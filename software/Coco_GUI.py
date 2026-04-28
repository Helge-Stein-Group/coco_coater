"""
coco_gui.py
===========
Graphical user interface for the CoCo (Combinatorial Coater) 
electrode coating platform.

CoCo is built on a modified Creality Ender-3 SE 3D printer and 
provides automated substrate mapping, blade handling, and three 
programmable coating modes (uniform, zonal, gradient).

Requirements
------------
    pip install customtkinter pyserial numpy openpyxl matplotlib

Usage
-----
    python coco_gui.py

For hardware setup, assembly instructions, and commissioning,
see the Supporting Information of the accompanying publication:

    Nuss et al., Digital Discovery, 2026
    DOI: [to be added]

License: MIT
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import serial
import time
import numpy as np
import os
import threading
import queue
from datetime import datetime
from openpyxl import Workbook, load_workbook
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


# ============================================================
# CONFIGURATION
# Adjust these settings for your system before first use.
# ============================================================

# Serial port of the Creality Ender-3 SE
# Windows: "COM3" (check Device Manager)
# Linux:   "/dev/ttyUSB0"
# macOS:   "/dev/tty.usbserial-*"
SERIAL_PORT = "COM3"
BAUD_RATE   = 115200

# Output directory for mapping and thickness data
BASE_DIR    = os.path.join(os.path.expanduser("~"), 
                           "Desktop", "Coating_Results")


# ============================================================
# COATING FIELD DEFINITIONS
# Each field is defined as (x_start, x_end, y_start, y_end)
# in printer coordinates [mm].
# ============================================================
Field_Areas = {
    1: (155, 220, 134, 218),
    2: (155, 220,  32, 116),
    3: ( 48, 138,  32, 116),
    4: ( 48, 138, 134, 218),
}

# Zone boundaries for zonal coating mode
# Each entry: (y_start, y_split, y_end)
# Zone 1: y >= y_split, Zone 2: y < y_split
Zone_Boundaries = {
    1: (134, 176, 218),
    2: ( 32,  74, 116),
    3: ( 32,  74, 116),
    4: (134, 176, 218),
}

# Offset between printer coordinate origin and 
# coating field coordinate system [mm]
X_OFFSET = 24.25
Y_OFFSET = 15.0

# Measurement grid spacing [mm]
STEP_MM = 7

# Derived paths -- do not change
MAPPING_DIR = os.path.join(BASE_DIR, "mappings")
RESULTS_DIR = os.path.join(BASE_DIR, "dry_results")
EXCEL_PATH  = os.path.join(RESULTS_DIR, "Coating_Results.xlsx")

os.makedirs(BASE_DIR,    exist_ok=True)
os.makedirs(MAPPING_DIR, exist_ok=True)
os.makedirs(RESULTS_DIR, exist_ok=True)

# Global serial connection object
ser = None


# ============================================================
# SERIAL COMMUNICATION
# ============================================================

def commando(ser, command):
    """
    Send a G-code command and wait for 'ok' response.

    Parameters
    ----------
    ser : serial.Serial
        Active serial connection to the printer.
    command : str
        G-code command string including line terminator.

    Returns
    -------
    list of str
        All response lines received before 'ok'.
    """
    if ser is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")
    ser.write(str.encode(command))
    time.sleep(0.02)
    output_data = []
    while True:
        line = ser.readline()
        decoded_line = line.decode("utf-8").strip()
        if decoded_line:
            output_data.append(decoded_line)
        if "ok" in decoded_line.lower():
            break
    return output_data


def heatmap_commando(ser_obj, command):
    """
    Send a G-code command for bed levelling and wait for response.

    Parameters
    ----------
    ser_obj : serial.Serial
        Active serial connection to the printer.
    command : str
        G-code command string.

    Returns
    -------
    list of str
        All response lines received.
    """
    if ser_obj is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")
    ser_obj.write(command.encode())
    time.sleep(0.1)
    output_data = []
    while True:
        decoded_line = ser_obj.readline().decode("utf-8").strip()
        if decoded_line:
            output_data.append(decoded_line)
        if "ok" in decoded_line.lower():
            break
    return output_data


# ============================================================
# BED LEVEL HEATMAP (Quick Bed Check tab)
# ============================================================

def strip_heatmap_output(output_data, keyword):
    """Extract the bilinear levelling grid from Marlin output."""
    for i, line in enumerate(output_data):
        if keyword in line:
            return output_data[i + 1:i + 6]
    return []


def generate_heatmap_data(ser_obj):
    """
    Execute a Marlin bilinear bed-level scan and return the 
    height grid as a numpy array.

    Uses G29 with fixed boundaries covering the full 
    vacuum plate area (L15 R195 F15 B195).

    Returns
    -------
    np.ndarray, shape (4, 4)
        Z-height values [mm] at the 4x4 probing grid.
    """
    if ser_obj is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")
    heatmap_commando(ser_obj, "G28\r\n")
    output = heatmap_commando(
        ser_obj, "G29 L15 R195 F15 B195\r\n")
    heatmap_commando(ser_obj, "M420 S1\r\n")

    keyword = "Bilinear Leveling Grid"
    table_section = strip_heatmap_output(output, keyword)

    if len(table_section) < 2:
        raise RuntimeError(
            "No valid Bilinear Leveling Grid found in response.")

    table = []
    for line in table_section[1:]:
        values = line.split()[1:]
        row = [float(v) for v in values]
        table.append(row)

    return np.array(table, dtype=float)


def show_heatmap_in_gui(data, frame):
    """
    Render a bed-level heatmap in the GUI Quick Bed Check tab.

    Displays the full bed and all four coating quadrants with 
    peak-to-valley deviations in µm.

    Parameters
    ----------
    data : np.ndarray
        4x4 Z-height array from generate_heatmap_data().
    frame : ctk.CTkFrame
        GUI frame to render the plot into.
    """
    for widget in frame.winfo_children():
        widget.destroy()

    # Split into quadrants matching the four coating fields
    q1 = data[2:, 2:][::-1, :]   # Field 1: top right
    q2 = data[:2, 2:][::-1, :]   # Field 2: bottom right
    q3 = data[:2, :2][::-1, :]   # Field 3: bottom left
    q4 = data[2:, :2][::-1, :]   # Field 4: top left

    delta = int((data.max() - data.min()) * 1000)
    d1 = int((q1.max() - q1.min()) * 1000)
    d2 = int((q2.max() - q2.min()) * 1000)
    d3 = int((q3.max() - q3.min()) * 1000)
    d4 = int((q4.max() - q4.min()) * 1000)

    vmin, vmax = data.min(), data.max()

    fig  = Figure(figsize=(10, 8), dpi=100)
    axs  = fig.subplots(3, 2)

    im = axs[0, 0].imshow(
        data, cmap="viridis", origin="lower", 
        vmin=vmin, vmax=vmax)
    fig.colorbar(im, ax=axs[0, 0], label="Height [mm]")
    axs[0, 0].set_title(f"Whole Bed – Δ={delta} µm")
    axs[0, 0].set_xlabel("X")
    axs[0, 0].set_ylabel("Y")
    axs[0, 1].axis("off")

    for ax, q, d, label in [
        (axs[1, 0], q4, d4, "Field 4"),
        (axs[1, 1], q1, d1, "Field 1"),
        (axs[2, 0], q3, d3, "Field 3"),
        (axs[2, 1], q2, d2, "Field 2"),
    ]:
        im_q = ax.imshow(
            q, cmap="viridis", origin="lower",
            vmin=vmin, vmax=vmax)
        fig.colorbar(im_q, ax=ax)
        ax.set_title(f"{label} – Δ={d} µm")

    fig.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)


# ============================================================
# GRID AND MAPPING
# ============================================================

def generate_grid(x_start, x_end, y_start, y_end):
    """
    Generate the measurement grid for a coating field.

    Returns x and y positions at STEP_MM spacing. 
    Y values are returned in reverse order (high to low) 
    to match the coating direction.
    """
    x_vals = list(np.arange(x_start, x_end + 0.0001, STEP_MM))
    y_vals = list(np.arange(y_start, y_end + 0.0001, STEP_MM))
    return x_vals, list(reversed(y_vals))


def Map_CoatingArea(field_id, root):
    """
    Map the substrate surface of a coating field point-by-point.

    Uses G30 single-point probing at each grid position. 
    Results are returned as a 2D numpy array of Z-heights [mm].

    Parameters
    ----------
    field_id : int
        Coating field number (1--4).
    root : CoatingApp
        GUI instance for progress logging.

    Returns
    -------
    z_data : np.ndarray, shape (n_rows, n_cols)
        Measured Z-heights [mm].
    x_vals : list of float
        X-positions of measurement grid [mm].
    y_vals : list of float
        Y-positions of measurement grid [mm].
    """
    if ser is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")

    root.log_message(
        f"[INFO] Starting mapping for Field {field_id}")
    commando(ser, "G28\r\n")
    commando(ser, "M420 S0\r\n")  # disable mesh compensation

    x_start, x_end, y_start, y_end = Field_Areas[field_id]
    x_vals, y_vals = generate_grid(
        x_start, x_end, y_start, y_end)
    z_data    = []
    total_rows = len(y_vals)

    for row_idx, y in enumerate(y_vals, start=1):
        row = []
        for x in x_vals:
            response = commando(
                ser, 
                f"G30 X{x - X_OFFSET} Y{y - Y_OFFSET}\r\n")
            z_val = None
            for line in response:
                if "Bed" in line and "Z:" in line:
                    try:
                        z_val = float(
                            line.split("Z:")[1].strip())
                        break
                    except ValueError:
                        continue
            row.append(z_val if z_val is not None else 0.0)
        z_data.append(row)
        root.log_message(
            f"[INFO] Field {field_id}: "
            f"row {row_idx}/{total_rows} at Y={y:.2f}")

    root.log_message(
        f"[INFO] Mapping complete for Field {field_id}")
    return np.array(z_data), x_vals, y_vals


def save_blank_mapping(field_id, z_array, x_vals, y_vals):
    """Save pre-coating substrate mapping to disk."""
    path = os.path.join(
        MAPPING_DIR, f"field{field_id}_blank_latest.npz")
    np.savez(path, z_array=z_array, 
             x_vals=x_vals, y_vals=y_vals)


def load_blank_mapping(field_id):
    """
    Load the most recent pre-coating mapping for a field.

    Returns None if no mapping exists.
    """
    path = os.path.join(
        MAPPING_DIR, f"field{field_id}_blank_latest.npz")
    if not os.path.exists(path):
        return None
    data = np.load(path)
    return data["z_array"], data["x_vals"], data["y_vals"]


def map_fields(field_ids, root):
    """Map multiple coating fields sequentially."""
    if ser is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")
    for field_id in field_ids:
        z_blank, x_vals, y_vals = Map_CoatingArea(
            field_id, root)
        save_blank_mapping(field_id, z_blank, x_vals, y_vals)
        root.log_message(
            f"[INFO] Blank mapping saved for Field {field_id}")
    return True


# ============================================================
# COATING CONFIGURATION
# ============================================================

def save_coating_config(field_id, mode, params):
    """Save coating parameters alongside the blank mapping."""
    path = os.path.join(
        MAPPING_DIR, 
        f"field{field_id}_coating_latest.npz")
    data = {
        "mode":      np.array([mode]),
        "timestamp": np.array([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    }
    for key, value in params.items():
        if isinstance(value, (list, np.ndarray)):
            data[key] = np.array(value, dtype=float)
        else:
            data[key] = np.array([value], dtype=float)
    np.savez(path, **data)


def load_coating_config(field_id):
    """Load the most recent coating configuration for a field."""
    path = os.path.join(
        MAPPING_DIR, 
        f"field{field_id}_coating_latest.npz")
    if not os.path.exists(path):
        return None
    raw    = np.load(path, allow_pickle=True)
    config = {}
    for key in raw.files:
        value = raw[key]
        if key == "mode":
            config[key] = str(value[0])
        elif value.ndim == 0:
            config[key] = value.item()
        elif len(value) == 1:
            try:
                config[key] = float(value[0])
            except Exception:
                config[key] = value[0]
        else:
            config[key] = value.tolist()
    return config


# ============================================================
# COATING MODE HELPERS
# ============================================================

def get_zone_index(y, field_id):
    """
    Return zone index (1 or 2) for a given Y-position.

    Zone 1 corresponds to the upper half of the field 
    (y >= y_split), Zone 2 to the lower half.
    """
    _, y_split, _ = Zone_Boundaries[field_id]
    return 1 if y >= y_split else 2


def generate_gradient(height_start, height_end, num_points):
    """
    Generate a linearly spaced list of gap heights for 
    gradient coating mode.
    """
    if num_points < 2:
        return [float(height_start)]
    delta = (float(height_end) - float(height_start)) \
            / (num_points - 1)
    return [float(height_start) + i * delta 
            for i in range(num_points)]


def get_row_setpoints(field_id, y_vals, config):
    """
    Return per-row gap heights and coating speeds based on 
    the coating mode stored in config.

    Returns
    -------
    heights : list of float
        Target gap height [µm] for each Y-row.
    speeds : list of float
        Coating speed [mm/s] for each Y-row.
    """
    mode = config["mode"]

    if mode == "UNIFORM":
        return (
            [float(config["height_um"])] * len(y_vals),
            [float(config["speed_mm_s"])] * len(y_vals)
        )
    elif mode == "ZONE":
        zone_heights = config["zone_heights_um"]
        zone_speeds  = config["zone_speeds_mm_s"]
        return (
            [float(zone_heights[
                get_zone_index(y, field_id) - 1]) 
             for y in y_vals],
            [float(zone_speeds[
                get_zone_index(y, field_id) - 1]) 
             for y in y_vals]
        )
    elif mode == "GRADIENT":
        heights = generate_gradient(
            float(config["height_start_um"]),
            float(config["height_end_um"]),
            len(y_vals)
        )
        return heights, \
               [float(config["speed_mm_s"])] * len(y_vals)
    else:
        raise ValueError(f"Unknown coating mode: {mode}")


# ============================================================
# BLADE HANDLING
# ============================================================

def pick_up_blade(blade_number):
    """
    Automated blade pickup from the storage rack.

    Moves to the slot corresponding to blade_number, lifts 
    the blade onto the magnetic holder, and positions the 
    holder above the coating field.

    Parameters
    ----------
    blade_number : int
        Slot number (1--4). Each slot has a different Z-height
        to account for blade stacking in the rack.
    """
    if ser is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")

    # Z-height for each blade slot [mm]
    # 30 mm offset per slot accounts for blade stacking
    z_positions = {1: 80, 2: 110, 3: 140, 4: 170}
    z = z_positions.get(blade_number, 80)

    commands = [
        f"G0 X135 Z{z}\r\n",       # move to rack X position
        "G0 Y12\r\n",              # move into rack
        f"G0 Z{z + 2}\r\n",        # lift blade off peg
        "G0 Y80\r\n",              # retract from rack
        "G0 X110 Y233 F800\r\n",   # move to coating position
        "M211 S0\r\n",             # disable soft endstops
        "G0 Z5 F800\r\n",          # lower toward substrate
        "G0 Z-0.4 F300\r\n",       # final approach
        "M400\r\n",                # wait for moves to finish
        "G4 P1000\r\n",            # dwell 1 s
        "G0 Z10 F800\r\n",         # lift blade
        "M211 S1\r\n",             # re-enable soft endstops
    ]
    for cmd in commands:
        commando(ser, cmd)


def drop_blade():
    """
    Eject the used blade into the collection tray.

    Moves the blade holder to the ejection position above 
    the tray and releases the blade by moving down.
    """
    if ser is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")

    commands = [
        "G0 Z30 F800\r\n",   # raise blade
        "G0 X135\r\n",       # move to ejection X
        "G0 Y15\r\n",        # move to ejection Y
        "G0 Z55\r\n",        # lower into tray
        "G0 Y60\r\n",        # push blade off magnets
    ]
    for cmd in commands:
        commando(ser, cmd)


# ============================================================
# COATING FUNCTIONS
# ============================================================

def coating_uniform(field_id, speed, height, root):
    """
    Execute a uniform coating run.

    Applies a constant gap height and coating speed across 
    the entire field with row-wise blade height referencing 
    derived from the pre-coating substrate mapping.

    Parameters
    ----------
    field_id : int
        Coating field (1--4).
    speed : float
        Coating speed [mm/s].
    height : float
        Nominal gap height [µm].
    root : CoatingApp
        GUI instance for logging and dialogs.
    """
    if ser is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")

    loaded = load_blank_mapping(field_id)
    if loaded is None:
        raise RuntimeError(
            f"No blank mapping found for Field {field_id}. "
            "Run mapping first.")

    save_coating_config(field_id, "UNIFORM", 
                        {"speed_mm_s": speed, 
                         "height_um":  height})
    z_blank, x_vals, y_vals = loaded
    z_max_per_row = np.max(z_blank, axis=1)

    x_probe = 172 if field_id in [1, 2] else 60
    y_probe = 231 if field_id in [1, 4] else 129
    feed    = float(speed) * 60.0  # mm/s to mm/min

    pick_up_blade(field_id)
    commando(ser, f"G0 X{x_probe} Y{y_probe}\r\n")
    commando(ser, "G0 Z0\r\n")
    commando(ser, "M400\r\n")
    commando(ser, "G4 P3000\r\n")  # wait 3 s for slurry

    root.show_info(
        "Confirmation",
        f"Ready to coat Field {field_id}.\n"
        "Apply slurry in front of the blade and click OK.")

    for idx, (y, z_row_base) in enumerate(
            zip(y_vals, z_max_per_row), start=1):
        # Row-wise blade referencing: blade positioned at
        # row maximum + nominal gap
        z_target = float(z_row_base) + float(height) / 1000.0
        root.log_message(
            f"[COATING] Field {field_id}, "
            f"Row {idx}: Y={y:.2f}, Z={z_target:.4f}")
        commando(ser,
            f"G1 X{x_probe:.3f} "
            f"Y{float(y) + 13:.3f} "
            f"Z{z_target:.4f} "
            f"F{feed:.1f}\r\n")

    commando(ser, "G0 Z10\r\n")
    drop_blade()
    return True


def coating_zones(field_id, speeds, heights, root):
    """
    Execute a zonal coating run with two independent zones.

    Zone 1 and Zone 2 receive independent gap heights and 
    speeds, applied from the same slurry batch in a single 
    run.

    Parameters
    ----------
    field_id : int
        Coating field (1--4).
    speeds : list of float
        Coating speeds [mm/s] for [Zone 1, Zone 2].
    heights : list of float
        Nominal gap heights [µm] for [Zone 1, Zone 2].
    root : CoatingApp
        GUI instance for logging and dialogs.
    """
    if ser is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")

    loaded = load_blank_mapping(field_id)
    if loaded is None:
        raise RuntimeError(
            f"No blank mapping found for Field {field_id}. "
            "Run mapping first.")

    save_coating_config(field_id, "ZONE", {
        "zone_speeds_mm_s":  [float(s) for s in speeds],
        "zone_heights_um":   [float(h) for h in heights]
    })

    z_blank, x_vals, y_vals = loaded
    z_max_per_row = np.max(z_blank, axis=1)
    x_probe = 172 if field_id in [1, 2] else 60
    y_probe = 231 if field_id in [1, 4] else 129

    pick_up_blade(field_id)
    commando(ser, f"G0 X{x_probe} Y{y_probe}\r\n")
    commando(ser, "G0 Z0\r\n")
    commando(ser, "M400\r\n")
    commando(ser, "G4 P3000\r\n")

    root.show_info(
        "Confirmation",
        f"Ready to coat Field {field_id}.\n"
        "Apply slurry in front of the blade and click OK.")

    for idx, y in enumerate(y_vals, start=1):
        zone_idx = get_zone_index(y, field_id) - 1
        speed    = float(speeds[zone_idx])
        height   = float(heights[zone_idx])
        feed     = speed * 60.0
        z_target = float(z_max_per_row[idx - 1]) \
                   + height / 1000.0
        root.log_message(
            f"[COATING] Field {field_id}, Row {idx}: "
            f"Zone {zone_idx + 1}, Y={y:.2f}, "
            f"Z={z_target:.4f}")
        commando(ser,
            f"G1 X{x_probe:.3f} "
            f"Y{float(y) + 13:.3f} "
            f"Z{z_target:.4f} "
            f"F{feed:.1f}\r\n")

    commando(ser, "G0 Z10\r\n")
    drop_blade()
    return True


def coating_gradient(field_id, speed, 
                     height_start, height_end, root):
    """
    Execute a gradient coating run.

    The gap height varies linearly from height_start to 
    height_end along the Y-direction. All other parameters 
    remain constant.

    Parameters
    ----------
    field_id : int
        Coating field (1--4).
    speed : float
        Coating speed [mm/s].
    height_start : float
        Gap height at the start of the coating [µm].
    height_end : float
        Gap height at the end of the coating [µm].
    root : CoatingApp
        GUI instance for logging and dialogs.
    """
    if ser is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")

    loaded = load_blank_mapping(field_id)
    if loaded is None:
        raise RuntimeError(
            f"No blank mapping found for Field {field_id}. "
            "Run mapping first.")

    save_coating_config(field_id, "GRADIENT", {
        "speed_mm_s":      float(speed),
        "height_start_um": float(height_start),
        "height_end_um":   float(height_end)
    })

    z_blank, x_vals, y_vals = loaded
    z_max_per_row = np.max(z_blank, axis=1)
    heights = generate_gradient(
        float(height_start), float(height_end), len(y_vals))
    x_probe = 172 if field_id in [1, 2] else 60
    y_probe = 231 if field_id in [1, 4] else 129
    feed    = float(speed) * 60.0

    pick_up_blade(field_id)
    commando(ser, f"G0 X{x_probe} Y{y_probe}\r\n")
    commando(ser, "G0 Z0\r\n")
    commando(ser, "M400\r\n")
    commando(ser, "G4 P3000\r\n")

    root.show_info(
        "Confirmation",
        f"Ready to coat Field {field_id}.\n"
        "Apply slurry in front of the blade and click OK.")

    for idx, (y, z_row_base, height) in enumerate(
            zip(y_vals, z_max_per_row, heights), start=1):
        z_target = float(z_row_base) + height / 1000.0
        root.log_message(
            f"[COATING] Field {field_id}, Row {idx}: "
            f"Gradient height={height:.1f} µm, "
            f"Y={y:.2f}, Z={z_target:.4f}")
        commando(ser,
            f"G1 X{x_probe:.3f} "
            f"Y{float(y) + 13:.3f} "
            f"Z{z_target:.4f} "
            f"F{feed:.1f}\r\n")

    commando(ser, "G0 Z10\r\n")
    drop_blade()
    return True


# ============================================================
# DRY THICKNESS ANALYSIS
# ============================================================

def measure_dry_thickness(field_sample_pairs, root):
    """
    Map coated fields and compute point-wise dry thickness.

    For each selected field, the coated surface is re-mapped 
    at the same grid positions used for the blank mapping. 
    Dry thickness is computed as:

        t_dry(x, y) = z_dry(x, y) - z_blank(x, y)

    Results are appended as a new sheet to Coating_Results.xlsx.

    Parameters
    ----------
    field_sample_pairs : list of (int, str)
        List of (field_id, sample_name) tuples.
    root : CoatingApp
        GUI instance for logging and dialogs.
    """
    if ser is None:
        raise RuntimeError(
            "No serial connection. "
            "Please connect to the printer first.")

    for field_id, sample_name in field_sample_pairs:
        root.log_message(
            f"[ANALYSIS] Starting Field {field_id} "
            f"({sample_name})")

        loaded_blank = load_blank_mapping(field_id)
        if loaded_blank is None:
            root.log_message(
                f"[ANALYSIS] No blank mapping for "
                f"Field {field_id} -- skipping")
            continue

        config = load_coating_config(field_id)
        if config is None:
            root.log_message(
                f"[ANALYSIS] No coating config for "
                f"Field {field_id} -- skipping")
            continue

        z_dry, x_dry, y_dry = Map_CoatingArea(
            field_id, root)
        z_blank, x_blank, y_blank = loaded_blank
        row_heights, row_speeds = get_row_setpoints(
            field_id, y_blank, config)

        point_rows = []
        for row_idx, y in enumerate(y_blank):
            set_height   = float(row_heights[row_idx])
            set_speed    = float(row_speeds[row_idx])
            z_row_base   = float(np.max(z_blank[row_idx]))
            z_target_row = z_row_base + set_height / 1000.0

            for col_idx, x in enumerate(x_blank):
                z_blank_val   = float(z_blank[row_idx, col_idx])
                z_dry_val     = float(z_dry[row_idx, col_idx])
                thickness_mm  = z_dry_val - z_blank_val

                point_rows.append([
                    sample_name,
                    config["mode"],
                    float(x),
                    float(y),
                    round(z_blank_val,              4),
                    round(z_dry_val,                4),
                    round(z_target_row - z_blank_val, 4),
                    round(thickness_mm,             4),
                    round(thickness_mm * 1000,      3),
                    get_zone_index(float(y), field_id),
                    round(set_speed,   3),
                    round(set_height,  2),
                    "stats"
                ])

        header = [
            "sample_name", "coating_mode",
            "x_mm", "y_mm",
            "z_blank_mm", "z_dry_mm",
            "z_actual_gap_mm",
            "thickness_mm", "thickness_um",
            "zone_index",
            "set_speed_mm_s", "set_height_um",
            "stats_included"
        ]

        mode_letter = str(config["mode"])[0].upper()
        sheet_name  = (
            f"F{field_id}_{mode_letter}_"
            f"{sample_name}_"
            f"{datetime.now().strftime('%d%m%Y_%H%M%S')}"
        )[:31]

        if os.path.exists(EXCEL_PATH):
            wb = load_workbook(EXCEL_PATH)
        else:
            wb = Workbook()
            del wb[wb.sheetnames[0]]

        sheet = wb.create_sheet(title=sheet_name)
        for col, val in enumerate(header, 1):
            sheet.cell(row=1, column=col, value=val)
        for r_idx, row in enumerate(point_rows, 2):
            for c_idx, val in enumerate(row, 1):
                sheet.cell(row=r_idx, column=c_idx, value=val)

        wb.save(EXCEL_PATH)
        root.log_message(
            f"[ANALYSIS] Field {field_id} complete. "
            f"Saved to {EXCEL_PATH}")

    root.show_info(
        "Done", 
        "Analysis completed for all selected fields.")
    return True


# ============================================================
# GUI
# ============================================================

class CoatingApp(ctk.CTk):
    """Main application window for the CoCo coating system."""

    def __init__(self):
        super().__init__()
        self.title("CoCo – Combinatorial Coater")
        self.geometry("1200x900")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title_font  = ctk.CTkFont(size=16, weight="bold")
        self.label_font  = ctk.CTkFont(size=12)
        self.button_font = ctk.CTkFont(size=12, weight="bold")

        self.log_queue = queue.Queue()
        self.busy      = False

        self.create_widgets()
        self.after(100, self.process_log_queue)

    def create_widgets(self):
        # Connection bar
        self.conn_frame = ctk.CTkFrame(self)
        self.conn_frame.pack(pady=10, padx=10, fill="x")

        self.conn_button = ctk.CTkButton(
            self.conn_frame,
            text="Connect to Device",
            command=self.toggle_connection,
            fg_color="#3498db", hover_color="#2980b9",
            font=self.button_font)
        self.conn_button.pack(side="left", padx=5, pady=5)

        self.conn_status = ctk.CTkLabel(
            self.conn_frame,
            text="Status: Not Connected",
            text_color="orange",
            font=self.label_font)
        self.conn_status.pack(side="left", padx=5, pady=5)

        # Tab view
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.pack(
            pady=10, padx=10, fill="both", expand=True)

        for tab in ["Quick Bed Check", "Mapping", 
                    "Coating", "Analysis"]:
            self.tab_view.add(tab)

        self.setup_heatmap_tab()
        self.setup_mapping_tab()
        self.setup_coating_tab()
        self.setup_analysis_tab()

        # Log console
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.pack(pady=10, padx=10, fill="x")

        ctk.CTkLabel(
            self.log_frame, text="Operation Log:",
            font=self.title_font).pack(pady=(0, 5))

        self.log_text = tk.Text(
            self.log_frame, height=10, width=100,
            state="disabled", wrap="word",
            font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)

        scrollbar = ctk.CTkScrollbar(
            self.log_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

    def setup_heatmap_tab(self):
        frame = ctk.CTkFrame(
            self.tab_view.tab("Quick Bed Check"))
        frame.pack(pady=10, padx=10, fill="both", expand=True)

        ctk.CTkLabel(
            frame, text="Quick Bed Check",
            font=self.title_font).pack(pady=(0, 10))

        self.heatmap_frame = ctk.CTkFrame(frame)
        self.heatmap_frame.pack(
            fill="both", expand=True, pady=10)

        self.heatmap_button = ctk.CTkButton(
            frame, text="Generate Heatmap",
            command=self.generate_heatmap,
            fg_color="#27ae60", hover_color="#219653",
            font=self.button_font)
        self.heatmap_button.pack(pady=10)

    def setup_mapping_tab(self):
        frame = ctk.CTkFrame(self.tab_view.tab("Mapping"))
        frame.pack(pady=10, padx=10, fill="both", expand=True)

        ctk.CTkLabel(
            frame, text="Field Mapping",
            font=self.title_font).pack(pady=(0, 10))

        sel = ctk.CTkFrame(frame)
        sel.pack(fill="x", pady=5)
        ctk.CTkLabel(
            sel, text="Select Fields:",
            font=self.label_font).pack(side="left", padx=5)

        self.field_vars = {f: tk.IntVar() for f in [1,2,3,4]}
        for f in [1, 2, 3, 4]:
            ctk.CTkCheckBox(
                sel, text=f"Field {f}",
                variable=self.field_vars[f],
                font=self.label_font
            ).pack(side="left", padx=10)

        self.mapping_button = ctk.CTkButton(
            frame, text="Start Mapping",
            command=self.run_mapping,
            fg_color="#2c3e50", hover_color="#1a252f",
            font=self.button_font)
        self.mapping_button.pack(pady=20)

    def setup_coating_tab(self):
        frame = ctk.CTkFrame(self.tab_view.tab("Coating"))
        frame.pack(pady=10, padx=10, fill="both", expand=True)

        ctk.CTkLabel(
            frame, text="Coating Process",
            font=self.title_font).pack(pady=(0, 10))

        # Field selection
        ff = ctk.CTkFrame(frame)
        ff.pack(fill="x", pady=5)
        ctk.CTkLabel(
            ff, text="Select Field:",
            font=self.label_font).pack(side="left", padx=5)
        self.coat_field_var = ctk.StringVar(value="1")
        ctk.CTkOptionMenu(
            ff, variable=self.coat_field_var,
            values=["1", "2", "3", "4"],
            font=self.label_font).pack(side="left", padx=5)

        # Mode selection
        mf = ctk.CTkFrame(frame)
        mf.pack(fill="x", pady=10)
        ctk.CTkLabel(
            mf, text="Coating Mode:",
            font=self.label_font).pack(side="left", padx=5)
        self.coat_mode_var = ctk.StringVar(value="UNIFORM")
        ctk.CTkOptionMenu(
            mf, variable=self.coat_mode_var,
            values=["UNIFORM", "ZONE", "GRADIENT"],
            command=self.update_coat_params,
            font=self.label_font).pack(side="left", padx=5)

        # Parameter frames
        self.coat_params_frame = ctk.CTkFrame(frame)
        self.coat_params_frame.pack(fill="x", pady=10)

        # UNIFORM
        self.uniform_frame = ctk.CTkFrame(
            self.coat_params_frame)
        ctk.CTkLabel(
            self.uniform_frame, text="Speed (mm/s):",
            font=self.label_font).pack(side="left", padx=5)
        self.uniform_speed = ctk.CTkEntry(
            self.uniform_frame, placeholder_text="e.g. 5",
            font=self.label_font)
        self.uniform_speed.pack(side="left", padx=5)
        ctk.CTkLabel(
            self.uniform_frame, text="Height (µm):",
            font=self.label_font).pack(side="left", padx=5)
        self.uniform_height = ctk.CTkEntry(
            self.uniform_frame, placeholder_text="e.g. 200",
            font=self.label_font)
        self.uniform_height.pack(side="left", padx=5)

        # ZONE
        self.zones_frame = ctk.CTkFrame(
            self.coat_params_frame)
        for i, zone in enumerate(["Zone 1", "Zone 2"], 1):
            zr = ctk.CTkFrame(self.zones_frame)
            zr.pack(fill="x", pady=2)
            ctk.CTkLabel(
                zr, text=f"{zone} Speed (mm/s):",
                font=self.label_font).pack(
                    side="left", padx=5)
            setattr(self, f"zone{i}_speed",
                    ctk.CTkEntry(zr, 
                        placeholder_text="e.g. 5",
                        font=self.label_font))
            getattr(self, f"zone{i}_speed").pack(
                side="left", padx=5)
            ctk.CTkLabel(
                zr, text=f"{zone} Height (µm):",
                font=self.label_font).pack(
                    side="left", padx=5)
            setattr(self, f"zone{i}_height",
                    ctk.CTkEntry(zr, 
                        placeholder_text="e.g. 200",
                        font=self.label_font))
            getattr(self, f"zone{i}_height").pack(
                side="left", padx=5)

        # GRADIENT
        self.gradient_frame = ctk.CTkFrame(
            self.coat_params_frame)
        gr1 = ctk.CTkFrame(self.gradient_frame)
        gr1.pack(fill="x", pady=2)
        ctk.CTkLabel(
            gr1, text="Speed (mm/s):",
            font=self.label_font).pack(side="left", padx=5)
        self.gradient_speed = ctk.CTkEntry(
            gr1, placeholder_text="e.g. 5",
            font=self.label_font)
        self.gradient_speed.pack(side="left", padx=5)
        ctk.CTkLabel(
            gr1, text="Start Height (µm):",
            font=self.label_font).pack(side="left", padx=5)
        self.gradient_start = ctk.CTkEntry(
            gr1, placeholder_text="e.g. 100",
            font=self.label_font)
        self.gradient_start.pack(side="left", padx=5)
        gr2 = ctk.CTkFrame(self.gradient_frame)
        gr2.pack(fill="x", pady=2)
        ctk.CTkLabel(
            gr2, text="End Height (µm):",
            font=self.label_font).pack(side="left", padx=5)
        self.gradient_end = ctk.CTkEntry(
            gr2, placeholder_text="e.g. 300",
            font=self.label_font)
        self.gradient_end.pack(side="left", padx=5)

        self.coating_button = ctk.CTkButton(
            frame, text="Start Coating Process",
            command=self.run_coating,
            fg_color="#27ae60", hover_color="#219653",
            font=self.button_font)
        self.coating_button.pack(pady=20)
        self.update_coat_params()

    def setup_analysis_tab(self):
        frame = ctk.CTkFrame(self.tab_view.tab("Analysis"))
        frame.pack(pady=10, padx=10, fill="both", expand=True)

        ctk.CTkLabel(
            frame, text="Dry Thickness Analysis",
            font=self.title_font).pack(pady=(0, 10))

        self.analysis_entries = {}
        for f in [1, 2, 3, 4]:
            ff = ctk.CTkFrame(frame)
            ff.pack(fill="x", pady=5)
            var = tk.IntVar()
            ctk.CTkCheckBox(
                ff, text=f"Field {f}",
                variable=var,
                font=self.label_font).pack(
                    side="left", padx=5)
            entry = ctk.CTkEntry(
                ff,
                placeholder_text=f"Sample name for Field {f}",
                font=self.label_font, width=200)
            entry.pack(side="left", padx=10)
            self.analysis_entries[f] = (var, entry)

        self.analysis_button = ctk.CTkButton(
            frame, text="Start Analysis",
            command=self.run_analysis,
            fg_color="#e74c3c", hover_color="#c0392b",
            font=self.button_font)
        self.analysis_button.pack(pady=20)

    def update_coat_params(self, *args):
        """Show/hide parameter inputs based on selected mode."""
        for f in [self.uniform_frame, 
                  self.zones_frame, 
                  self.gradient_frame]:
            f.pack_forget()
        mode = self.coat_mode_var.get()
        if mode == "UNIFORM":
            self.uniform_frame.pack()
        elif mode == "ZONE":
            self.zones_frame.pack()
        elif mode == "GRADIENT":
            self.gradient_frame.pack()

    def log_message(self, message):
        """Append a timestamped message to the log console."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"{timestamp} - {message}")

    def process_log_queue(self):
        """Drain the log queue and update the text widget."""
        messages = []
        while True:
            try:
                messages.append(self.log_queue.get_nowait())
            except queue.Empty:
                break
        if messages:
            self.log_text.config(state="normal")
            self.log_text.insert(
                "end", "\n".join(messages) + "\n")
            total = int(
                self.log_text.index("end-1c").split(".")[0])
            if total > 2000:
                self.log_text.delete(
                    "1.0", f"{total - 2000}.0")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(100, self.process_log_queue)

    def show_info(self, title, text):
        """Thread-safe info dialog."""
        event = threading.Event()
        def _show():
            messagebox.showinfo(title, text)
            event.set()
        self.after(0, _show)
        event.wait()

    def show_error(self, title, text):
        """Thread-safe error dialog."""
        self.after(0, lambda: messagebox.showerror(
            title, text))

    def set_busy(self, busy, source_text=None):
        """Disable/enable all action buttons during operations."""
        self.busy  = busy
        state = "disabled" if busy else "normal"
        for btn in [self.conn_button, self.heatmap_button,
                    self.mapping_button, self.coating_button,
                    self.analysis_button]:
            btn.configure(state=state)
        if busy and source_text:
            self.log_message(source_text)

    def start_background_task(self, worker, 
                               start_message=None):
        """Run a worker function in a background thread."""
        if self.busy:
            self.show_error(
                "Busy", 
                "Another operation is currently running.")
            return
        self.set_busy(True, start_message)
        def wrapped():
            try:
                worker()
            except Exception as e:
                self.log_message(f"Error: {e}")
                self.show_error("Error", str(e))
            finally:
                self.after(0, lambda: self.set_busy(False))
        threading.Thread(target=wrapped, daemon=True).start()

    def toggle_connection(self):
        """Connect to or disconnect from the printer."""
        global ser
        if self.busy:
            self.show_error(
                "Busy",
                "Cannot connect while an operation is running.")
            return
        if ser is None:
            try:
                ser = serial.Serial(
                    SERIAL_PORT, BAUD_RATE, timeout=1)
                time.sleep(2)
                ser.reset_input_buffer()
                ser.reset_output_buffer()
                self.conn_status.configure(
                    text="Status: Connected",
                    text_color="green")
                self.conn_button.configure(
                    text="Disconnect")
                self.log_message(
                    f"Connected on {SERIAL_PORT}")
            except Exception as e:
                ser = None
                self.log_message(f"Connection failed: {e}")
                self.show_error("Error", str(e))
        else:
            try:
                ser.close()
            except Exception:
                pass
            ser = None
            self.conn_status.configure(
                text="Status: Not Connected",
                text_color="orange")
            self.conn_button.configure(
                text="Connect to Device")
            self.log_message("Disconnected")

    def generate_heatmap(self):
        def worker():
            data = generate_heatmap_data(ser)
            self.after(0, lambda: show_heatmap_in_gui(
                data, self.heatmap_frame))
            self.log_message("Heatmap generated.")
        self.start_background_task(
            worker, "Generating bed heatmap...")

    def run_mapping(self):
        fields = [f for f in [1,2,3,4] 
                  if self.field_vars[f].get()]
        if not fields:
            self.show_error(
                "Error", "Select at least one field.")
            return
        def worker():
            map_fields(fields, self)
            self.show_info(
                "Done", 
                "Mapping completed for all selected fields.")
        self.start_background_task(
            worker, f"Mapping fields: {fields}")

    def run_coating(self):
        try:
            field_id = int(self.coat_field_var.get())
            mode     = self.coat_mode_var.get()

            if mode == "UNIFORM":
                speed  = float(self.uniform_speed.get())
                height = float(self.uniform_height.get())
                def worker():
                    coating_uniform(
                        field_id, speed, height, self)
                    self.log_message(
                        f"Coating complete: Field {field_id}")

            elif mode == "ZONE":
                speeds  = [float(getattr(
                    self, f"zone{i}_speed").get()) 
                           for i in [1, 2]]
                heights = [float(getattr(
                    self, f"zone{i}_height").get()) 
                           for i in [1, 2]]
                def worker():
                    coating_zones(
                        field_id, speeds, heights, self)
                    self.log_message(
                        f"Coating complete: Field {field_id}")

            elif mode == "GRADIENT":
                speed   = float(self.gradient_speed.get())
                h_start = float(self.gradient_start.get())
                h_end   = float(self.gradient_end.get())
                def worker():
                    coating_gradient(
                        field_id, speed, 
                        h_start, h_end, self)
                    self.log_message(
                        f"Coating complete: Field {field_id}")
            else:
                raise ValueError(f"Unknown mode: {mode}")

            self.start_background_task(
                worker,
                f"Starting {mode} coating: Field {field_id}")

        except Exception as e:
            self.log_message(f"Error: {e}")
            self.show_error("Error", str(e))

    def run_analysis(self):
        try:
            pairs = []
            for f, (var, entry) in \
                    self.analysis_entries.items():
                if var.get():
                    name = entry.get().strip() or \
                        f"Sample_F{f}_" \
                        f"{datetime.now().strftime('%Y%m%d')}"
                    pairs.append((f, name))
            if not pairs:
                self.show_error(
                    "Error", "Select at least one field.")
                return
            def worker():
                measure_dry_thickness(pairs, self)
                self.log_message(
                    f"Analysis complete. "
                    f"Saved to {EXCEL_PATH}")
            self.start_background_task(
                worker,
                f"Analysing fields: "
                f"{[f for f, _ in pairs]}")
        except Exception as e:
            self.log_message(f"Error: {e}")
            self.show_error("Error", str(e))


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    app = CoatingApp()
    app.mainloop()
