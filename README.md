# CoCo – Combinatorial Coater

CoCo is an open-source automated electrode coating platform built on a modified Creality Ender-3 SE 3D printer. It integrates pre-coating substrate surface mapping, automated blade handling, three programmable deposition modes, and post-coating dry thickness mapping into a single workflow operated through a Python-based GUI.

**Total hardware cost: ~€200**

> This repository accompanies the publication:
> *CoCo, a Low-Cost Open Electrode Coating 
> Platform with Integrated Process 
> Characterisation*  
> L. Nuss, F. Ganster, H. S. Stein  

---

## Features

- Automated blade pickup and ejection (up to 4 blades per session)
- Pre-coating substrate topography mapping on a 7 mm grid
- Three deposition modes:
  - **Uniform** – constant gap height and speed
  - **Zonal** – two independent parameter combinations from a single slurry batch
  - **Gradient** – linearly varying gap height along the coating direction
- Post-coating dry thickness mapping at identical grid positions
- Automatic export of structured process records (CSV/Excel) for every experiment
- Python-based GUI – no programming knowledge required

---

## Hardware Requirements

- Creality Ender-3 SE (modified)
- 3D printed parts (PLA)
- Standard laboratory compressed air supply (4–6 bar) for Venturi vacuum
- PC with USB port (Windows, Linux, or macOS)

See the Supporting Information of the paper for the complete bill of materials and assembly instructions.
 ⚠️ **Important:** The Creality Ender-3 **SE** is required. The Ender-3 KE looks similar but does not have a serial communication interface and is incompatible for CoCo.

---

## Software Requirements

Python 3.8 or later is required. 
Install all dependencies with:

```bash
pip install customtkinter pyserial numpy 
    openpyxl matplotlib
```

---

## Installation

```bash
git clone https://github.com/[your-username]/coco-coater.git
cd coco-coater
pip install customtkinter pyserial numpy openpyxl matplotlib
```

---

## Usage

### 1. Configure serial port

Open `coco_gui.py` and set the correct serial port for your system:

```python
SERIAL_PORT = "COM3"       # Windows
# SERIAL_PORT = "/dev/ttyUSB0"  # Linux
# SERIAL_PORT = "/dev/tty.usbserial-*"  # macOS
```

### 2. Launch the GUI

```bash
python coco_gui.py
```

### 3. Workflow

The GUI guides you through four tabs in sequence:

| Tab | Function |
|-----|----------|
| **Quick Bed Check** | Coarse bed-level scan and tilt heatmap |
| **Mapping** | Pre-coating substrate surface mapping |
| **Coating** | Select mode and parameters, start coating |
| **Analysis** | Post-coating dry thickness mapping and export |

### 4. Output data

All results are saved automatically to 
`~/Desktop/Coating_Results/`:
Coating_Results/
  mappings/ #raw Z-height arrays (.npz)
  dry_results/ #Coating_Results.xlsx

Each row in the Excel file corresponds to one measurement point and contains:
x/y position, pre- and post-coating Z-heights, effective gap, dry thickness, zone index, set speed, and set gap height

---

## System Commissioning

Before first use:

1. **Bed levelling** – use the *Quick Bed Check* tab to verify peak-to-valley tilt < 200 µm
2. **Z-offset calibration** – via the printer display (see Supporting Information)
3. **Functional verification** – run the dry-run check described in the Supporting Information

---

## Repository Structure
coco-coater/
├── README.md
├── LICENSE              # MIT (software)
├── software/
│   └── coco_gui.py
├── hardware/
│   ├── step/            # STEP files (editable)
│   ├── stl/             # STL files (print-ready)
│   └── fusion360/       # F3D files
└── data/
└── example_result.xlsx

---

## Licence

All software and hardware design files are released under the MIT Licence.

---

## Citation

If you use CoCo in your research, please cite:

---

## Contact

Leah Nuss  
Technical University of Munich  
leah.nuss@tum.de
