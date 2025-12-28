# MatTek
A.D.A.M - Avichay &amp; Daniel Analysis Microplate reader - Simplified GUI for bacterial growth analysis
# A.D.A.M - Avichay & Daniel Analysis Microplate reader

## Overview
A focused MATLAB GUI tool for microplate reader analysis. Specifically designed for bacterial growth analysis with OD and Luciferase measurements.

## Features
- **Step 1: Data Preparation** - Import and format Excel data from plate readers
- **Step 2: Raw OD/LUC Analysis** - Analyze optical density and luciferase readings with growth curves
- **Step 3: Luciferase Analysis** - Detailed luciferase enzyme activity analysis
- **Step 4: Lum vs. OD** - 2D comparative analysis between luminescence and optical density
- **Step 5: 3D Lum vs. OD vs. Time** - 3D visualization of all parameters

## How to Use
Run the main function in MATLAB:
```matlab
ADAM_main
```

A GUI window will open with 5 analysis steps. Follow them sequentially or jump to any step.

## Input Requirements
- Excel files (.xlsx) from microplate reader instruments
- Expected data: Well ID, Time points, OD600 measurements, Luciferase readings

## Output
- Growth curves and analysis plots
- Statistical metrics (doubling time, max growth rate)
- Comparative visualizations
- Exported analysis files

## Dependencies
- Requires `growth_analysis_module.m`

## Requirements
- MATLAB (R2019b or later)
- Image Processing Toolbox
- Curve Fitting Toolbox

## Author
Avichay & Daniel

## Status
Active Development
