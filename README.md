# RECC-scenarios

This repository has two main components:

1. The target table excel sheet template - *scenario_target_tables.xlsx*
2. Python script in the *interpolate_target_tables* folder that does the following:
	1. Read the values in the target tables' sheets
	2. Interpolate them for intermediate years (using a spline interpolation algorithm)
	3. Create ODYM format data files (excel sheets) with the full interpolated time series data, one file for each target table sheet

Questions, comments, please contact Tomer tomer.fishman@idc.ac.il
https://zenodo.org/badge/237208012.svg
