# RECC-scenarios
This is the repository for a prospective scenario formulation and quantification approach. The method is described in:

> Tomer Fishman, Niko Heeren, Stefan Pauliuk, Peter Berrill, Qingshi Tu, Paul Wolfram, and Edgar G. Hertwich. “A Comprehensive Set of Global Scenarios of Housing, Mobility, and Material Efficiency for Material Cycles and Energy Systems Modeling.” Journal of Industrial Ecology 25, no. 2 (March 31, 2021): 305–20. https://doi.org/10.1111/jiec.13122.

It is also a component of the RECC (Resource Efficiency and Climate change) project.

This repository has two main components:

1. The target table excel sheet template - *scenario_target_tables.xlsx*
2. Python script in the *interpolate_target_tables* folder that does the following:
	1. Read the values in the target tables' sheets
	2. Interpolate them for intermediate years (using a spline interpolation algorithm)
	3. Create ODYM format data files (excel sheets) with the full interpolated time series data, one file for each target table sheet

Questions, comments, please contact Tomer t.fishman@cml.leiden-univ.nl


[![DOI](https://zenodo.org/badge/237208012.svg)](https://zenodo.org/badge/latestdoi/237208012)
