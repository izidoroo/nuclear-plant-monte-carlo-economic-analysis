# JEK2 Nuclear PP – Monte Carlo Economic Analysis
Stochastic Monte Carlo model for the economic assessment of the proposed second unit of the Krško nuclear power plant (JEK2) in Slovenia.

A full techno-economic simulation model for the proposed second unit of the Krško nuclear power plant (JEK2).
The script computes NPV, LCOE, and the breakeven electricity price (NPV = 0), and propagates uncertainty using Monte Carlo sampling from multiple distributions.
The model includes detailed cost structures, financing architecture, FCFF calculations, and shutdown risk modelling.

## Repository Contents
`src/jek2_monte_carlo.py` Main Python simulation script (Monte Carlo, NPV, LCOE, FCFF)

`data/jek2_input_data.xlsx` Input parameters and scenario assumptions

`docs/Economic_study_JEK2_MZPP_2025.pdf` Full techno-economic study (MZPP, 2025)

## Requirements

```bash
pip install -r requirements.txt
