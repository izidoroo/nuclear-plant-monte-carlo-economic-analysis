# -*- coding: utf-8 -*-
"""
Created on Sun Oct 20 13:59:27 2024

@author: izidoro

Stochastic techno-economic model for a nuclear power plant (JEK2):
Monte Carlo simulation assessing the economic feasibility of investing
in JEK2 by calculating NPV, LCOE, and the breakeven electricity price (NPV = 0),
and comparing the latter two with projected wholesale electricity prices.

Context
    This script was developed within an independent, pro bono analytical project on the
    economics of the proposed second unit at Krško (JEK2).

Methodological overview
    The model constructs a complete techno-economic representation of a nuclear
    investment, including:
        • CAPEX schedules, O&M cost structures, fuel,
          decommissioning, waste management, insurance, taxes, and extraordinary
          repairs;
        • explicit financing architecture (equity, bonds, concessional loans,
          EXIM financing) with interest capitalisation and construction-period
          borrowing;
        • yearly FCFF calculations, WACC, NPV, LCOE, and the endogenous
          electricity price that yields NPV = 0.
    Uncertainty is propagated using a Monte Carlo simulation:
        • stochastic sampling from triangular, KDE-based, Student-t-bounded,
          discrete, and piecewise distributions;
        • conditional modelling of operating regimes (remont vs. non-remont)
          and political/technical shutdown risks;

Outputs
    • Probability distributions for NPV, LCOE, and the breakeven electricity price
    (NPV = 0)
    • Comparison of the LCOE and breakeven price with projected wholesale electricity price.
    • Matplotlib visualisations summarising distributions and comparative metrics.

Note
    Inline comments throughout the script are written in Slovenian,
    reflecting the language of the original research context.

"""

# %% VNOS KNJIŽNIC
import os
import xlwings as xw
import numpy as np
import pandas as pd
from scipy.stats import  kstest, norm, lognorm, gamma, expon, triang, chi2, weibull_min, pareto, t, gumbel_r, uniform, gaussian_kde
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from scipy.integrate import quad
from scipy.optimize import minimize_scalar
import numpy
from collections import Counter

#ŠTEVILO ITERACIJ MONTE CARLO ANALIZE
#definiraš koliko iteracij se bo izvedlo pri Monte Carlo simulaciji
n_trials = 10000

# AKTIVNI GRAFI - zakomentiraj/odkomentiraj po želji za prikaz več/manj grafov
active_plots = {
    "Razporeditev končnega OCC za JEK2 (OCC_final)",
    "Razporeditev neto sedanje vrednosti (NPV) projekta JEK2",
    "Razporeditev prodajne cene električne energije iz JEK2",
    "Razporeditev prodajne cene električne energije iz JEK2 ter primerjava med povprečno prodajno in borzno ceno elektrike",
    "Razporeditev lastne cene električne energije iz JEK2",
    "Povprečna prodajna in lastna cena projekta JEK2 ter borzna cena električne energije",
    # --- zakomentirane grafike (odkomentiraj za prikaz) ---
    # "Razporeditev življenske dobe JEK2",
    # "Inštalirana moč JEK2",
    # "Dolžina gradnje JEK2",
    # "Razporeditev povprečnega faktorja obremenitve",
    # "Razporeditev faktorja obremenitve v letih brez remonta",
    # "Razporeditev faktorja obremenitve v letih remonta",
    # "Razporeditev variabilnih stroškov obratovanja in vzdrževanja (O&M)",
    # "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 20. let",
    # "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 21-40 let",
    # "Razporeditev sroškov obratovanja in vzdrževanja (O&M) za obdboje 41-60 let",
    # "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 61-80 let",
    # "Razporeditev stroškov goriva",
    # "Razporeditev stroškov razgradnje",
    # "Razporeditev upravljanja z odpadki",
    # "Razporeditev začetnega OCC (OCC_initial)",
    # "Razporeditev stroška lastniškega kapitala",
    # "Razporeditev stroška kredita izvozno-uvozne banke (EXIM)",
    # "Razporeditev stroška državno jamčenega kredita pri komercialnih bankah",
    # "Razporeditev stroška državnih obveznic",
    # "Razporeditev tehtanega povprečja stroškov kapitala (WACC)",
    # "Razporeditev političnega tveganja tekom gradnje",
    # "Razporeditev političnega tveganja tekom obratovanja",
    # "Razporeditev tehničnega tveganja",   
    # "Razporeditev borznih cen električne energije osnovnih treh scenarijev",
    # "Razporeditev cen električne energije znotraj treh osnovnih scenarijev in scenarija energetkse krize",
    # "Razporeditev pojavljanja energetske krize",
    # "Razporeditev izbrane scenarije cen električne energije",
}


#reprodukcija rezultatov
np.random.seed(42)

# %% POVEZAVA Z EXCELOM
# Definiramo pot do excel dokumenta
file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "C:\\workspace\\JEK2 - python koda\\Final price of JEK2 - monte carlo simulation - github.xlsx")


wb = xw.Book(file_path)

# %% ŽIVLJENSKA DOBA
sheet = wb.sheets['Življenska doba JEK2']
lifetime_values = sheet.range('A2:A5').value
lifetime_weights = sheet.range('B2:B5').value

#sample lifetimes temelječ na definiranih utežeh
def sample_lifetime_years_with_weights(n_trials):
    return np.random.choice(lifetime_values, size=n_trials, p=lifetime_weights)

# %% INŠTALIRANA KAPACITETA - KONČANO - PAZI DA JE POVSOD PRAVILNO S MW
sheet = wb.sheets['Inštalirana kapaciteta']
installed_capacities = sheet.range('B1:B3').value

def choose_installed_capacity(installed_capacities):
    #discrete uniform distribution
    return np.random.choice(installed_capacities)

# %%  ISKANJE BEST FIT DISTRIBUCIJE - DOKONČANO 
distributions = {
    "Normalna": norm,
    "Lognormalna": lognorm,
    "gamma": gamma,
    #"beta": beta,
    "exponential": expon,
    "Trikotna": triang,
    "chi2": chi2,
    "weibull_min": weibull_min,
    "pareto": pareto,
    #"cauchy": cauchy,
    "Studentova t": t,  
    "gumbel": gumbel_r,  
    "Enakomerna": uniform  
}

#funkcija da najde best fit distribucijo, uporabljaš Kolmogorov-Smirnov test
def fit_distributions(data):
    #rezultati: name of distribution as key, D, p and params as values
    results = {}
    
    for name, distribution in distributions.items():
        try:
            #fit-a distribucijo podatkom
            params = distribution.fit(data)
            
            #izvedeš Kolmogorov-Smirnov test
            D, p_value = kstest(data, distribution.name, args=params)
            
            results[name] = {
                "D-statistic": D,
                "p-value": p_value,
                "params": params,
            }
        except Exception as e:
            print(f"Could not fit {name}: {e}")
            continue
    
    #identificiraj best fit temeljče na najnižji D-statistiki
    best_fit_name = min(results, key=lambda k: results[k]["D-statistic"])
    best_fit_data = results[best_fit_name]
    
    #print(f"Best fitting distribution: {best_fit_name}")
    #print(f"Parameters: {best_fit_data}")
    
    return best_fit_name, best_fit_data
# %% ISKANJE KDE - DOKONČANO
def find_kde(data):
    #vrne ocenjenje density values za podatke
    kde = gaussian_kde(data)
    return kde
# %% CAPACITY FACTOR (sedaj trikotna distribucija)- KONČANO
sheet = wb.sheets['Capacity factor']
min_capacity_factor = sheet.range('B2').value
mode_capacity_factor = sheet.range('B3').value
max_capacity_factor = sheet.range('B4').value

def calculate_capacity_factor(min_capacity_factor, mode_capacity_factor, max_capacity_factor):

    capacity_factor_average = np.random.triangular(min_capacity_factor, mode_capacity_factor, max_capacity_factor)
    #2/3 let je remont, 1/3 let pa je leto brez remonta
    weight_remont = 0.667
    weight_ne_remont = 0.333
    #izhajajoč iz enačbe: capacity factor average = weight remont * CF remont + weight neremont * CF neremont
    #vemo, kaj je capacity factor povprečen, mamo uteži, vemo tudi razmerje CF remont = CF neremont * 11/12
    cf_ne_remont = capacity_factor_average / (weight_remont * (11 / 12) + weight_ne_remont)
    cf_remont = cf_ne_remont * (11 / 12)
                        
    return capacity_factor_average, cf_ne_remont, cf_remont
        
# %% ODHODKI
# %% O&M OPERATION AND MAINTENANCE COST - DOKONČANO
sheet = wb.sheets['O&M']
min_OM_20 = sheet.range('B19').value
mean_OM_20 = sheet.range('C19').value
max_OM_20 = sheet.range('D19').value
min_OM_40 = sheet.range('B20').value
mean_OM_40 = sheet.range('C20').value
max_OM_40 = sheet.range('D20').value
min_OM_60 = sheet.range('B21').value
mean_OM_60 = sheet.range('C21').value
max_OM_60 = sheet.range('D21').value
min_OM_80 = sheet.range('B22').value
mean_OM_80 = sheet.range('C22').value
max_OM_80 = sheet.range('D22').value
escalation_rate_O_and_M = sheet.range('B23').value
var_OM_value_min = sheet.range('B24').value
var_OM_value_mean = sheet.range('C24').value
var_OM_value_max = sheet.range('D24').value
prob_max_OMcost = sheet.range('B25').value
prob_min_OMcost = sheet.range('B26').value
shareof_OM_duringconstruction = sheet.range('B27').value
average_OMcost_over80years = sheet.range('B28').value

#definiraj piecewise linear probability density function (PDF)
#x = single value, kjer bo PDF ocenjen
def piecewise_pdf_OM(x, min_OM, mean_OM, max_OM):
    if min_OM <= x < mean_OM:
        
        return prob_min_OMcost * (x - min_OM) / (mean_OM - min_OM)
    elif mean_OM <= x <= max_OM:
        
        return prob_max_OMcost * (max_OM - x) / (max_OM - mean_OM)
    else:
        return 0

#normaliziraj PDF
#total area = 1
def calculate_totalarea_OM(piecewise_pdf_OM, min_OM, mean_OM, max_OM):
    total_area_OM, _ = quad(lambda x: piecewise_pdf_OM(x, min_OM, mean_OM, max_OM), min_OM, max_OM)
    return total_area_OM

# nova funkcija normalized_piecewise_pdf, ki prilagodi višina originale piecewise_pdf z deljem z total_area.
def normalized_piecewise_pdf_OM(x, total_area_OM, min_OM, mean_OM, max_OM):
    #Tu dodaš () parametre, saj eksplicitno kličeš funkcijo, želiš jo uporabiti
    return piecewise_pdf_OM(x, min_OM, mean_OM, max_OM) / total_area_OM

#generiraš random samples iz izbrane PDF
def sample_from_piecewise_distribution_OM(min_OM, mean_OM, max_OM, total_area_OM):

    x_values = np.linspace(min_OM, max_OM, 1000)

    pdf_values = np.array([normalized_piecewise_pdf_OM(x, total_area_OM, min_OM, mean_OM, max_OM) for x in x_values])
    #Convert PDF to CDF (ključna za inverse transform sampling)
    cdf_values = np.cumsum(pdf_values) / np.sum(pdf_values)

    random_samples = np.random.rand()

    sampled_value_OM = np.interp(random_samples, cdf_values, x_values)
    return sampled_value_OM

def execute_triang_dist_OM_var(mean_OM, min_OM, max_OM):
    c_OM = (mean_OM - min_OM) / (max_OM - min_OM)
    triang_dist_OM = triang(c_OM, loc = min_OM, scale = (max_OM - min_OM))
    var_OM_value = triang_dist_OM.rvs()
    return var_OM_value                    

#O&M v času gradnje
def OMcosts_during_constructionperiod(construction_period, escalation_rate, installed_capacity):
    O_and_M_costs_construction = []
    for year in range(construction_period):
        #*1000 ker imaš EUR/kW, installed capacity pa v MW
        yearly_OM_constructioncost = shareof_OM_duringconstruction * average_OMcost_over80years * ((1 + escalation_rate) ** year) * installed_capacity * 1000
        O_and_M_costs_construction.append((yearly_OM_constructioncost))
    return O_and_M_costs_construction

#O&M prvih 20 let
def calculate_O_and_M_over_20_years(installed_capacity, escalation_rate, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, first_period = 20):

    total_area_OM_20 = calculate_totalarea_OM(piecewise_pdf_OM, min_OM_20, mean_OM_20, max_OM_20)
    # Pridobi fiksno O&M za obdobje 20.let
    O_and_M_20 = sample_from_piecewise_distribution_OM(min_OM_20, mean_OM_20, max_OM_20, total_area_OM_20)
    #Pridobi variabilno O&M za obdobje 20  
    O_and_M_costs_20 = []
    OM_var_costs_20 = []
    #pogledaš če je first period večja od življenske dobe, ki jo lahko politična odločitev zelo skrajša
    if first_period > lifetime_years:
        first_period = lifetime_years
    for year in range(first_period):        
        O_and_M_cost_20 = O_and_M_20 * ((1 + escalation_rate) ** year) * installed_capacity * 1000
        #Deliš s 3 in če je ostanek 0, potem izbereš to if vrstico. Tu imaš ne-remont obdobje, saj skačeš po 3 naprej [0, 3, 6, ...]
        if (year % 3) == 0:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
        #leta, ko je remont
        else:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont
        OM_var_cost_20 = var_OM_value * yearly_ee_produced_mwh * ((1 + escalation_rate) ** year)
        O_and_M_costs_20.append(O_and_M_cost_20)
        OM_var_costs_20.append(OM_var_cost_20)
    
    #sešteješ fiksne in variabilne stroške po letih (za leto 1, za leto 2, ...)
    total_OM_costs_20 = [fixed + variable for fixed, variable in zip(O_and_M_costs_20, OM_var_costs_20)]
    
    #vrneš več spremenljivk, da preveriš, če se vse pravilno računa v excelu
    return total_OM_costs_20, O_and_M_20, O_and_M_costs_20, OM_var_costs_20

# ---- O&M 21-40 let ----
def calculate_O_and_M_over_40_years (installed_capacity, escalation_rate, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, second_period = 40):
    if lifetime_years <= 20:
        #vrni  list namest None, saj potem seštevam in rabim nekaj, da zapiše v excel in da s tem računa, ne pa None
        #štiri vrednosti, saj maš tudi spodaj pri return 4 vrednosti
        total_OM_costs_40 = [] 
        O_and_M_costs_40 = [] 
        OM_var_costs_40 = []
        O_and_M_40 = 0
        return total_OM_costs_40, O_and_M_40, O_and_M_costs_40, OM_var_costs_40
    
    total_area_OM_40 = calculate_totalarea_OM(piecewise_pdf_OM, min_OM_40, mean_OM_40, max_OM_40)
    O_and_M_40 = sample_from_piecewise_distribution_OM(min_OM_40, mean_OM_40, max_OM_40, total_area_OM_40)
    O_and_M_costs_40 = []
    OM_var_costs_40 = []
    if second_period > lifetime_years:
        second_period = lifetime_years
    for year in range(20, second_period): 
        O_and_M_cost_40 = O_and_M_40 * ((1 + escalation_rate) ** year) * installed_capacity * 1000
        #Deliš s 3 in če je ostanek 0, potem izbereš to if vrstico. Tu imaš ne-remont obdobje, saj skačeš po 3 naprej [0, 3, 6, ...]
        if (year % 3) == 0:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
        #leta, ko je remont
        else:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont     
        OM_var_cost_40 = var_OM_value * yearly_ee_produced_mwh * ((1 + escalation_rate) ** year)
        O_and_M_costs_40.append(O_and_M_cost_40)
        OM_var_costs_40.append(OM_var_cost_40)
        
    total_OM_costs_40 = [fixed + variable for fixed, variable in zip(O_and_M_costs_40, OM_var_costs_40)]

    return total_OM_costs_40, O_and_M_40, O_and_M_costs_40, OM_var_costs_40

# ----O&M 41-60 let
def calculate_O_and_M_over_60_years (installed_capacity, escalation_rate, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, third_period = 60):
    if lifetime_years <= 40:
        #vrni  list namest None, saj potem seštevam in rabim nekaj, da zapiše v excel in da s tem računa, ne pa None
        #štiri vrednosti, saj maš tudi spodaj pri return 4 vrednosti
        total_OM_costs_60 = [] 
        O_and_M_costs_60 = [] 
        OM_var_costs_60 = []
        O_and_M_60 = 0
        return total_OM_costs_60, O_and_M_60, O_and_M_costs_60, OM_var_costs_60
    
    #če je življenska doba <=60, pomeni, da jim v obdobju 41-60 let ni treba nadgrajevati JE za delovanje 80 let, zato vzamemo vrednosti od prejšnjih 20let
    if 40 < lifetime_years <= 60:
        total_area_OM_60 = calculate_totalarea_OM(piecewise_pdf_OM, min_OM_40, mean_OM_40, max_OM_40)
        O_and_M_60 = sample_from_piecewise_distribution_OM(min_OM_40, mean_OM_40, max_OM_40, total_area_OM_60)    
    #če je življenska doba > 60, pomeni, da morajo v odbodju 41-60 nadtraditi JE za delovanje 80 let in zato višji stroški
    if lifetime_years > 60:
        total_area_OM_60 = calculate_totalarea_OM(piecewise_pdf_OM, min_OM_60, mean_OM_60, max_OM_60)
        O_and_M_60 = sample_from_piecewise_distribution_OM(min_OM_60, mean_OM_60, max_OM_60, total_area_OM_60)
    
    O_and_M_costs_60 = []
    OM_var_costs_60 = []
    if third_period > lifetime_years:
        third_period = lifetime_years
    for year in range(40, third_period): 
        O_and_M_cost_60 = O_and_M_60 * ((1 + escalation_rate) ** year) * installed_capacity * 1000
        #leta, ko ni remonta
        if (year % 3) == 0:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
        #leta, ko je remont
        else:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont
        OM_var_cost_60 = var_OM_value * yearly_ee_produced_mwh * ((1 + escalation_rate) ** year)
        O_and_M_costs_60.append(O_and_M_cost_60)
        OM_var_costs_60.append(OM_var_cost_60)
        
    total_OM_costs_60 = [fixed + variable for fixed, variable in zip(O_and_M_costs_60, OM_var_costs_60)]
    
    return total_OM_costs_60, O_and_M_60, O_and_M_costs_60, OM_var_costs_60

# ----O&M 61-80 let
def calculate_O_and_M_over_80_years (installed_capacity, escalation_rate, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, fourth_period = 80):
    if lifetime_years <= 60:
        #vrni  list namest None, saj potem seštevam in rabim nekaj, da zapiše v excel in da s tem računa, ne pa None
        #štiri vrednosti, saj maš tudi spodaj pri return 4 vrednosti
        total_OM_costs_80 = [] 
        O_and_M_costs_80 = [] 
        OM_var_costs_80 = []
        O_and_M_80 = 0
        return total_OM_costs_80, O_and_M_80, O_and_M_costs_80, OM_var_costs_80
    
    #če življenska doba med 60 in 80 let, potem imaš nižje O&M stroške
    if 60 < lifetime_years <= 80:
        total_area_OM_80 = calculate_totalarea_OM(piecewise_pdf_OM, min_OM_80, mean_OM_80, max_OM_80)
        O_and_M_80 = sample_from_piecewise_distribution_OM(min_OM_80, mean_OM_80, max_OM_80, total_area_OM_80)
    
    #če življenska doba nad 80, potem moraš nadgradit za delovanje do 100 let
    if lifetime_years > 80:
        total_area_OM_80 = calculate_totalarea_OM(piecewise_pdf_OM, min_OM_60, mean_OM_60, max_OM_60)
        O_and_M_80 = sample_from_piecewise_distribution_OM(min_OM_60, mean_OM_60, max_OM_60, total_area_OM_80)
        
    O_and_M_costs_80 = []
    OM_var_costs_80 = []
    if fourth_period > lifetime_years:
        fourth_period = lifetime_years
    for year in range(60, fourth_period): 
        O_and_M_cost_80 = O_and_M_80 * ((1 + escalation_rate) ** year) * installed_capacity * 1000
        #leta, ko ni remonta
        if (year % 3) == 0:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
        #leta, ko je remont
        else:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont        
        OM_var_cost_80 = var_OM_value * yearly_ee_produced_mwh * ((1 + escalation_rate) ** year)
        O_and_M_costs_80.append(O_and_M_cost_80)
        OM_var_costs_80.append(OM_var_cost_80)
        
    total_OM_costs_80 = [fixed + variable for fixed, variable in zip(O_and_M_costs_80, OM_var_costs_80)]

    return total_OM_costs_80, O_and_M_80, O_and_M_costs_80, OM_var_costs_80

# ----O&M 81-100 let
def calculate_O_and_M_over_100_years (installed_capacity, escalation_rate, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, fifth_period = 100):
    if lifetime_years <= 80:
        #vrni  list namest None, saj potem seštevam in rabim nekaj, da zapiše v excel in da s tem računa, ne pa None
        #štiri vrednosti, saj maš tudi spodaj pri return 4 vrednosti
        total_OM_costs_100 = [] 
        O_and_M_costs_100 = [] 
        OM_var_costs_100 = []
        O_and_M_100 = 0
        return total_OM_costs_100, O_and_M_100, O_and_M_costs_100, OM_var_costs_100
    
    total_area_OM_100 = calculate_totalarea_OM(piecewise_pdf_OM, min_OM_80, mean_OM_80, max_OM_80)
    O_and_M_100 = sample_from_piecewise_distribution_OM(min_OM_80, mean_OM_80, max_OM_80, total_area_OM_100)
        
    O_and_M_costs_100 = []
    OM_var_costs_100 = []
    if fifth_period > lifetime_years:
        fifth_period = lifetime_years
    for year in range(80, fifth_period): 
        O_and_M_cost_100 = O_and_M_100 * ((1 + escalation_rate) ** year) * installed_capacity * 1000
        #leta, ko ni remonta
        if (year % 3) == 0:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
        #leta, ko je remont
        else:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont        
        OM_var_cost_100 = var_OM_value * yearly_ee_produced_mwh * ((1 + escalation_rate) ** year)
        O_and_M_costs_100.append(O_and_M_cost_100)
        OM_var_costs_100.append(OM_var_cost_100)
        
    total_OM_costs_100 = [fixed + variable for fixed, variable in zip(O_and_M_costs_100, OM_var_costs_100)]

    return total_OM_costs_100, O_and_M_100, O_and_M_costs_100, OM_var_costs_100


# OM stroški skozi celotno obdobje (gradnja + obratovanje) - sešteješ šest časovnih vrst
def construct_OMcosts_over_lifetime(O_and_M_costs_construction, total_OM_costs_20, total_OM_costs_40, total_OM_costs_60, total_OM_costs_80, total_OM_costs_100):
    OM_costs_over_lifetime = O_and_M_costs_construction + total_OM_costs_20 + total_OM_costs_40 + total_OM_costs_60 + total_OM_costs_80 + total_OM_costs_100
    return OM_costs_over_lifetime
    
# %% FUEL COSTS - DOKONČANO
sheet = wb.sheets['Fuel cost']
fuel_cost_min = sheet.range('B1').value
fuel_cost_avg = sheet.range('B2').value
fuel_cost_max = sheet.range('B3').value
escalation_rate_fuel_cost = sheet.range('B4').value
prob_min_fuelcost = sheet.range('B5').value
prob_max_fuelcost = sheet.range('B6').value

#definiraj piecewise linear probability density function (PDF)
def piecewise_pdf(x):
    if fuel_cost_min <= x < fuel_cost_avg:
        return prob_min_fuelcost * (x - fuel_cost_min) / (fuel_cost_avg - fuel_cost_min)
    elif fuel_cost_avg <= x <= fuel_cost_max:

        return prob_max_fuelcost * (fuel_cost_max - x) / (fuel_cost_max - fuel_cost_avg)
    else:
        return 0

#normaliziraj PDF: total area 1
total_area_fuelcost, _ = quad(piecewise_pdf, fuel_cost_min, fuel_cost_max)

def normalized_piecewise_pdf(x):
    return piecewise_pdf(x) / total_area_fuelcost

#generate random samples from PDF, uporabljamo inverse transform sampling
def sample_from_piecewise_distribution(num_samples=1):

    x_values = np.linspace(fuel_cost_min, fuel_cost_max, 1000)

    pdf_values = np.array([normalized_piecewise_pdf(x) for x in x_values])
    #Convert PDF to CDF
    cdf_values = np.cumsum(pdf_values) / np.sum(pdf_values)

    random_samples = np.random.rand(num_samples)

    #inverz CDF, da najdeš x vrednosti
    sampled_values = np.interp(random_samples, cdf_values, x_values)

    return sampled_values[0] if num_samples == 1 else sampled_values

#izračunaš stroške goriva preko življenske dobe
def calculate_fuel_costs_over_lifetime(construction_period, lifetime_years, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont):
    fuel_costs_over_lifetime = []

    fuel_cost = sample_from_piecewise_distribution()
    
    #0 v času gradnje JEK2
    for constr_year in range(construction_period):
        fuel_costs_over_lifetime.append(0)
    
    #Vrednosti goriv v času delovanja JEK2
    for year in range(lifetime_years):
        
        #leta, ko ni remonta
        if (year % 3) == 0:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
        #leta, ko je remont
        else:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont
        
        #escalation rate. Prvo leto ^0 = 1, tako da se eskalacijski faktor ne upošteva
        yearly_fuel_cost = fuel_cost * ((1 + escalation_rate_fuel_cost) ** year) * yearly_ee_produced_mwh
        
        fuel_costs_over_lifetime.append(yearly_fuel_cost)
    
    return fuel_costs_over_lifetime, fuel_cost
# %% RAZGRADNJA - DOKONČANO
#računamo sedaj na EUR/MWh
sheet = wb.sheets['Stroški razgradnje']
decom_values_40 = sheet.range('J3:J14').value
decom_values_60 = sheet.range('J17:J28').value
decom_values_80 = sheet.range('J31:J42').value
decom_values_100 = sheet.range('J45:J56').value

min_decom_value_40 = sheet.range('N9').value
max_decom_value_40 = sheet.range('N10').value
min_decom_value_60 = sheet.range('N23').value
max_decom_value_60 = sheet.range('N24').value
min_decom_value_80 = sheet.range('N36').value
max_decom_value_80 = sheet.range('N37').value
min_decom_value_100 = sheet.range('N49').value
max_decom_value_100 = sheet.range('N50').value

kde_decom_40 = find_kde(decom_values_40)
kde_decom_60 = find_kde(decom_values_60)
kde_decom_80 = find_kde(decom_values_80)
kde_decom_100 = find_kde(decom_values_100)

def calculate_decom_costs(lifetime_years_initial, lifetime_years, construction_period, yearly_ee_produced_mwh_ne_remont, yearly_ee_produced_mwh_remont):
    if lifetime_years_initial <= 40:
        kde = kde_decom_40
        min_decom_value = min_decom_value_40
        max_decom_value = max_decom_value_40
        
    elif 40 < lifetime_years_initial <= 60:
        kde = kde_decom_60
        min_decom_value = min_decom_value_60
        max_decom_value = max_decom_value_60
        
    
    elif 60 < lifetime_years_initial <= 80:
        kde = kde_decom_80
        min_decom_value = min_decom_value_80
        max_decom_value = max_decom_value_80
        
    else: 
        kde = kde_decom_100
        min_decom_value = min_decom_value_100
        max_decom_value = max_decom_value_100
    
    decom_costs_over_lifetime = [0] * construction_period
    
    while True:
        decom_cost_eur_mwh = kde.resample(size=1)[0][0]
        if min_decom_value <= decom_cost_eur_mwh <= max_decom_value:
            for year in range(lifetime_years):
                
                #leta, ko ni remonta
                if (year % 3) == 0:
                    yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
                #leta, ko je remont
                else:
                    yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont
                decom_cost_yearly = decom_cost_eur_mwh * yearly_ee_produced_mwh
                decom_costs_over_lifetime.append(decom_cost_yearly)

            return decom_costs_over_lifetime, decom_cost_eur_mwh

    
# %% ODLAGAJANJE ODPADKOV, WASTE DISPOSAL - KONČANO
#računamo sedaj na EUR/MWh
sheet = wb.sheets['Stroški odpadkov']
waste_values_40 = sheet.range('J3:J17').value
waste_values_60 = sheet.range('J20:J34').value
waste_values_80 = sheet.range('J37:J51').value
waste_values_100 = sheet.range('J54:J68').value

#KDE for each range
kde_waste_40 = find_kde(waste_values_40)
kde_waste_60 = find_kde(waste_values_60)
kde_waste_80 = find_kde(waste_values_80)
kde_waste_100 = find_kde(waste_values_100)

#najdi min, max
min_waste_value_40, max_waste_value_40 = min(waste_values_40), max(waste_values_40)
min_waste_value_60, max_waste_value_60 = min(waste_values_60), max(waste_values_60)
min_waste_value_80, max_waste_value_80 = min(waste_values_80), max(waste_values_80)
min_waste_value_100, max_waste_value_100 = min(waste_values_100), max(waste_values_100)

def calculate_waste_costs(lifetime_years_initial, lifetime_years, construction_period, yearly_ee_produced_mwh_ne_remont, yearly_ee_produced_mwh_remont):
    if lifetime_years_initial <= 40:
        kde = kde_waste_40
        min_waste_value = min_waste_value_40
        max_waste_value = max_waste_value_40
    elif 40 < lifetime_years_initial <= 60:
        kde = kde_waste_60
        min_waste_value = min_waste_value_60
        max_waste_value = max_waste_value_60
    
    elif 60 < lifetime_years_initial <= 80:
        kde = kde_waste_80
        min_waste_value = min_waste_value_80
        max_waste_value = max_waste_value_80
        
    else: 
        kde = kde_waste_100
        min_waste_value = min_waste_value_100
        max_waste_value = max_waste_value_100
    
    waste_costs_over_lifetime = [0] * construction_period
    
    #generiraj vrednosti dokler ni v valid rangu
    while True:
        waste_cost_eur_mwh = kde.resample(1)[0][0]
        
        if min_waste_value <= waste_cost_eur_mwh <= max_waste_value:
            for year in range(lifetime_years):
                #leta, ko ni remonta
                if (year % 3) == 0:
                    yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
                #leta, ko je remont
                else:
                    yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont
                    
                waste_cost_yearly = waste_cost_eur_mwh * yearly_ee_produced_mwh
                waste_costs_over_lifetime.append(waste_cost_yearly)
            return waste_costs_over_lifetime, waste_cost_eur_mwh
        
        
# %% PLAČILO OBČINI IN VODARINA - KONČANO
sheet = wb.sheets['Nadomestilo občini in voda']
placilo_obcini_1000 = sheet.range('B2').value
placilo_obcini_1250 = sheet.range('B3').value
placilo_obcini_1650 = sheet.range('B4').value

def calculate_payment_to_municipality(installed_capacity, lifetime_years, construction_period):
    if installed_capacity == 1000:
        placilo_obcini = placilo_obcini_1000
    elif installed_capacity == 1250:
        placilo_obcini = placilo_obcini_1250
    else:
        placilo_obcini = placilo_obcini_1650
    placila_obcini = [placilo_obcini] * (lifetime_years + construction_period)
    return placila_obcini

# %% NEPRIČAKOVANI IZPADI IN STROŠKI POPRAVIL - UNEXPECTED OUTAGES, EXTRAORDINARY REPAIRS - DOKONČANO
#VKLJUČEN V DRUGIH STROŠKOV, ZATO TRENUTNO 0
sheet = wb.sheets['Extraordinary repairs']
average_unexp_outages_per_year = sheet.range('B5').value
extraordinary_repair_cost_engineering = sheet.range('B6').value
extraordinary_repair_cost_electricity = sheet.range('B7').value
costescalation_rate_extraordinaryrepairs = sheet.range('B12').value

# U-shaped kvadratna funkcija da dobiš število in vrednosti nepričakovanih incidentov, ki v povprečju znašajo 2,78 (kot je JEK1 povprečje)
#dobim lifetime_years vrednosti, ki sledijo U-shaped kvadratni funkciji. dobim za vsako leto število nepričakovanih incidentov, ki v povprečju skozi boravnavano obdobje znašajo 2,78 (kot je JEK1 povprečje)
def u_shaped_unexpected_outage(lifetime_years, average_unexp_outages_per_year):
    # število let obratovanja
    years = np.linspace(0, lifetime_years, lifetime_years)
    #srednja vrednost, ki je najnižja vrednost v funkciji
    midpoint = lifetime_years / 2
    #osnovna, neskladirana U-shaped kvadratna funkcija
    unscaled_unexp_outage = (((years - midpoint) ** 2)) / (midpoint ** 2)
    
    #scaling faktor
    scale_factor = average_unexp_outages_per_year / np.mean(unscaled_unexp_outage)
    
    #final U shape function
    scaled_unexp_outages = unscaled_unexp_outage * scale_factor

    return scaled_unexp_outages

def calculate_extraordinary_repairs_over_lifetime(lifetime_years, construction_period, average_unexp_outages_per_year): 
    extraordinary_repair_costs_over_lifetime = [0] * construction_period
    
    #dobim lifetime_years vrednosti, ki sledijo U-shaped kvadratni funkciji. dobim za vsako leto število nepričakovanih incidentov, ki v povprečju skozi boravnavano obdobje znašajo 2,78 (kot je JEK1 povprečje)
    scaled_unexp_outages = u_shaped_unexpected_outage(lifetime_years, average_unexp_outages_per_year)
    
    for year in range(lifetime_years):
        yearly_engineering_cost_outage = scaled_unexp_outages[year] * (extraordinary_repair_cost_engineering * ((1 + costescalation_rate_extraordinaryrepairs) ** year))
        yearly_ee_cost_outage = scaled_unexp_outages[year] * extraordinary_repair_cost_electricity
        yearly_combined_cost_outages = yearly_engineering_cost_outage + yearly_ee_cost_outage
        
        extraordinary_repair_costs_over_lifetime.append(yearly_combined_cost_outages)
    
    scaled_unexp_outages = np.concatenate((np.zeros(construction_period), scaled_unexp_outages))
    
    return extraordinary_repair_costs_over_lifetime, scaled_unexp_outages

# %% STROŠKI NADGRADNJE V PRIMERU 80 LET - KONČANO
sheet = wb.sheets['Stroški nadgradnje na 80 let']
upgradecost_80_yearly_kw = sheet.range('B4').value 

def calculate_costs_upgrade_80(lifetime_years, installed_capacity, construction_period):
    if lifetime_years <= 60:

        return [0] * (construction_period + lifetime_years)
    fullupgradecost_80_yearly = upgradecost_80_yearly_kw * installed_capacity * 1000
    fullupgradecost80_over_lifetime = [0] * lifetime_years

    fullupgradecost80_over_lifetime[40:60] = [fullupgradecost_80_yearly] * 20
    
    #dodaš 0 za čas gradnje čist na začetku, dodaš 0 pred prvo vrednostjo (index 0)
    fullupgradecost80_over_lifetime[:0] = [0] * construction_period
    
    return fullupgradecost80_over_lifetime
# %% ZAVAROVALNIŠKA PREMIJA ZA PRIMER JEDRSKE NESREČE - KONČANO
sheet = wb.sheets['NPP accident insurance premium']
insurance_premium_value = sheet.range('B1').value

def calculate_insurance_premium(yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, construction_period, lifetime_years):
    insurance_premium_over_lifetime = [0] * construction_period
    for year in range(lifetime_years):
        #leta, ko ni remonta
        if (year % 3) == 0:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
        #leta, ko je remont
        else:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont 
        insurance_premium_yearly = insurance_premium_value * yearly_ee_produced_mwh
        insurance_premium_over_lifetime.append(insurance_premium_yearly)
    return insurance_premium_over_lifetime

# %% CAPEX & construction period & cost escalation rate
sheet = wb.sheets['CAPEX']

#occ initial
min_occ_initial = sheet.range('B6').value
max_occ_initial = sheet.range('B7').value

occ_initial_values = sheet.range('G9:G89').value
#očisti vseh neštevilčnik in praznih celic OCC initial
occ_initial_values_cleaned = pd.to_numeric(occ_initial_values, errors='coerce')

occ_initial_values_cleaned = occ_initial_values_cleaned[~np.isnan(occ_initial_values_cleaned)]

#cost escalation rate za CAPEX
cost_escalation_rate_occ = sheet.range('B1').value
#depreciation period for OCC
depreciation_rate = sheet.range('B2').value
depreciation_rate = int(depreciation_rate)

#najbolj optimalna distribucija = student t distribucija
student_t_params = t.fit(occ_initial_values_cleaned)

def get_occ_initial_depreciation_occ_final_over_lifetime(student_t_params, construction_period, lifetime_years):

    occ_initial = t.rvs(*student_t_params)
    #zahtevaš spoštovanje min in max
    while occ_initial < min_occ_initial or occ_initial > max_occ_initial:
        occ_initial = t.rvs(*student_t_params)
        
    #construction cost - 5: 5 let je dodatek od priprave dokumentacije in relokacije energetske infrastrukture ki je dodan zgodovinskim podatkom
    occ_final_eurkw = occ_initial * ((1 + cost_escalation_rate_occ) ** (construction_period - 5 ))* (1 + 0) ** 5
    occ_final = occ_initial * ((1 + cost_escalation_rate_occ) ** (construction_period - 5 )) * (installed_capacity * 1000) * (1 + 0) ** 5
    
    occ_final_depreciation = occ_final / depreciation_rate
    
    # V času gradnje je 0
    occ_final_depreciation_over_lifetime = [0] * construction_period
    
    #da boš videl, če ti je še kaj depreciacije ostalo
    total_depreciation = 0

    # Za vsako leto obratovanj dodaš depreciacijo, če je ta nižje kot depreciacijsko obdobje
    for year in range(lifetime_years):
        if year < depreciation_rate:
            occ_final_depreciation_over_lifetime.append(occ_final_depreciation)
            total_depreciation += occ_final_depreciation
        #0 ko je konec depreciacije
        else:
            occ_final_depreciation_over_lifetime.append(0)

    #poglej če ostanek depreciacije ostane
    remaining_depreciation = occ_final - total_depreciation
    if remaining_depreciation > 0: 
        #dodaš ostanek depreciacije zadnje leto
        occ_final_depreciation_over_lifetime[-1] += remaining_depreciation

    #kako dolga mora bit časovna vrsta (sešteješ gradnjo + življ dobo)
    target_length = construction_period + lifetime_years
    #kako dolga je trenutno časovna vrsta
    current_length = len(occ_final_depreciation_over_lifetime)

    if current_length < target_length:
        #dodaj 0-ke na koncu, da dosežeš željeno dolžino, to je target_length
        occ_final_depreciation_over_lifetime.extend([0] * (target_length - current_length))
    elif current_length > target_length:
        #skrči časovno vrsto, da dosežeš target_length (odvzameš vrednosti od zadaj)
        occ_final_depreciation_over_lifetime = occ_final_depreciation_over_lifetime[:target_length]

    return occ_final_depreciation_over_lifetime, occ_final, occ_initial, occ_final_eurkw

#očisti vseh neštevilčnik in praznih celic construction period
construction_period_values = sheet.range('D9:D89').value
#min in max construction period
construction_period_min = sheet.range('C4').value
construction_period_max = sheet.range('C6').value

construction_period_values_cleaned = pd.to_numeric(construction_period_values, errors = 'coerce')
construction_period_values_cleaned = construction_period_values_cleaned[~np.isnan(construction_period_values_cleaned)]

#ker student t kot best fit ni najboljša in ker imamo nasploh malo podatkov za gradnjo, gremo naprej s KDE 
kde_construction_period = find_kde(construction_period_values_cleaned)

#pridobi construction period
def get_construction_period(kde_construction_period):
 
    #[0][0]: pridobiš skalar
    construction_period = kde_construction_period.resample(size=1)[0][0]

    while (construction_period > construction_period_max or construction_period < construction_period_min):
        construction_period = kde_construction_period.resample(size=1)[0][0]
    
    return construction_period

def get_occ_final_cashflow(occ_final, construction_period, lifetime_years):

    x = np.linspace(1, construction_period, construction_period)
    mean_investmentcost = construction_period / 2
    if construction_period < 10:
        std_dev = construction_period / 6
    elif 10 < construction_period <= 15:
        std_dev = construction_period / 7 
    elif 15 < construction_period <= 20: 
        std_dev = construction_period / 8 
    elif 20 < construction_period <= 25:
        std_dev = construction_period / 9
    else:
        std_dev = construction_period / 10
        
    #generiraj krivuljo normalne distribucije
    normal_curve = np.exp(-0.5 * ((x - mean_investmentcost) / std_dev) ** 2)
    #normaliziraj - vsaka vrednost posebej je zdeljena s seštevkom vrednosti
    normalized_curve = normal_curve / normal_curve.sum()
    #dobiš vrsto vrednosti: skaliraj na celoten CAPEX
    investment_cash_flow = occ_final * normalized_curve
    
    #ustvari occ final cashflow za celotno obdobje
    occ_final_cashflow_over_lifetime = list(investment_cash_flow) + [0] * lifetime_years
        
    return investment_cash_flow, occ_final_cashflow_over_lifetime

# %% COST OF EQUITY - KONČANO

#velika vloga države za zdaj default
velika_vloga_drzave = True 

sheet = wb.sheets['WACC']

"""
#brez vloge države, mala vloga države
min_cost_of_equity_value_nostate = sheet.range('D2').value
mean_cost_of_equity_value_nostate = sheet.range('E2').value
max_cost_of_equity_value_nostate = sheet.range('F2').value

c_equity_nostate = (mean_cost_of_equity_value_nostate - min_cost_of_equity_value_nostate) / (max_cost_of_equity_value_nostate - min_cost_of_equity_value_nostate)

#triangular distribution
triang_dist_cost_of_equity_nostate = triang(c_equity_nostate, loc=min_cost_of_equity_value_nostate, scale=(max_cost_of_equity_value_nostate - min_cost_of_equity_value_nostate))
"""

#z vlogo države, velika vloga države
share_cost_of_equity = sheet.range('B3').value

min_cost_of_equity_value_state = sheet.range('D3').value
mean_cost_of_equity_value_state = sheet.range('E3').value
max_cost_of_equity_value_state = sheet.range('F3').value

#c: relativna lokacija mode(peak)
c_equity_state = (mean_cost_of_equity_value_state - min_cost_of_equity_value_state) / (max_cost_of_equity_value_state - min_cost_of_equity_value_state)

triang_dist_cost_of_equity_state = triang(c_equity_state, loc=min_cost_of_equity_value_state, scale=(max_cost_of_equity_value_state - min_cost_of_equity_value_state))

def cost_of_equity_calculation(velika_vloga_drzave):
    if velika_vloga_drzave:
         cost_of_equity = triang_dist_cost_of_equity_state.rvs()
    #else:
    #     cost_of_equity = triang_dist_cost_of_equity_nostate.rvs()
    return cost_of_equity


# %% COST OF DEBT
sheet = wb.sheets['WACC']

strosek_odobritve_kredita = sheet.range('B10').value

#nižje vrednosti kreditov in obveznice so notri
if velika_vloga_drzave:
    #izvozni kredit pri EXIM banki (dolgoročna posojila) - exim
    share_cost_of_exim = sheet.range('B4').value
    exim_repayment_period = sheet.range('C4').value
    interest_rate_min_exim = sheet.range('D4').value
    interest_rate_mean_exim = sheet.range('E4').value
    interest_rate_max_exim = sheet.range('F4').value

    c_exim = (interest_rate_mean_exim - interest_rate_min_exim) / (interest_rate_max_exim - interest_rate_min_exim)

    triang_dist_cost_of_exim = triang(c_exim, loc=interest_rate_min_exim, scale=(interest_rate_max_exim - interest_rate_min_exim))
    
    #krediti pri bankah z in brez državnega poroštva - loan
    share_cost_of_loan = sheet.range('B5').value
    loan_repayment_period = sheet.range('C5').value
    interest_rate_min_loan = sheet.range('D5').value
    interest_rate_mean_loan = sheet.range('E5').value
    interest_rate_max_loan = sheet.range('F5').value

    c_loan = (interest_rate_mean_loan - interest_rate_min_loan) / (interest_rate_max_loan - interest_rate_min_loan)

    triang_dist_cost_of_loan = triang(c_loan, loc=interest_rate_min_loan, scale=(interest_rate_max_loan - interest_rate_min_loan))
    
    #obveznice - bond
    share_cost_of_bond = sheet.range('B6').value
    bond_repayment_period = sheet.range('C6').value
    interest_rate_min_bond = sheet.range('D6').value
    interest_rate_mean_bond = sheet.range('E6').value
    interest_rate_max_bond = sheet.range('F6').value

    c_bond = (interest_rate_mean_bond - interest_rate_min_bond) / (interest_rate_max_bond - interest_rate_min_bond)

    triang_dist_cost_of_bond = triang(c_bond, loc=interest_rate_min_bond, scale=(interest_rate_max_bond - interest_rate_min_bond))

#vloga države majhna -> ni obveznic, višji krediti pri bankah
else:
    #izvozni kredit pri EXIM banki (dolgoročna posojila) - exim
    share_cost_of_exim = sheet.range('B7').value
    exim_repayment_period = sheet.range('C7').value
    interest_rate_min_exim = sheet.range('D7').value
    interest_rate_mean_exim = sheet.range('E7').value
    interest_rate_max_exim = sheet.range('F7').value

    c_exim = (interest_rate_mean_exim - interest_rate_min_exim) / (interest_rate_max_exim - interest_rate_min_exim)

    triang_dist_cost_of_exim = triang(c_exim, loc=interest_rate_min_exim, scale=(interest_rate_max_exim - interest_rate_min_exim))

    #krediti pri bankah z in brez državnega poroštva - loan
    share_cost_of_loan = sheet.range('B8').value
    loan_repayment_period = sheet.range('C8').value
    interest_rate_min_loan = sheet.range('D8').value
    interest_rate_mean_loan = sheet.range('E8').value
    interest_rate_max_loan = sheet.range('F8').value

    c_loan = (interest_rate_mean_loan - interest_rate_min_loan) / (interest_rate_max_loan - interest_rate_min_loan)

    triang_dist_cost_of_loan = triang(c_loan, loc=interest_rate_min_loan, scale=(interest_rate_max_loan - interest_rate_min_loan))

    #obveznice - bond
    share_cost_of_bond = sheet.range('B9').value
    bond_repayment_period = sheet.range('C9').value
    interest_rate_min_bond = sheet.range('D9').value
    interest_rate_mean_bond = sheet.range('E9').value
    interest_rate_max_bond = sheet.range('F9').value

    c_bond = (interest_rate_mean_bond - interest_rate_min_bond) / (interest_rate_max_bond - interest_rate_min_bond)

    triang_dist_cost_of_bond = triang(c_bond, loc=interest_rate_min_bond, scale=(interest_rate_max_bond - interest_rate_min_bond))


def calculate_cost_of_debt(share_cost_of_debt, occ_final, triang_dist_cost_of_debt, debt_repayment_period, cash_flow_interest, construction_period, lifetime_years):
    
    debt_repayment_period = int(debt_repayment_period)
    
    #glavnica
    principal = share_cost_of_debt * occ_final
    interest_rate = triang_dist_cost_of_debt.rvs()
    
    #Iščeš prvi i (1,2,3), kjer je x (prava vrednosti z časovnega vektorja) ne-0
    start_year = next((i for i, x in enumerate(cash_flow_interest) if x != 0), None)
    
    #ves dolg se poplača, četudi gre kaj narobe z investicijo
    #to je smiselna predpostavka, saj so vsi dolgovi vezani tudi na državo
    if debt_repayment_period > (lifetime_years + construction_period - start_year):
        debt_repayment_period = (lifetime_years + construction_period - start_year)   
   
    principal_per_year = principal / debt_repayment_period

    yearly_principal = []
    yearly_interest = []
    yearly_debt_payment = []
    yearly_balance = []
    
    for year in range(debt_repayment_period):
        
        if year == 0:
            #glavnica - plača vsako leto enako
            yearly_principal.append(principal_per_year)
            
            #obresti se izračunajo kot zmnožek obrestne mere in polne glavnice
            interest = principal * interest_rate
            
            #dodaš še strošek odobritve kredita
            letni_strosek_odobritve_kredita = principal * strosek_odobritve_kredita
            
            interest = interest + letni_strosek_odobritve_kredita
            
            yearly_interest.append(interest)
            
            #anuiteta kot seštevek glavnice in obresti
            full_yearly_payment_for_debt = principal_per_year + interest
            yearly_debt_payment.append(full_yearly_payment_for_debt)
            
            #balans: zniževanje še neodplačane vsote z glavnico
            balance = principal - principal_per_year
            yearly_balance.append(balance)
                   
        else:
            #glavnica
            yearly_principal.append(principal_per_year)
            
            #obresti
            interest = yearly_balance[year - 1] * interest_rate
            yearly_interest.append(interest)
            
            #anuiteta
            full_yearly_payment_for_debt = principal_per_year + interest
            yearly_debt_payment.append(full_yearly_payment_for_debt)
            
            #balans
            balance = yearly_balance[year - 1] - principal_per_year
            yearly_balance.append(balance)
    
    #da se nato shrani časovno vrsto različnih obresti
    interests_over_lifetime = [0] * (construction_period + lifetime_years)
    
    #da se da kredit v pravo časovno vrsto
    if start_year is not None:
        for i, interest in enumerate(yearly_interest): 
            interests_over_lifetime[start_year + i] = interest
    
    return  interest_rate, yearly_interest, interests_over_lifetime

# %% DOLOČITEV, KDAJ ZNOTRAJ DOBE GRADNJE NASTANEJO KREDITI
def define_cashflow_of_credits(investment_cash_flow, occ_final):
    #koliko prispevka vsak vir financiranja
    equity_amount = occ_final * share_cost_of_equity
    bond_amount = occ_final * share_cost_of_bond
    exim_amount = occ_final * share_cost_of_exim
    loan_amount = occ_final * share_cost_of_loan

    # določi preostanek vrednosti, ki so sedaj še polne
    remaining_equity = equity_amount
    remaining_bond = bond_amount
    remaining_exim = exim_amount
    remaining_loan = loan_amount

    equity_cash_flow = []
    bond_cash_flow = []
    exim_cash_flow = []
    loan_cash_flow = []

    #vsako leto investicijske dobe
    #vsak loop je ena cifra (allocated amount ali pa 0) prilepljena
    for year, yearly_cash_flow in enumerate(investment_cash_flow):
        # določiš denarni tok prvo na lastniški kapital
        if remaining_equity > 0:
            allocated_equity = min(remaining_equity, yearly_cash_flow)
            equity_cash_flow.append(allocated_equity)
            remaining_equity -= allocated_equity
            yearly_cash_flow -= allocated_equity
        else:
            equity_cash_flow.append(0)

        # naslednja stopnja so obveznice
        if remaining_bond > 0:
            allocated_bond = min(remaining_bond, yearly_cash_flow)
            bond_cash_flow.append(allocated_bond)
            remaining_bond -= allocated_bond
            yearly_cash_flow -= allocated_bond
        else:
            bond_cash_flow.append(0)

        # EXIM krediti
        if remaining_exim > 0:
            allocated_exim = min(remaining_exim, yearly_cash_flow)
            exim_cash_flow.append(allocated_exim)
            remaining_exim -= allocated_exim
            yearly_cash_flow -= allocated_exim
        else:
            exim_cash_flow.append(0)

        #na koncu sfinacirajo investicijo bančni krediti
        if remaining_loan > 0:
            allocated_loan = min(remaining_loan, yearly_cash_flow)
            loan_cash_flow.append(allocated_loan)
            remaining_loan -= allocated_loan
            yearly_cash_flow -= allocated_loan
        else:
            loan_cash_flow.append(0)
            
    return equity_cash_flow, bond_cash_flow, exim_cash_flow, loan_cash_flow


# %% VZETI VEČ DODATNIH KREDITOV ZA POPLAČILO OBRESTI TEKOM GRADNJE, KO ŠE NI PRIHODKOV OD OBRATOVANJA + NAREDITI PRAVILNO ČASOVNO VRSTO - KONČANO
def calculate_additional_loan_during_construction(
    bond_interests_over_lifetime, 
    exim_interests_over_lifetime, 
    loan_interests_over_lifetime, 
    construction_period, 
    lifetime_years, 
    loan_interest_rate,
    loan_repayment_period=20,
):

    start_year = next(
        (
            i for i in range(len(bond_interests_over_lifetime))
            if bond_interests_over_lifetime[i] > 0 or 
               exim_interests_over_lifetime[i] > 0 or 
               loan_interests_over_lifetime[i] > 0
        ),
        #vrne None, če ne najde nobene vrednosti več kot 0
        None
    )
    #če je None, potem naredi časovno vrsto 0 in gre ven iz funkcije (to rabimo, ker moramo potem računat)
    if start_year is None:

            all_additionalloans_interests_over_lifetime = [[0] * (construction_period + lifetime_years)]
            return all_additionalloans_interests_over_lifetime
    
    end_year = construction_period
    
    #list of lists
    all_additionalloans_interests_over_lifetime = [[0] * (construction_period + lifetime_years)]
    
    for loan_start_year in range(start_year, end_year):
        yearly_loan_amount_sum = (bond_interests_over_lifetime[loan_start_year] + 
        exim_interests_over_lifetime[loan_start_year] +
        loan_interests_over_lifetime[loan_start_year] +       
        sum(one_interest_vector_list[loan_start_year] for one_interest_vector_list in all_additionalloans_interests_over_lifetime))
        
        
        # Izračunat kakšna je resnična repayment_period, zakaj? da ti lifetime_year lost of political support ne zniža preveč
        available_loan_repaymnent_period = construction_period - loan_start_year + lifetime_years
        loan_repayment_period = min(available_loan_repaymnent_period, loan_repayment_period)

        #calculate yearly payments
        principal_per_year = yearly_loan_amount_sum / loan_repayment_period

        yearly_principal = []
        yearly_interest = []
        yearly_debt_payment = []
        yearly_balance = []
        
        for year in range(loan_repayment_period):
            if year == 0:
                #first year of payment
                yearly_principal.append(principal_per_year)
                interest = yearly_loan_amount_sum * loan_interest_rate
                yearly_interest.append(interest)
                full_yearly_payment = principal_per_year + interest
                yearly_debt_payment.append(full_yearly_payment)
                balance = yearly_loan_amount_sum - principal_per_year
                yearly_balance.append(balance)
            else:
                #subsequent years
                yearly_principal.append(principal_per_year)
                interest = yearly_balance[year - 1] * loan_interest_rate
                yearly_interest.append(interest)
                full_yearly_payment = principal_per_year + interest
                yearly_debt_payment.append(full_yearly_payment)
                balance = yearly_balance[year - 1] - principal_per_year
                yearly_balance.append(balance)
        
        #narediš časovno vrsto, v katero vstaviš na pravo mesto obresti od dodatnih kreditov
        additional_loan_interests_over_lifetime = [0] * (construction_period + lifetime_years)
        
        #s tem nastaviš na pravo mesto obresti znotraj celotne časovne vrste
        for y, interest in enumerate(yearly_interest):
           additional_loan_interests_over_lifetime[loan_start_year + y] = interest
        
        #dobiš listov listov, kjer so shranjene vse popolne časovne vrste obresti dodatnih kreditov 
        all_additionalloans_interests_over_lifetime.append(additional_loan_interests_over_lifetime)   
            
    return all_additionalloans_interests_over_lifetime


# %% IZRAČUN WACC - KONČANO
sheet = wb.sheets['Ostale postavke']
davek_na_dobiček = sheet.range('B1').value

def calculate_WACC(cost_of_equity, cost_of_exim, cost_of_loan, cost_of_bond):
    cumulative_share = share_cost_of_equity + share_cost_of_exim + share_cost_of_loan + share_cost_of_bond
    WACC =(
    ((share_cost_of_equity / cumulative_share) * cost_of_equity) +
    ((share_cost_of_exim / cumulative_share) * cost_of_exim * (1-davek_na_dobiček)) +
    ((share_cost_of_loan / cumulative_share) * cost_of_loan * (1-davek_na_dobiček)) +
    ((share_cost_of_bond / cumulative_share) * cost_of_bond * (1-davek_na_dobiček))
    )
    return WACC   

# %% TVEGANJA - KONČANO
sheet = wb.sheets['Tveganja']
lost_of_political_support_construction_prob = sheet.range('B1').value
lost_of_political_support_operation_prob = sheet.range('B2').value
technical_reason_closure_prob = sheet.range('B3').value

# %% PRIHODKI
# %% PRODAJA POTRDIL O IZVORU - KONČANO
sheet = wb.sheets['Potrdila o izvoru']
potrdila_o_izvoru = sheet.range('B1').value
def potrdilaoizvoru_prihodki_over_lifetime(construction_period, lifetime_years, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont):
    
    prihodki_potrdilaoizvoru_over_lifetime = []
    
    for yearconstruction in range(construction_period):
        prihodki_potrdilaoizvoru_over_lifetime.append(0)
    
    for year in range(lifetime_years):
        #leta, ko ni remonta
        if (year % 3) == 0:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
        #leta, ko je remont
        else:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont 
        
        yearly_potrdilaoizvoru_prihodki = potrdila_o_izvoru * yearly_ee_produced_mwh
        prihodki_potrdilaoizvoru_over_lifetime.append(yearly_potrdilaoizvoru_prihodki)
    
    return prihodki_potrdilaoizvoru_over_lifetime
# %% PRODAJA ELEKTRIČNE ENERGIJE - KONČANO
sheet = wb.sheets['Cena elektrike']
#osnovni scenariji
prob_low_scenario = sheet.range('B2').value
mean_low_scenario = sheet.range('C2').value
stdev_low_scenario = sheet.range('D2').value
prob_mean_scenario = sheet.range('B3').value
mean_mean_scenario = sheet.range('C3').value
stdev_mean_scenario = sheet.range('D3').value
prob_high_scenario = sheet.range('B4').value
mean_high_scenario = sheet.range('C4').value
stdev_high_scenario = sheet.range('D4').value
scenarios = ['Scenarij nizkih cen', 'Osnovni scenarij', 'Scenarij visokih cen']
weights_scenarios = [prob_low_scenario, prob_mean_scenario, prob_high_scenario]
real_escalation_rate_ee = sheet.range('B6').value
#scenarij če so visoke cene EE
prob_high_ee_prices = sheet.range('B8').value
high_prices_min = sheet.range('B10').value
high_prices_mean = sheet.range('B11').value
high_prices_max = sheet.range('B12').value
c_highprices_dist = (high_prices_mean - high_prices_min) / (high_prices_max - high_prices_min)
loc_highprices_dist = high_prices_min
scale_highprices_dist = high_prices_max - high_prices_min
#vzpostavi triangularno distribucijo za scenarij visokih cen EE
triang_dist_highprices = triang(c_highprices_dist, loc=loc_highprices_dist, scale=scale_highprices_dist)

def electricity_prihodki_over_lifetime(construction_period, lifetime_years, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont):    
    ee_prihodki_over_lifetime =[]

    chosen_scenario = np.random.choice(scenarios, size=1, replace=False, p=weights_scenarios)[0]
    
    scenario_parameters = {
        'Scenarij nizkih cen': (mean_low_scenario, stdev_low_scenario),
        'Osnovni scenarij': (mean_mean_scenario, stdev_mean_scenario),
        'Scenarij visokih cen': (mean_high_scenario, stdev_high_scenario)
    }
    
    chosen_mean_scenario, chosen_stdev_scenario = scenario_parameters[chosen_scenario]
    ee_price = np.random.normal(loc=chosen_mean_scenario, scale=chosen_stdev_scenario, size = 1)[0]
    
    ee_prihodki_over_lifetime.extend([0] * construction_period)
    
    #narediš vektor cen EE, saj lahko tako preveriš naknadno, če vse pravilno računa
    cene_elektrike_over_lifetime = [0] * construction_period    
    
    #izračunaj ceno EE za osnovni in high prices scenario      
    for year in range(lifetime_years):  
        #če se zgodijo visoke cene
        high_prices_happen = np.random.binomial(1, prob_high_ee_prices)        
        if high_prices_happen:
             #vzeti eno vrednost za visoke cene EE 
             ee_price_used = triang_dist_highprices.rvs()
        else:
             ee_price_used = ee_price
        #leta, ko ni remonta
        if (year % 3) == 0:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_ne_remont
        #leta, ko je remont
        else:
            yearly_ee_produced_mwh = yearly_ee_produced_mwh_remont 
        yearly_ee_prihodki = (ee_price_used * ((1 + real_escalation_rate_ee) ** year)) * yearly_ee_produced_mwh
        
        ee_prihodki_over_lifetime.append(yearly_ee_prihodki)
        cene_elektrike_over_lifetime.append(ee_price_used)
        
    return ee_prihodki_over_lifetime, cene_elektrike_over_lifetime, chosen_scenario, ee_price, high_prices_happen


# %% OBRATNI KAPITAL
dnevi_vezave_terjatev = 30 
dnevi_vezave_obveznosti = 15 

# %% SEŠTEVANJE PO STOLPCIH Z UPORABO NUMPY ARRAY-OV
def sum_columns_by_year(*lists):

    return np.sum(lists,axis=0)

# %% ISKANJE PRODAJNE CENE EE IZ JEK2, DA JE NPV = 0 - KONČANO

def simulate_fcff_given_price(
    price_of_electricity,
    construction_period,
    lifetime_years,
    yearly_ee_produced_mwh_remont,
    yearly_ee_produced_mwh_ne_remont,
    # Odhodki - te ostanejo enaki:
    skupni_odhodki_po_letih,
    # Prihodki od potrdil o izvoru - te ostanejo enaki:
    prihodki_potrdilaoizvoru_over_lifetime,
    # Depreciation and CAPEX - to ostane enako
    occ_final_depreciation_over_lifetime,
    occ_final_cashflow_over_lifetime,
    # Interest arrays - te ostanejo enaki:
    skupne_obresti,  
    #da sešteveš skupaj
    sum_columns_by_year,
    #Davki in working capital:
    davek_na_dobiček,
    dnevi_vezave_terjatev,
    dnevi_vezave_obveznosti
):

    electricity_revenues = []
    #ni produkcije med konstrucijo
    for _ in range(construction_period):
        electricity_revenues.append(0)
    #produkcija tekom obratovanja
    for year in range(lifetime_years):
        if (year % 3) == 0:
            production_mwh = yearly_ee_produced_mwh_ne_remont
        else:
            production_mwh = yearly_ee_produced_mwh_remont
        
        yearly_revenue = price_of_electricity * production_mwh
        electricity_revenues.append(yearly_revenue)
    
    electricity_revenues = np.array(electricity_revenues)

    prihodki_potrdilaoizvoru_over_lifetime = np.array(prihodki_potrdilaoizvoru_over_lifetime)
        
    total_revenues = electricity_revenues + prihodki_potrdilaoizvoru_over_lifetime

    total_costs = np.array(skupni_odhodki_po_letih)

    ebitda = total_revenues - total_costs

    occ_depreciation = np.array(occ_final_depreciation_over_lifetime)
    ebit = ebitda - occ_depreciation

    skupne_obresti = np.array(skupne_obresti)

    ebt = ebit - skupne_obresti

    tax_array = []
    for val in ebt:
        if val <= 0:
            tax_array.append(0)
        else:
            tax_array.append(val * davek_na_dobiček)
    tax_array = np.array(tax_array)

    net_income = ebt - tax_array

    receivables = total_revenues * (dnevi_vezave_terjatev / 365.0)
    payables    = total_costs * (dnevi_vezave_obveznosti / 365.0)

    delta_receivables = np.zeros_like(receivables)
    delta_payables    = np.zeros_like(payables)
    #leto 0
    delta_receivables[0] = 0 - receivables[0]
    delta_payables[0]    = payables[0] - 0
    # naslednja leta
    delta_receivables[1:] = receivables[:-1] - receivables[1:]
    delta_payables[1:] = payables[1:] - payables[:-1]

    spr_obratnega_kapitala_po_letih = delta_payables + delta_receivables

    skupne_obresti_prilagojene_za_davek = skupne_obresti * (1 - davek_na_dobiček)

    occ_final_capex = np.array(occ_final_cashflow_over_lifetime)
    
    fcff = (
        net_income
        + occ_depreciation
        - occ_final_capex
        + skupne_obresti_prilagojene_za_davek
        - spr_obratnega_kapitala_po_letih
    )

    return fcff

def compute_npv(fcff, wacc):
    # diskontiraj vsako leto
    years = np.arange(1, len(fcff)+1)
    return (fcff / ((1 + wacc)**years)).sum()

def find_electricity_price_npv0(
    construction_period,
    lifetime_years,
    yearly_ee_produced_mwh_remont,
    yearly_ee_produced_mwh_ne_remont,
    skupni_odhodki_po_letih,
    prihodki_potrdilaoizvoru_over_lifetime,
    occ_final_depreciation_over_lifetime,
    occ_final_cashflow_over_lifetime,
    skupne_obresti,
    davek_na_dobiček,
    dnevi_vezave_terjatev,
    dnevi_vezave_obveznosti,
    wacc,
):

    def objective(price_of_electricity):
        fcff = simulate_fcff_given_price(
            price_of_electricity,
            construction_period,
            lifetime_years,
            yearly_ee_produced_mwh_remont,
            yearly_ee_produced_mwh_ne_remont,
            # Odhodki - te ostanejo enaki:
            skupni_odhodki_po_letih,
            # Prihodki od potrdil o izvoru - te ostanejo enaki:
            prihodki_potrdilaoizvoru_over_lifetime,
            # Depreciation in CAPEX - to ostane enako
            occ_final_depreciation_over_lifetime,
            occ_final_cashflow_over_lifetime,
            # Interest arrays - te ostanejo enaki:
            skupne_obresti,  
            #da sešteveš skupaj
            sum_columns_by_year,
            #Davki in working capital:
            davek_na_dobiček,
            dnevi_vezave_terjatev,
            dnevi_vezave_obveznosti
    )
        #NPV izračunaj
        npv_value = compute_npv(fcff, wacc)
        #minimize abs(NPV)
        return abs(npv_value)

    result = minimize_scalar(objective, bounds=(0, 1500), method='bounded')
    return result.x


# %% ZAKLJUČEK - RAČUNOVODSKI IZKAZI, NPV & PRODAJNA CENA EE, DA JE NPV 0
življenskadoba_database = []
inštaliranakapaciteta_database = []
capacityfactoravg_database = []
capacityfactor_neremont_database = []
capacityfactor_remont_database = []
OM20_database = []
OM40_database = []
OM60_database = []
OM80_database = []
OM100_database = []
OMvar_database = []
fuelcost_database = []
razgradnja_database = []
odpadki_database = []
occ_initial_database = []
occ_final_database = []
constructionperiod_database = []
costofequity_database = []
costofexim_database = []
costofloan_database = []
costofbond_database = []
wacc_database = []
političnotveganjegradnja_database = []
političnotveganjeobratovanje_database = []
tehničnotveganje_database = []
izberaniscenarijceneee_database = []
cenaelektrike_database = []
vseceneelektrike_database = []
energetskakriza_database = []
npv_database = []
prodajnacena_ee_npv_0_database = []
lastna_cena_database = []
lcoe_database = []

#dobiš časovno vrsto živlenjskih dob za vse iteracije
lifetime_years_data = sample_lifetime_years_with_weights(n_trials)

for lifetime_years in lifetime_years_data:
    
    #OSNOVNI PARAMETRI - inštalirana moč, življenska doba, gradbena doba, capacity factor
    
    #osnovna življenska doba pred vsemi tveganji, ki znižujejo življensko dobo
    lifetime_years_initial = round(lifetime_years)
    
    #lifetime_years_initial = 80
    
    #izguba politične podpore, ki pripelje do predčasnega zaprtja JE
    #result: 1 or 0
    lost_of_political_support_operation_happens = np.random.binomial(1, lost_of_political_support_operation_prob)
    
    if lost_of_political_support_operation_happens:
        #randomly selects a year between 1 and lifetime_years
        random_year_1 = np.random.randint(1, lifetime_years + 1)
        lifetime_years = random_year_1
    
    #tehnična okvara, ki pripelje do zaprta JE tekom obratovanja
    technical_reason_closure_happens = np.random.binomial(1, technical_reason_closure_prob)
    
    if technical_reason_closure_happens:
        random_year_2 = np.random.randint(1, lifetime_years + 1)
        lifetime_years = random_year_2
    
    #raje kot int() uporabim round, saj se zaokroži k najbližjemu polnemu številu (medtem ko int k spodnjemu polnemu številu)
    #za izračune
    #lifetime_years = 80
    lifetime_years = round(lifetime_years)
    
    #določi inštalirano moč
    installed_capacity = choose_installed_capacity(installed_capacities)

    #construction period
    #pridobi construction period
    construction_period = get_construction_period(kde_construction_period)
       
    construction_period = round(construction_period)

    #določi capacity factor in letno proizvedena EE
    capacity_factor_average, cf_ne_remont, cf_remont = calculate_capacity_factor(min_capacity_factor, mode_capacity_factor, max_capacity_factor)

    #letna proizvedena EE s CF_remont in CF_ne_remont
    yearly_ee_produced_mwh_remont = cf_remont * 8760 * installed_capacity
    yearly_ee_produced_mwh_ne_remont = cf_ne_remont * 8760 * installed_capacity

    
    ## IZKAZ POSLOVNEGA IZIDA
    
    
    #ODHODKI        
    #določi O&M costs
    var_OM_value = execute_triang_dist_OM_var(var_OM_value_mean, var_OM_value_min, var_OM_value_max)
    O_and_M_costs_construction = OMcosts_during_constructionperiod(construction_period, escalation_rate_O_and_M, installed_capacity)     
    total_OM_costs_20, O_and_M_20, O_and_M_costs_20, OM_var_costs_20 = calculate_O_and_M_over_20_years(installed_capacity, escalation_rate_O_and_M, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, first_period = 20)    
    total_OM_costs_40, O_and_M_40, O_and_M_costs_40, OM_var_costs_40 = calculate_O_and_M_over_40_years(installed_capacity, escalation_rate_O_and_M, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, second_period = 40)
    total_OM_costs_60, O_and_M_60, O_and_M_costs_60, OM_var_costs_60 = calculate_O_and_M_over_60_years(installed_capacity, escalation_rate_O_and_M, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, third_period = 60)
    total_OM_costs_80, O_and_M_80, O_and_M_costs_80, OM_var_costs_80 = calculate_O_and_M_over_80_years(installed_capacity, escalation_rate_O_and_M, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, fourth_period = 80)
    total_OM_costs_100, O_and_M_100, O_and_M_costs_100, OM_var_costs_100 = calculate_O_and_M_over_100_years(installed_capacity, escalation_rate_O_and_M, lifetime_years, var_OM_value, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, fifth_period = 100)

    OM_costs_over_lifetime = construct_OMcosts_over_lifetime(O_and_M_costs_construction, total_OM_costs_20, total_OM_costs_40, total_OM_costs_60, total_OM_costs_80, total_OM_costs_100)

    #določi fuel cost
    fuel_costs_over_lifetime, fuel_cost = calculate_fuel_costs_over_lifetime(construction_period, lifetime_years, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont)
    
    #razgradnja    
    decom_costs_over_lifetime, decom_cost_eur_mw = calculate_decom_costs(lifetime_years_initial, lifetime_years, construction_period, yearly_ee_produced_mwh_ne_remont, yearly_ee_produced_mwh_remont)
    
    #odlaganje odpadkov    
    waste_costs_over_lifetime, waste_cost_eur_mw = calculate_waste_costs(lifetime_years_initial, lifetime_years, construction_period, yearly_ee_produced_mwh_ne_remont, yearly_ee_produced_mwh_remont)

    #plačilo občini in vodarina
    payment_to_municipality_over_lifetime = calculate_payment_to_municipality(installed_capacity, lifetime_years, construction_period)

    #stroški nepričakovanih izpadov in popravil
    extraordinary_repair_costs_over_lifetime, scaled_unexp_outages = calculate_extraordinary_repairs_over_lifetime(lifetime_years, construction_period, average_unexp_outages_per_year)
    
    #stroški nadgradnje v primeru življenske dobe 80 let
    fullupgradecost80_over_lifetime = calculate_costs_upgrade_80(lifetime_years, installed_capacity, construction_period)

    #zavarovanje za primer jedrske nesreče, nuclear premium    
    insurance_premium_over_lifetime = calculate_insurance_premium(yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont, construction_period, lifetime_years)  
    
    #SKUPNI ODHODKI
    skupni_odhodki_po_letih = sum_columns_by_year(
        OM_costs_over_lifetime,
        fuel_costs_over_lifetime,
        decom_costs_over_lifetime,
        waste_costs_over_lifetime,
        payment_to_municipality_over_lifetime,
        extraordinary_repair_costs_over_lifetime,
        fullupgradecost80_over_lifetime,
        insurance_premium_over_lifetime        
    )
    #PRIHODKI
    #potrdila o izvoru prihodki
    prihodki_potrdilaoizvoru_over_lifetime = potrdilaoizvoru_prihodki_over_lifetime(construction_period, lifetime_years, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont)

    #elektrika prihodki
    ee_prihodki_over_lifetime, cene_elektrike_over_lifetime, chosen_scenario, ee_price, high_prices_happen = electricity_prihodki_over_lifetime(construction_period, lifetime_years, yearly_ee_produced_mwh_remont, yearly_ee_produced_mwh_ne_remont)

    #SKUPNI PRIHODKI
    skupni_prihodki_po_letih = sum_columns_by_year(
        prihodki_potrdilaoizvoru_over_lifetime,
        ee_prihodki_over_lifetime
    )

    #EBITDA
    #spremenit v array, saj če delaš razliko, gre vsak element posebej, element-wise operation
    skupni_odhodki_po_letih = np.array(skupni_odhodki_po_letih)
    skupni_prihodki_po_letih = np.array(skupni_prihodki_po_letih)
   
    #izračunat EBITDA
    ebitda = (skupni_prihodki_po_letih - skupni_odhodki_po_letih)
    
    #DEPRECIACIJA
    occ_final_depreciation_over_lifetime, occ_final, occ_initial, occ_final_eurkw = get_occ_initial_depreciation_occ_final_over_lifetime(
        student_t_params,
        construction_period,
        lifetime_years
    )
    
    occ_final_depreciation_over_lifetime = np.array(occ_final_depreciation_over_lifetime)
    
    #že tu izračunaš occ final cash flow, saj potrebuješ podatke pri obrestih
    investment_cash_flow, occ_final_cashflow_over_lifetime = get_occ_final_cashflow(occ_final, construction_period, lifetime_years)
    occ_final_cashflow_over_lifetime = np.array(occ_final_cashflow_over_lifetime)
    
    #EBIT 
    #izračunat EBIT
    ebit = ebitda - occ_final_depreciation_over_lifetime
    
    #OBRESTI
    #določi kdaj/katero leto se začne vsak od kreditov 
    equity_cash_flow, bond_cash_flow, exim_cash_flow, loan_cash_flow = define_cashflow_of_credits(investment_cash_flow, occ_final)

    #obresti treh virov financiranja - obveznice, exim kredit, bančni kredit
    bond_interest_rate, yearly_interest_bond, bond_interests_over_lifetime = calculate_cost_of_debt(share_cost_of_bond, occ_final, triang_dist_cost_of_bond, bond_repayment_period, bond_cash_flow, construction_period, lifetime_years)
    exim_interest_rate, yearly_interest_exim, exim_interests_over_lifetime = calculate_cost_of_debt(share_cost_of_exim, occ_final, triang_dist_cost_of_exim, exim_repayment_period, exim_cash_flow, construction_period, lifetime_years)
    loan_interest_rate, yearly_interest_loan, loan_interests_over_lifetime = calculate_cost_of_debt(share_cost_of_loan, occ_final, triang_dist_cost_of_loan, loan_repayment_period, loan_cash_flow, construction_period, lifetime_years)
    
    #obresti za celotno obdobje gradnje in življenske dobe elektrarne    
    #vsi dodatni krediti za pokrivanje obresti tekom gradnje
    all_additionalloans_interests_over_lifetime = calculate_additional_loan_during_construction(bond_interests_over_lifetime, exim_interests_over_lifetime, loan_interests_over_lifetime, construction_period, lifetime_years, loan_interest_rate, loan_repayment_period=20)

    #sešteješ dodatne kredite oz. obresti z istim pristopom kot zgoraj
    #* list v individualne argumente
    all_additionalloans_interests_sumed_over_lifetime = sum_columns_by_year(*all_additionalloans_interests_over_lifetime)

    #sešteješ obresti
    skupne_obresti = sum_columns_by_year(
        bond_interests_over_lifetime,
        exim_interests_over_lifetime,
        loan_interests_over_lifetime,
        all_additionalloans_interests_sumed_over_lifetime
    )

    #EBT
    ebt = ebit - skupne_obresti
    
    #DAVEK
    davek_over_lifetime = []
    for yearly_ebt in ebt:
        if yearly_ebt <= 0:
            davek_yearly = 0
        elif yearly_ebt > 0:
            davek_yearly = yearly_ebt * davek_na_dobiček
        davek_over_lifetime.append(davek_yearly)

    
    davek_over_lifetime = np.array(davek_over_lifetime)
    
    #ČISTI POSLOVNI IZID, neto prihodek, net income
    čisti_poslovni_izid = ebt - davek_over_lifetime

    
    ## DENARNI TOKOVI
    
    
    #PRILAGODITVE
        
    #obrestni odhodki, prilagojeni za davek (+)
    skupne_obresti_prilagojene_za_davek = skupne_obresti * (1 - davek_na_dobiček)

    #spremembe obratnega kapitala (-)
    #terjatve do kupcev - prilivi
    terjatve_do_kupcev = skupni_prihodki_po_letih * (dnevi_vezave_terjatev / 365)
    
    #vzpostaviš spremembe terjatve do kupcev spremenljvko
    #zanima te spremembe terjatve lasnko leto (ki jih nato dobiš v tem letu) minus terjatve v tem letu (ki še čakaš da jih dobiš)
    #vrstica samih 0, da lahko potem daješ notri vrednosti
    spr_terjatve_do_kupcev = np.zeros_like(terjatve_do_kupcev)
    
    #prva vrednost pri spremembi: 0 - prva vrednost terjatve do kupcev, saj delaš y-1 vrednost - y vrednost
    spr_terjatve_do_kupcev[0] = 0 - terjatve_do_kupcev[0]
    #spremembe terjatev do druge vrednosti naprej: y-1 vrednost - y vrednost
    spr_terjatve_do_kupcev[1:] = terjatve_do_kupcev[:-1] - terjatve_do_kupcev[1:]
    
    #obveznosti do dobaviteljev
    obveznosti_do_dobaviteljev = skupni_odhodki_po_letih * (dnevi_vezave_obveznosti / 365)
    
    spr_obveznosti_do_dobaviteljev = np.zeros_like(obveznosti_do_dobaviteljev)
    
    spr_obveznosti_do_dobaviteljev[0] = obveznosti_do_dobaviteljev[0] - 0
    
    spr_obveznosti_do_dobaviteljev[1:] = obveznosti_do_dobaviteljev[1:] - obveznosti_do_dobaviteljev[:-1]
    
    #sprememba obratnega kapitala kot seštevek terjatev do kupcev in obveznosti do dobaviteljev
    spr_obratnega_kapitala_po_letih = spr_obveznosti_do_dobaviteljev + spr_terjatve_do_kupcev
    
    #prištel boš depreciacijo (+)
    #odštel boš investicijski, occ final cash flow (-)
    
    #FCFF: prosti denarni tok podjetja, free cash flow to the firm
    fcff = (čisti_poslovni_izid +
        occ_final_depreciation_over_lifetime -
        occ_final_cashflow_over_lifetime +
        skupne_obresti_prilagojene_za_davek -
        spr_obratnega_kapitala_po_letih
    )
    
    #politično tveganje med gradnjo JE, ki prekine gradnjo
    lost_of_political_support_construction_happens = np.random.binomial(1, lost_of_political_support_construction_prob)
    
    if lost_of_political_support_construction_happens:
        #randomly izbereš leto med 1 and construction period
        random_year_3 = np.random.randint(1, construction_period + 1)
        fcff = fcff[:random_year_3]
 
    #WACC
    #določiš strošek lastniškega kapitala
    cost_of_equity = cost_of_equity_calculation(velika_vloga_drzave)

    #izračunaš wacc
    wacc = calculate_WACC(cost_of_equity, exim_interest_rate, loan_interest_rate, bond_interest_rate)

    #NPV NETO SEDANJA VREDNOST
    #excel npv: diskontiranje že prvo vrednost, python npv: prva vrednost ni diskontirana, zato zapišemo direktno formulo
    #https://stackoverflow.com/questions/62696450/is-there-a-difference-between-numpy-npv-and-excel-npv
    npv = (fcff / (1 + wacc) ** np.arange(1, len(fcff) + 1)).sum()
    #npv = npf.npv(wacc, fcff)
    
    #IZRAČUNAJ PRODAJNO CENO EE JEK2, DA PRIDE NPV = 0
    if not lost_of_political_support_construction_happens:
        electricity_price_npv_0 = find_electricity_price_npv0(
            construction_period=construction_period,
            lifetime_years=lifetime_years,
            yearly_ee_produced_mwh_remont=yearly_ee_produced_mwh_remont,
            yearly_ee_produced_mwh_ne_remont=yearly_ee_produced_mwh_ne_remont,
            skupni_odhodki_po_letih=skupni_odhodki_po_letih,
            prihodki_potrdilaoizvoru_over_lifetime=prihodki_potrdilaoizvoru_over_lifetime,
            occ_final_depreciation_over_lifetime=occ_final_depreciation_over_lifetime,
            occ_final_cashflow_over_lifetime=occ_final_cashflow_over_lifetime,
            skupne_obresti=skupne_obresti,
            davek_na_dobiček=davek_na_dobiček,
            dnevi_vezave_terjatev=dnevi_vezave_terjatev,
            dnevi_vezave_obveznosti=dnevi_vezave_obveznosti,
            wacc=wacc,
        )

        prodajnacena_ee_npv_0_database.append(electricity_price_npv_0)
    
    #IZRAČUNI LASTNA CENA

    vsa_proizvedena_ee_po_letih = []
    #0 med gradnjo
    vsa_proizvedena_ee_po_letih.extend([0] * construction_period)

    for year in range(lifetime_years):
        if (year % 3) == 0:
            vsa_proizvedena_ee_po_letih.append(yearly_ee_produced_mwh_ne_remont)
        else:
            vsa_proizvedena_ee_po_letih.append(yearly_ee_produced_mwh_remont)
    
    vsa_proizvedena_ee_po_letih = np.array(vsa_proizvedena_ee_po_letih)
        
    #IZRAČUNAJ LASTNO CENO EE - BREZ DISKONTIRANJA
    vsi_odhodki_po_letih = skupni_odhodki_po_letih + occ_final_cashflow_over_lifetime + skupne_obresti + davek_over_lifetime
    vsi_odhodki_živeljska_doba = np.sum(vsi_odhodki_po_letih)
    vsa_proizvedena_ee = np.sum(vsa_proizvedena_ee_po_letih)
    lastna_cena_ee = vsi_odhodki_živeljska_doba / vsa_proizvedena_ee
    lastna_cena_database.append(lastna_cena_ee)
    
    
    #IZRAČUNAJ LASTNO CENO EE - Z DISKONTIRANJEM
    vsi_odhodki_po_letih = skupni_odhodki_po_letih + occ_final_cashflow_over_lifetime + skupne_obresti + davek_over_lifetime
    letna_proizvedena_ee = capacity_factor_average * 8760 * installed_capacity
    diskontna_stopnja_dolg = exim_interest_rate * 0.333 + loan_interest_rate * 0.2 + bond_interest_rate * 0.467
    #discount factors per year
    discount_factors = np.array([(1 + diskontna_stopnja_dolg) ** (-t) for t in range(construction_period + lifetime_years)])
    #discounting
    diskontirani_odhodki = vsi_odhodki_po_letih * discount_factors
    diskontirana_proizvodnja = vsa_proizvedena_ee_po_letih * discount_factors
    lcoe = np.sum(diskontirani_odhodki) / np.sum(diskontirana_proizvodnja)
    lcoe_database.append(lcoe)
    
    
    #SHRANI PODATKE RAZLIČNINH SPREMENLJIVK V BAZO
    življenskadoba_database.append(lifetime_years_initial)
    inštaliranakapaciteta_database.append(installed_capacity)
    constructionperiod_database.append(construction_period)
    capacityfactoravg_database.append(capacity_factor_average)
    capacityfactor_neremont_database.append(cf_ne_remont)
    capacityfactor_remont_database.append(cf_remont)
    OMvar_database.append(var_OM_value)
    OM20_database.append(O_and_M_20)
    OM40_database.append(O_and_M_40)
    OM60_database.append(O_and_M_60)
    OM80_database.append(O_and_M_80)
    OM100_database.append(O_and_M_100)
    fuelcost_database.append(fuel_cost)
    razgradnja_database.append(decom_cost_eur_mw)
    odpadki_database.append(waste_cost_eur_mw)
    occ_initial_database.append(occ_initial)
    occ_final_database.append(occ_final_eurkw)
    costofequity_database.append(cost_of_equity)
    costofexim_database.append(exim_interest_rate)
    costofloan_database.append(loan_interest_rate)
    costofbond_database.append(bond_interest_rate)
    wacc_database.append(wacc)
    #shranjeno kot 0 ali 1
    političnotveganjegradnja_database.append(lost_of_political_support_construction_happens)
    političnotveganjeobratovanje_database.append(lost_of_political_support_operation_happens)
    tehničnotveganje_database.append(technical_reason_closure_happens)
    izberaniscenarijceneee_database.append(chosen_scenario) 
    #samo cene za osnovne tri scenarije
    cenaelektrike_database.append(ee_price)
    #list of lists
    vseceneelektrike_database.append(cene_elektrike_over_lifetime)
    energetskakriza_database.append(high_prices_happen)
    npv_database.append(npv)
    #shranjevanje cene EE da je NPV = 0 je že zgoraj


# %%NARIŠI GRAFE ZA PODATKE, UPORABLJENE V PYTHON MONTE CARLO ANALIZI, PO IZVEDENI ANALIZI

#za x vrstico
#"." -> "," 
#x: tick value
#_: unused second argument
def format_x_numbers(x, _):
    #American style
    #.1f: 1 decimal
    formatted = f"{x:,.1f}"
    #European style
    #, -> TEMP -> .
    #. -> ,
    formatted = formatted.replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")
    return formatted

# za y vrstico
def format_y_numbers(y, _):
    #American style
    formatted = f"{y:,.0f}"
    #European style
    formatted = formatted.replace(",", "TEMP").replace(".", ",").replace("TEMP", ".")
    return formatted


#NASLOVI GRAFOV
data_dict = {
    "Razporeditev življenske dobe JEK2": življenskadoba_database,
    "Inštalirana moč JEK2": inštaliranakapaciteta_database,
    "Dolžina gradnje JEK2": constructionperiod_database,
    "Razporeditev povprečnega faktorja obremenitve": capacityfactoravg_database,
    "Razporeditev faktorja obremenitve v letih brez remonta": capacityfactor_neremont_database,
    "Razporeditev faktorja obremenitve v letih remonta": capacityfactor_remont_database,
    "Razporeditev variabilnih stroškov obratovanja in vzdrževanja (O&M)": OMvar_database,
    "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 20. let": OM20_database,
    "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 21-40 let": OM40_database,
    "Razporeditev sroškov obratovanja in vzdrževanja (O&M) za obdboje 41-60 let": OM60_database,
    "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 61-80 let": OM80_database,
    "Razporeditev stroškov goriva": fuelcost_database,
    "Razporeditev stroškov razgradnje": razgradnja_database,
    "Razporeditev upravljanja z odpadki": odpadki_database,
    "Razporeditev začetnega OCC (OCC_initial)": occ_initial_database,
    "Razporeditev končnega OCC za JEK2 (OCC_final)": occ_final_database,
    "Razporeditev stroška lastniškega kapitala": costofequity_database,
    "Razporeditev stroška kredita izvozno-uvozne banke (EXIM)": costofexim_database,
    "Razporeditev stroška državno jamčenega kredita pri komercialnih bankah": costofloan_database,
    "Razporeditev stroška državnih obveznic": costofbond_database,
    "Razporeditev tehtanega povprečja stroškov kapitala (WACC)": wacc_database,
    "Razporeditev političnega tveganja tekom gradnje": političnotveganjegradnja_database,
    "Razporeditev političnega tveganja tekom obratovanja": političnotveganjeobratovanje_database,
    "Razporeditev tehnološkega tveganja": tehničnotveganje_database,
    "Razporeditev borznih cen električne energije osnovnih treh scenarijev": cenaelektrike_database,
    "Razporeditev cen električne energije znotraj treh osnovnih scenarijev in scenarija energetkse krize": vseceneelektrike_database,
    "Razporeditev pojavljanja energetske krize": energetskakriza_database,
    "Razporeditev neto sedanje vrednosti (NPV) projekta JEK2": npv_database,
    "Razporeditev izbranih scenarijev cen električne energije": izberaniscenarijceneee_database,
    "Razporeditev prodajne cene električne energije iz JEK2": prodajnacena_ee_npv_0_database,
    "Razporeditev prodajne cene električne energije iz JEK2 ter primerjava med povprečno prodajno in borzno ceno elektrike": prodajnacena_ee_npv_0_database,
    "Razporeditev lastne cene električne energije iz JEK2": lcoe_database,
    "Povprečna prodajna in lastna cena projekta JEK2 ter borzna cena električne energije": lcoe_database
}

#IMENA X OSI
#mapiranje data_dict to x-axis labels
x_axis_labels = {
    "Razporeditev življenske dobe JEK2": "Življenjska doba [leta]",
    "Inštalirana moč JEK2": "Inštalirana moč [MW]",
    "Dolžina gradnje JEK2": "Dolžina gradnje [leta]",
    "Razporeditev povprečnega faktorja obremenitve": "Povprečni faktor obremenitve [%]",
    "Razporeditev faktorja obremenitve v letih brez remonta": "Faktor obremenitve brez remonta [%]",
    "Razporeditev faktorja obremenitve v letih remonta": "Faktor obremenitve v letih remonta [%]",
    "Razporeditev variabilnih stroškov obratovanja in vzdrževanja (O&M)": "Variabilni stroški O&M [€/MWh]",
    "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 20. let": "Stroški O&M (20 let) [EUR/kW]",
    "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 21-40 let": "Stroški O&M (21–40 let) [EUR/kW]",
    "Razporeditev sroškov obratovanja in vzdrževanja (O&M) za obdboje 41-60 let": "Stroški O&M (41–60 let) [EUR/kW]",
    "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 61-80 let": "Stroški O&M (61–80 let) [EUR/kW]",
    "Razporeditev stroškov goriva": "Stroški goriva [EUR/MWh]",
    "Razporeditev stroškov razgradnje": "Stroški razgradnje [EUR/MW]",
    "Razporeditev upravljanja z odpadki": "Stroški upravljanja z odpadki [EUR/MW]",
    "Razporeditev začetnega OCC (OCC_initial)": "Začetni OCC [EUR/kW]",
    "Razporeditev končnega OCC (OCC_final)": "Končni OCC [EUR/kW]",
    "Razporeditev stroška lastniškega kapitala": "Strošek lastniškega kapitala [%]",
    "Razporeditev stroška kredita izvozno-uvozne banke (EXIM)": "Strošek kredita EXIM [%]",
    "Razporeditev stroška državno jamčenega kredita pri komercialnih bankah": "Strošek komercialnega kredita [%]",
    "Razporeditev stroška državnih obveznic": "Strošek državnih obveznic [%]",
    "Razporeditev tehtanega povprečja stroškov kapitala (WACC)": "WACC [%]",
    "Razporeditev borznih cen električne energije osnovnih treh scenarijev": "Cene električne energije [EUR/MWh]",
    "Razporeditev neto sedanje vrednosti (NPV) projekta JEK2": "Neto sedanja vrednost (NPV) [mrd. EUR]",
    "Razporeditev prodajne cene električne energije iz JEK2": "Prodajna cena električne energije [EUR/MWh]",
    "Razporeditev prodajne cene električne energije iz JEK2 ter primerjava med povprečno prodajno in borzno ceno elektrike": "Prodajna cena električne energije [EUR/MWh]",
    "Razporeditev lastne cene električne energije iz JEK2": "Lastna cena električne energije [EUR/MWh]"
}




for key, values in data_dict.items():
    if key not in active_plots:
        continue
    #plot
    plt.figure(figsize=(10, 6))
    
    # Specifični primer za list of lists
    if key == "Razporeditev cen električne energije znotraj treh osnovnih scenarijev in scenarija energetkse krize":

        #vsak price iz vsakega sublista v novi list flattened_prices
        flattened_prices = [price for sublist in values for price in sublist]
        flattened_prices = [price for price in flattened_prices if price != 0]
    
        #plot
        plt.hist(flattened_prices, bins=25, color='lightblue', edgecolor='black')
        plt.xlabel("Cena električne energije [EUR/MWh]")
        plt.ylabel("Število pojavitev")
        plt.title("Razporeditev cen električne energije znotraj treh osnovnih scenarijev in scenarija energetkse krize")
        
        #custom formatting
        plt.gca().xaxis.set_major_formatter(ticker.FuncFormatter(format_x_numbers))  # X-axis with decimals
        plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(format_y_numbers))  # Y-axis without decimals

        #adjust the spacing between elements of a plot to prevent overlapping
        plt.tight_layout()

    #specifičen primer
    elif key == "Razporeditev izbranih scenarijev cen električne energije":
        #outcome of Counter(values):
        #Keys: unique elements/categories
        #values: n of times vsak unique element se pokaže
        counts = Counter(values)
        desired_order = ["Scenarij nizkih cen", "Osnovni scenarij", "Scenarij visokih cen"]
        
        #nov dictionary z željeno vrstim redom (desired_order)
        ordered_counts = {key: counts[key] for key in desired_order if key in counts}
        
        plt.bar(ordered_counts.keys(), ordered_counts.values(), color='skyblue')
        plt.ylabel("Število pojavitev")
        plt.title("Razporeditev pojavljanja scenarijev cen električne energije")

        plt.tight_layout()

    #binary data (0 and 1)
    #v in [0, 1]: ali so vrednosti 0 ali 1
    elif all(v in [0, 1] for v in values):
        #kolikokrat 0 ali 1 se pojavijo
        counts = Counter(values)
        #counts je dictionary, dostopaš do value od key (ki je 0 oz 1)
        plt.bar(["Nepojavitev", "Pojavitev"], [counts[0], counts[1]], color='lightcoral')
        plt.ylabel("Število pojavitev")
        plt.title(f"{key}")
        plt.tight_layout()

    #spremenljivke, ki zahtevajo % na x osi
    elif key in [
        "Razporeditev tehtanega povprečja stroškov kapitala (WACC)",
        "Razporeditev stroška državnih obveznic",
        "Razporeditev stroška državno jamčenega kredita pri komercialnih bankah",
        "Razporeditev stroška kredita izvozno-uvozne banke (EXIM)",
        "Razporeditev stroška lastniškega kapitala",
        "Razporeditev povprečnega faktorja obremenitve",
        "Razporeditev faktorja obremenitve v letih brez remonta",
        "Razporeditev faktorja obremenitve v letih remonta",
    ]:
        #decimals -> percentages
        percentage_values = [value * 100 for value in values]
        plt.hist(percentage_values, bins=30, color='green', edgecolor='black')
        plt.xlabel(f"{x_axis_labels.get(key, key)}")
        plt.ylabel("Pogostost pojavljanja")
        plt.title(f"{key}")
        
        plt.gca().xaxis.set_major_formatter(ticker.FuncFormatter(format_x_numbers))
        plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(format_y_numbers))
        
        plt.tight_layout()

    #specifično za NPV
    elif key == "Razporeditev neto sedanje vrednosti (NPV) projekta JEK2":
        #v milijardo spremenit
        npv_values_in_billions = [value / 1e9 for value in values]
    
        #vrednosti % > 0
        percentage_above_zero = (sum(value > 0 for value in values) / len(values)) * 100
    
        plt.hist(npv_values_in_billions, bins=30, color='green', edgecolor='black', alpha=0.7)
    
        #vertical line pri 0
        plt.axvline(x=0, color='black', linestyle='--', linewidth=2, alpha=0.6)
        #dodaj tekst
        plt.text(0.5, plt.ylim()[1] * 0.9, f"Projekt je ekonomsko upravičen v {percentage_above_zero:.0f}. odstotkih", 
         fontsize=12, color='black')

        plt.xlabel("Neto sedanja vrednost (NPV) [milijarde EUR]")  # Adjusted x-axis label
        plt.ylabel("Pogostost pojavljanja")
        plt.title("Razporeditev neto sedanje vrednosti (NPV) projekta JEK2")
    
        plt.gca().xaxis.set_major_formatter(ticker.FuncFormatter(format_x_numbers))  # X-axis with decimals
        plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(format_y_numbers))  # Y-axis without decimals
    
        plt.tight_layout()

    #specifično za prodajno ceno EE


    elif key == "Razporeditev prodajne cene električne energije iz JEK2":
        electricity_values = np.array(values)
    
        #≤350 in preštej cifre >350
        capped_values = electricity_values[electricity_values <= 350]
        above_350_count = np.sum(electricity_values > 350)
    
        mean_price = np.mean(electricity_values)
    
        #fiksen bin širina
        bin_width = 10
        #regularni bini do 350
        bins = list(range(0, 351, bin_width))
        #bin 350+ ima enako širino
        bins.append(360)
    
        hist_values, bin_edges, _ = plt.hist(capped_values, bins=bins, color='green', edgecolor='black', alpha=0.7)
    
        #popravi višino za 350+ bin
        bin_area = above_350_count / (bin_edges[-1] - bin_edges[-2])
        plt.bar(355, bin_area, width=bin_width, color='green', edgecolor='black', alpha=0.7)
    
        #vertical mean price line
        plt.axvline(x=mean_price, color='black', linestyle='--', linewidth=2, alpha=0.6)
    
        #text
        plt.text(mean_price * 1.02, plt.ylim()[1] * 0.9, 
                 f"Povprečna prodajna cena JEK2: {mean_price:.1f}".replace('.', ',') + " EUR/MWh", 
                 fontsize=12, color='black')
    
        plt.xlabel(x_axis_labels.get(key, key))
        plt.ylabel("Pogostost pojavljanja")
        plt.title("Razporeditev prodajne cene električne energije iz JEK2")
    
        plt.gca().xaxis.set_major_formatter(ticker.FuncFormatter(format_x_numbers))
        plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(format_y_numbers))
    
        #remove 350 in ohrani le 350+ 
        tick_positions = [x for x in range(0, 351, 50) if x != 350] + [355]
        tick_labels = [str(x) for x in range(0, 351, 50) if x != 350] + ["350+"]
    
        plt.xticks(tick_positions, tick_labels)
    
        plt.tight_layout()


    #PRIMERJAVA PRODAJNA IN BORZNA CENA EE
    elif key == "Razporeditev prodajne cene električne energije iz JEK2 ter primerjava med povprečno prodajno in borzno ceno elektrike":
        electricity_values = np.array(values)
    
        #≤350 in preštej cifre >350
        capped_values = electricity_values[electricity_values <= 350]
        above_350_count = np.sum(electricity_values > 350)
    

        all_prices = np.concatenate(vseceneelektrike_database)   
        filtered_prices = all_prices[all_prices != 0]            
        mean_priceofee_4scenarios = np.mean(filtered_prices)
        mean_price = np.mean(electricity_values)
        mean_market_price = mean_priceofee_4scenarios
    
        #širina bin
        bin_width = 10
        bins = list(range(0, 351, bin_width))
        #zadnji bin za 350+ kot 350-360
        bins.append(360)
    
        hist_values, bin_edges, _ = plt.hist(
            capped_values, bins=bins, color='green', edgecolor='black', alpha=0.7
        )
    
        #add bar za 350+
        bin_area = above_350_count / (bin_edges[-1] - bin_edges[-2])
        plt.bar(355, bin_area, width=bin_width, color='green', edgecolor='black', alpha=0.7)
    
        #dodaj vertikalno linijo za mean JEK2 ceno
        plt.axvline(x=mean_price, color='black', linestyle='--', linewidth=2, alpha=0.8)
        plt.text(
            mean_price + 5, plt.ylim()[1] * 0.90,
            f"Povprečna prodajna cena JEK2: {mean_price:.1f}".replace('.', ',') + " EUR/MWh",
            color='black', fontsize=11
        )
    
        #dodaj vertikalno linijo za mean market price
        plt.axvline(x=mean_market_price, color='orange', linestyle='--', linewidth=2, alpha=0.8)
        plt.text(
            mean_market_price + 5, plt.ylim()[1] * 0.80,
            f"Povprečna borzna cena EE: {mean_market_price:.1f}".replace('.', ',') + " EUR/MWh",
            color='orange', fontsize=11
        )

        plt.xlabel(x_axis_labels.get(key, key))
        plt.ylabel("Pogostost pojavljanja")
        plt.title("Razporeditev prodajne cene električne energije iz JEK2 ter primerjava med povprečno prodajno in borzno ceno elektrike")
    
        plt.gca().xaxis.set_major_formatter(ticker.FuncFormatter(format_x_numbers))
        plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(format_y_numbers))
    
        #prikaži 350+    
        tick_positions = [x for x in range(0, 351, 50) if x != 350] + [355]
        tick_labels = [str(x) for x in range(0, 351, 50) if x != 350] + ["350+"]
        plt.xticks(tick_positions, tick_labels)
    
        plt.tight_layout()

    #LASTNA CENA LCOE
    
    elif key == "Razporeditev lastne cene električne energije iz JEK2":
        
        electricity_values = np.array(values)
    
        #<350 in count
        capped_values = electricity_values[electricity_values <= 350]
        above_350_count = np.sum(electricity_values > 350)
    
        mean_price = np.mean(electricity_values)
    
        bin_width = 10
        bins = list(range(0, 351, bin_width))
        bins.append(360)
    
        hist_values, bin_edges, _ = plt.hist(capped_values, bins=bins, color='green', edgecolor='black', alpha=0.7)
    
        #popravi višino 350+ - density upoštevaj 
        bin_area = above_350_count / (bin_edges[-1] - bin_edges[-2])
        plt.bar(355, bin_area, width=bin_width, color='green', edgecolor='black', alpha=0.7)
    
        #dodaj vertikalno mean line
        plt.axvline(x=mean_price, color='black', linestyle='--', linewidth=2, alpha=0.6)

        plt.text(mean_price * 1.02, plt.ylim()[1] * 0.9, 
                 f"Povprečna lastna cena JEK2: {mean_price:.1f}".replace('.', ',') + " EUR/MWh", 
                 fontsize=12, color='black')
    
        plt.xlabel(x_axis_labels.get(key, key))
        plt.ylabel("Pogostost pojavljanja")
        plt.title("Razporeditev lastne cene električne energije iz JEK2")
    
        plt.gca().xaxis.set_major_formatter(ticker.FuncFormatter(format_x_numbers))
        plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(format_y_numbers))
    
        #ne 350, ja 350+ 
        tick_positions = [x for x in range(0, 351, 50) if x != 350] + [355]
        tick_labels = [str(x) for x in range(0, 351, 50) if x != 350] + ["350+"]
    
        plt.xticks(tick_positions, tick_labels)
    
        plt.tight_layout()

    elif key == "Povprečna prodajna in lastna cena projekta JEK2 ter borzna cena električne energije":
    
        all_prices = np.concatenate(vseceneelektrike_database)
        filtered_prices = all_prices[all_prices != 0]
        mean_priceofee_4scenarios = np.mean(filtered_prices)
        mean_prodajnacena = np.mean(prodajnacena_ee_npv_0_database)
        mean_lcoe = np.mean(lcoe_database)

        values = [mean_prodajnacena, mean_priceofee_4scenarios, mean_lcoe]
        labels = [
            "Povprečna prodajna cena EE",
            "Povprečna borzna cena EE",
            "Povprečna lastna cena EE"
        ]
        colors = ['green', 'green', 'green']
    
        # Počeši figuro ki jo je zanka že ustvarila in jo resizaj na 8x6
        plt.gcf().set_size_inches(8, 6)
    
        bars = plt.bar(labels, values, color=colors, edgecolor='black', alpha=0.8)
    
       #value labels na vrh bar-ov
        for bar, value in zip(bars, values):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, height + 2,
                     f"{value:.1f}".replace('.', ',') + " EUR/MWh",
                     ha='center', va='bottom', fontsize=11)

        plt.ylabel("Cena električne energije [EUR/MWh]")
        plt.title("Povprečna prodajna in lastna cena projekta JEK2 ter borzna cena električne energije")
        
        #malo več space-a nad bari
        max_height = max(values)
        plt.ylim(0, max_height * 1.075)
        
        plt.grid(axis='y', linestyle='--', alpha=0.6)
        plt.tight_layout()

    #za numerical data
    else:
        #Odstrani 0 vrednosti iz O&M fiksnih stroškov, saj imajo 0, ko se JE predčasno zapre
        if key in [
            "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 20. let",
            "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 21-40 let",
            "Razporeditev sroškov obratovanja in vzdrževanja (O&M) za obdboje 41-60 let",
            "Razporeditev stroškov obratovanja in vzdrževanja (O&M) za obdobje 61-80 let",
        ]:
            values = [v for v in values if v != 0]
        
        #da dobiš vrednosti key z .get
        x_label = x_axis_labels.get(key, key)  
        plt.hist(values, bins=30, color='green', edgecolor='black')
        plt.xlabel(x_label)
        plt.ylabel("Pogostost pojavljanja")
        plt.title(f"{key}")

        plt.gca().xaxis.set_major_formatter(ticker.FuncFormatter(format_x_numbers))
        plt.gca().yaxis.set_major_formatter(ticker.FuncFormatter(format_y_numbers))

        plt.tight_layout()

# prikaže vse grafike skupaj
plt.show()