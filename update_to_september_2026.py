import openpyxl
import json
import os
import datetime

print("=== Starting update to September 2026 ===")

EXCEL_GLOBAL = 'excel_data/202609_Factory_Meeting_KPIs - Global.xlsx'
EXCEL_QHSE = 'excel_data/QHSE KPI 2026.xlsx'

wb_global = openpyxl.load_workbook(EXCEL_GLOBAL, data_only=True)
wb_qhse = openpyxl.load_workbook(EXCEL_QHSE, data_only=True)

def safe_num(v, mult=1.0):
    if v is None or v == '' or v == '#REF!' or v == '#VALUE!':
        return None
    try:
        val = float(str(v).replace('%', '').replace(',', '').strip())
        return round(val * mult, 4)
    except:
        return None

# ==============================================================================
# 1. UPDATE PROCESSUS PRODUCTION GLOBAL
# ==============================================================================
print("--> Processing Production...")
ws_eff = wb_global['Efficiency']
ws_qual = wb_global['Quality']

# In ws_eff:
# col 19: 2026-06, col 20: 2026-07, col 21: 2026-08, col 22: 2026-09, col 23: YTD, col 24: TARGET
eff_bu_rows = {
    'SAFETY': 4,
    'LEATHER': 12,
    'D|AIR': 20,
    'MEL-T': 28
}

prod_data_path = 'data/processus_production_global.json'
with open(prod_data_path, 'r', encoding='utf-8') as f:
    prod_json = json.load(f)

# Update efficiency, productivity, volume for each BU
for bu, r in eff_bu_rows.items():
    if bu not in prod_json['sites']:
        continue
    
    vol_row = r + 1
    std_row = r + 2
    wkd_row = r + 3
    eff_mth_row = r + 4
    eff_ytd_row = r + 5
    prod_row = r + 6
    
    # Values
    eff_vals = {
        'juin_2026': safe_num(ws_eff.cell(eff_mth_row, 19).value, 100.0),
        'juillet_2026': safe_num(ws_eff.cell(eff_mth_row, 20).value, 100.0),
        'aout_2026': safe_num(ws_eff.cell(eff_mth_row, 21).value, 100.0),
        'septembre_2026': safe_num(ws_eff.cell(eff_mth_row, 22).value, 100.0),
        'ytd_2026': safe_num(ws_eff.cell(eff_ytd_row, 23).value, 100.0),
        'cible': safe_num(ws_eff.cell(eff_ytd_row, 24).value, 100.0)
    }
    
    prod_vals = {
        'juin_2026': safe_num(ws_eff.cell(prod_row, 19).value),
        'juillet_2026': safe_num(ws_eff.cell(prod_row, 20).value),
        'aout_2026': safe_num(ws_eff.cell(prod_row, 21).value),
        'septembre_2026': safe_num(ws_eff.cell(prod_row, 22).value),
        'ytd_2026': safe_num(ws_eff.cell(prod_row, 23).value),
        'cible': safe_num(ws_eff.cell(prod_row, 24).value)
    }
    
    vol_vals = {
        'juin_2026': safe_num(ws_eff.cell(vol_row, 19).value),
        'juillet_2026': safe_num(ws_eff.cell(vol_row, 20).value),
        'aout_2026': safe_num(ws_eff.cell(vol_row, 21).value),
        'septembre_2026': safe_num(ws_eff.cell(vol_row, 22).value),
        'ytd_2026': safe_num(ws_eff.cell(vol_row, 23).value),
        'cible': None
    }
    
    for item in prod_json['sites'][bu]:
        if item['nom'] == 'Efficacité opérationnelle':
            item['valeurs'].update(eff_vals)
        elif item['nom'] == 'Productivité en pièces/heure':
            item['valeurs'].update(prod_vals)
        elif item['nom'] == 'Volume de production':
            item['valeurs'].update(vol_vals)

# Quality (FQC / OQC) in ws_qual
# Rows:
# FQC trend: row 14 Safety, 15 Leather, 17 D|air, 18 MEL-T
# col 11: Jun, 12: Jul, 13: Aug, 14: Sep
# FQC recap: row 6 Safety, 7 Leather, 9 D|air, 10 MEL-T. col 6 is Rej %, col 7 Target, col 14 YTD Rej %
qual_fqc_trend = {'SAFETY': 14, 'LEATHER': 15, 'D|AIR': 17, 'MEL-T': 18}
qual_fqc_recap = {'SAFETY': 6, 'LEATHER': 7, 'D|AIR': 9, 'MEL-T': 10}

qual_oqc_trend = {'SAFETY': 34, 'LEATHER': 35, 'D|AIR': 37, 'MEL-T': 38}
qual_oqc_recap = {'SAFETY': 26, 'LEATHER': 27, 'D|AIR': 29, 'MEL-T': 30}

for bu in ['SAFETY', 'LEATHER', 'D|AIR', 'MEL-T']:
    if bu not in prod_json['sites']:
        continue
    tr_fqc = qual_fqc_trend[bu]
    rc_fqc = qual_fqc_recap[bu]
    
    fqc_vals = {
        'juin_2026': safe_num(ws_qual.cell(tr_fqc, 11).value, 100.0),
        'juillet_2026': safe_num(ws_qual.cell(tr_fqc, 12).value, 100.0),
        'aout_2026': safe_num(ws_qual.cell(tr_fqc, 13).value, 100.0),
        'septembre_2026': safe_num(ws_qual.cell(tr_fqc, 14).value, 100.0),
        'ytd_2026': safe_num(ws_qual.cell(rc_fqc, 14).value, 100.0),
        'cible': safe_num(ws_qual.cell(rc_fqc, 7).value, 100.0)
    }
    
    tr_oqc = qual_oqc_trend[bu]
    rc_oqc = qual_oqc_recap[bu]
    oqc_vals = {
        'juin_2026': safe_num(ws_qual.cell(tr_oqc, 11).value, 100.0),
        'juillet_2026': safe_num(ws_qual.cell(tr_oqc, 12).value, 100.0),
        'aout_2026': safe_num(ws_qual.cell(tr_oqc, 13).value, 100.0),
        'septembre_2026': safe_num(ws_qual.cell(tr_oqc, 14).value, 100.0),
        'ytd_2026': safe_num(ws_qual.cell(rc_oqc, 14).value, 100.0),
        'cible': safe_num(ws_qual.cell(rc_oqc, 7).value, 100.0)
    }
    
    for item in prod_json['sites'][bu]:
        if item['nom'] == 'Taux de défaut FQC':
            item['valeurs'].update(fqc_vals)
        elif item['nom'] == 'Taux de défaut OQC':
            item['valeurs'].update(oqc_vals)

with open(prod_data_path, 'w', encoding='utf-8') as f:
    json.dump(prod_json, f, indent=2, ensure_ascii=False)
print("--> Updated", prod_data_path)


# ==============================================================================
# 2. UPDATE PROCESSUS FINANCE GLOBAL
# ==============================================================================
print("--> Processing Finance...")
ws_costs = wb_global['Costs & Profitability']
fin_data_path = 'data/processus_finance_global.json'
with open(fin_data_path, 'r', encoding='utf-8') as f:
    fin_json = json.load(f)

# Scrap: row 5 Safety, 6 Leather, 7 D|air, 8 MEL-T
# col 15: Jun, 16: Jul, 17: Aug, 18: Sep, 19: YTD
scrap_rows = {'SAFETY': 5, 'LEATHER': 6, 'D|AIR': 7, 'MEL-T': 8}

for bu, r in scrap_rows.items():
    if bu not in fin_json['sites']:
        continue
    scrap_vals = {
        'juin_2026': safe_num(ws_costs.cell(r, 15).value),
        'juillet_2026': safe_num(ws_costs.cell(r, 16).value),
        'aout_2026': safe_num(ws_costs.cell(r, 17).value),
        'septembre_2026': safe_num(ws_costs.cell(r, 18).value),
        'ytd_2026': safe_num(ws_costs.cell(r, 19).value)
    }
    for item in fin_json['sites'][bu]:
        if item['nom'] == 'Valeur des rebuts matière':
            item['valeurs'].update(scrap_vals)

# Maintenance cost: row 18 Total cost (TND)
# Maintenance cost/hour DL: row 23
maint_total_vals = {
    'juin_2026': safe_num(ws_costs.cell(18, 15).value),
    'juillet_2026': safe_num(ws_costs.cell(18, 16).value),
    'aout_2026': safe_num(ws_costs.cell(18, 17).value),
    'septembre_2026': safe_num(ws_costs.cell(18, 18).value),
    'ytd_2026': safe_num(ws_costs.cell(18, 19).value),
    'cible': safe_num(ws_costs.cell(18, 20).value)
}
maint_dl_vals = {
    'juin_2026': safe_num(ws_costs.cell(23, 15).value),
    'juillet_2026': safe_num(ws_costs.cell(23, 16).value),
    'aout_2026': safe_num(ws_costs.cell(23, 17).value),
    'septembre_2026': safe_num(ws_costs.cell(23, 18).value),
    'ytd_2026': safe_num(ws_costs.cell(23, 19).value),
    'cible': safe_num(ws_costs.cell(23, 20).value)
}

for bu in ['MEL-T']:
    for item in fin_json['sites'][bu]:
        if item['nom'] == 'Coût total de maintenance':
            item['valeurs'].update(maint_total_vals)
        elif item['nom'] == 'Coût de maintenance par heure directe':
            item['valeurs'].update(maint_dl_vals)

with open(fin_data_path, 'w', encoding='utf-8') as f:
    json.dump(fin_json, f, indent=2, ensure_ascii=False)
print("--> Updated", fin_data_path)


# ==============================================================================
# 3. UPDATE PROCESSUS RH GLOBAL
# ==============================================================================
print("--> Processing RH...")
ws_people = wb_global['People']
rh_data_path = 'data/processus_rh_global.json'
with open(rh_data_path, 'r', encoding='utf-8') as f:
    rh_json = json.load(f)

# In ws_people:
# col 14: Jun, 15: Jul, 16: Aug, 17: Sep, 18: Oct (target/budget)
# Rows:
# Safety: Beginning HC row 5, Turnover row 9, Turnover YTD row 10
# Leather: Beginning HC row 11, Turnover row 15, Turnover YTD row 16
# D|air: Turnover row 21
# MEL-T: Beginning HC row 29, Turnover row 33, Turnover YTD row 34
rh_hc_rows = {'SAFETY': 5, 'LEATHER': 11, 'MEL-T': 29}
rh_turnover_rows = {'SAFETY': 9, 'LEATHER': 15, 'D|AIR': 21, 'MEL-T': 33}
rh_turnover_ytd = {'SAFETY': 10, 'LEATHER': 16, 'MEL-T': 34}

for bu, r in rh_hc_rows.items():
    if bu not in rh_json['sites']: continue
    hc_vals = {
        'juin_2026': safe_num(ws_people.cell(r, 14).value),
        'juillet_2026': safe_num(ws_people.cell(r, 15).value),
        'aout_2026': safe_num(ws_people.cell(r, 16).value),
        'septembre_2026': safe_num(ws_people.cell(r, 17).value)
    }
    for item in rh_json['sites'][bu]:
        if item['nom'] == 'Effectif total':
            item['valeurs'].update(hc_vals)

for bu, r in rh_turnover_rows.items():
    if bu not in rh_json['sites']: continue
    ytd_r = rh_turnover_ytd.get(bu)
    to_vals = {
        'juin_2026': safe_num(ws_people.cell(r, 14).value, 100.0),
        'juillet_2026': safe_num(ws_people.cell(r, 15).value, 100.0),
        'aout_2026': safe_num(ws_people.cell(r, 16).value, 100.0),
        'septembre_2026': safe_num(ws_people.cell(r, 17).value, 100.0),
        'ytd_2026': safe_num(ws_people.cell(ytd_r, 17).value, 100.0) if ytd_r else None
    }
    for item in rh_json['sites'][bu]:
        if item['nom'] == 'Taux de turnover':
            item['valeurs'].update(to_vals)

with open(rh_data_path, 'w', encoding='utf-8') as f:
    json.dump(rh_json, f, indent=2, ensure_ascii=False)
print("--> Updated", rh_data_path)


# ==============================================================================
# 4. UPDATE PROCESSUS LOGISTIQUE GLOBAL (Warehouse POM & Inventory Accuracy)
# ==============================================================================
print("--> Processing Logistics Warehouse...")
ws_wh = wb_global['Warehouse & Inventory Manag']
log_data_path = 'data/processus_logistique_global.json'
with open(log_data_path, 'r', encoding='utf-8') as f:
    log_json = json.load(f)

# POM rows: Safety row 6, Leather row 7, D|air row 8, MEL-T row 9
# col 14: Jun, 15: Jul, 16: Aug, 17: Sep, 18: YTD
pom_rows = {'SAFETY': 6, 'LEATHER': 7, 'D|AIR': 8, 'MEL-T': 9}
for bu, r in pom_rows.items():
    if bu not in log_json['sites']: continue
    pom_vals = {
        'juin_2026': safe_num(ws_wh.cell(r, 14).value, 100.0),
        'juillet_2026': safe_num(ws_wh.cell(r, 15).value, 100.0),
        'aout_2026': safe_num(ws_wh.cell(r, 16).value, 100.0),
        'septembre_2026': safe_num(ws_wh.cell(r, 17).value, 100.0),
        'ytd_2026': safe_num(ws_wh.cell(r, 18).value, 100.0)
    }
    for item in log_json['sites'][bu]:
        if item['nom'] == 'Fiabilité des livraisons':
            item['valeurs'].update(pom_vals)

# Inventory Accuracy Rate: row 19, YTD is col 18 = 0.65 (65%)
inv_acc_vals = {
    'ytd_2026': safe_num(ws_wh.cell(19, 18).value, 100.0)
}
for bu in log_json['sites']:
    for item in log_json['sites'][bu]:
        if item['nom'] == 'Précision des stocks':
            item['valeurs'].update(inv_acc_vals)

with open(log_data_path, 'w', encoding='utf-8') as f:
    json.dump(log_json, f, indent=2, ensure_ascii=False)
print("--> Updated", log_data_path)


# ==============================================================================
# 5. UPDATE KPI_METRICS.JSON (QHSE SST & ENERGY)
# ==============================================================================
print("--> Processing QHSE Metrics...")
metrics_path = 'data/kpi_metrics.json'
with open(metrics_path, 'r', encoding='utf-8') as f:
    metrics = json.load(f)

ws_sst = wb_qhse['KPI - SST']
# Cols: 37: 2026-06, 38: 2026-07, 39: 2026-08, 40: 2026-09
# Sites rows in SST:
# Safety: row 14 hours, 15 accidents, 16 lost-time, 17 lost-days, 18 TF, 19 TG
# Leather: row 20 hours, 21 accidents, 22 lost-time, 23 lost-days, 24 TF, 25 TG
# D|air: row 26 hours, 27 accidents, 28 lost-time, 29 lost-days, 30 TF, 31 TG
# Warehouse: row 32 hours, 33 accidents, 34 lost-time, 35 lost-days, 36 TF, 37 TG
# MEL-T: row 38 hours, 39 accidents, 40 lost-time, 41 lost-days, 42 TF, 43 TG

sst_site_blocks = {
    'Safety': 14,
    'Leather': 20,
    'D|AIR': 26,
    'Warehouse': 32,
    'MEL-T': 39
}

sst_inds = [
    (1, 'Accidents du travail'),
    (2, 'Accidents avec arrêt'),
    (3, 'Jours perdus'),
    (4, 'Taux de fréquence (TF)'),
    (5, 'Taux de gravité (TG)')
]

new_entries = []
q3_periods = [
    (38, '2026-07'),
    (39, '2026-08'),
    (40, '2026-09')
]

for site, start_r in sst_site_blocks.items():
    for offset, ind_name in sst_inds:
        r = start_r + offset
        for col, period_key in q3_periods:
            v = safe_num(ws_sst.cell(r, col).value)
            if v is not None:
                new_entries.append({
                    "section": "s1_kpi_sst__sant___s_curit__au_travail",
                    "site": site.lower(),
                    "site_label": site,
                    "indicateur": ind_name,
                    "unite": None,
                    "periode": period_key,
                    "valeur": v
                })

# Energy in ws_nrg
ws_nrg = wb_qhse['KPI ENERGY']
# Cols: 38: 2026-07, 39: 2026-08, 40: 2026-09
# Safety 27B & Leather 27A:
# row 14: KWh, row 15: TEP, row 16: TCO2, row 19: KWh/piece
# D|air LOT 5:
# row 20: KWh, row 21: TEP, row 24: KWh/piece
# MEL-T:
# row 29: KWh, row 30: TEP, row 31: CO2, row 33: KWh/piece
# Diesel MEL-T:
# row 48: Litre, row 49: TEP, row 50: tCO2
# Bilan global:
# row 61: MEL-T TEP total, row 63: MEL-T CO2 total

for col, period_key in q3_periods:
    # Safety & Leather
    kwh_sl = safe_num(ws_nrg.cell(14, col).value)
    co2_sl = safe_num(ws_nrg.cell(16, col).value)
    ie_sl = safe_num(ws_nrg.cell(19, col).value)
    for s in ['Safety', 'Leather']:
        if kwh_sl is not None:
            new_entries.append({
                "section": "s2_kpi_energy___consommation__lectrique",
                "site": s.lower(),
                "site_label": s,
                "indicateur": "Consommation d'électricité",
                "unite": "kWh",
                "periode": period_key,
                "valeur": kwh_sl
            })
        if co2_sl is not None:
            new_entries.append({
                "section": "s2_kpi_energy___consommation__lectrique",
                "site": s.lower(),
                "site_label": s,
                "indicateur": "Émission d'électricité",
                "unite": "tCO₂",
                "periode": period_key,
                "valeur": co2_sl
            })
        if ie_sl is not None:
            new_entries.append({
                "section": "s2_kpi_energy___consommation__lectrique",
                "site": s.lower(),
                "site_label": s,
                "indicateur": "Intensité électrique (IE)",
                "unite": "kWh/pièce",
                "periode": period_key,
                "valeur": ie_sl
            })
            
    # D|air
    kwh_dair = safe_num(ws_nrg.cell(20, col).value)
    ie_dair = safe_num(ws_nrg.cell(24, col).value)
    if kwh_dair is not None:
        new_entries.append({
            "section": "s2_kpi_energy___consommation__lectrique",
            "site": "dair",
            "site_label": "D|AIR",
            "indicateur": "Consommation d'électricité",
            "unite": "kWh",
            "periode": period_key,
            "valeur": kwh_dair
        })
    if ie_dair is not None:
        new_entries.append({
            "section": "s2_kpi_energy___consommation__lectrique",
            "site": "dair",
            "site_label": "D|AIR",
            "indicateur": "Intensité électrique (IE)",
            "unite": "kWh/pièce",
            "periode": period_key,
            "valeur": ie_dair
        })
        
    # MEL-T
    kwh_melt = safe_num(ws_nrg.cell(29, col).value)
    tep_melt = safe_num(ws_nrg.cell(61, col).value)
    co2_melt = safe_num(ws_nrg.cell(63, col).value)
    diesel_melt = safe_num(ws_nrg.cell(48, col).value)
    ie_melt = safe_num(ws_nrg.cell(33, col).value)
    if kwh_melt is not None:
        new_entries.append({
            "section": "s2_kpi_energy___consommation__lectrique",
            "site": "melt",
            "site_label": "MEL-T",
            "indicateur": "Consommation d'électricité",
            "unite": "kWh",
            "periode": period_key,
            "valeur": kwh_melt
        })
    if tep_melt is not None:
        new_entries.append({
            "section": "s4_kpi_energy___bilan__nerg_tique_global",
            "site": "melt",
            "site_label": "MEL-T",
            "indicateur": "TEP",
            "unite": "TEP",
            "periode": period_key,
            "valeur": tep_melt
        })
    if co2_melt is not None:
        new_entries.append({
            "section": "s2_kpi_energy___consommation__lectrique",
            "site": "melt",
            "site_label": "MEL-T",
            "indicateur": "Émissions totales CO₂",
            "unite": "tCO₂",
            "periode": period_key,
            "valeur": co2_melt
        })
    if diesel_melt is not None:
        new_entries.append({
            "section": "s3_kpi_energy___consommation_gasoil",
            "site": "melt",
            "site_label": "MEL-T",
            "indicateur": "Consommation Gasoil sans soufre",
            "unite": "L",
            "periode": period_key,
            "valeur": diesel_melt
        })
    if ie_melt is not None:
        new_entries.append({
            "section": "s2_kpi_energy___consommation__lectrique",
            "site": "melt",
            "site_label": "MEL-T",
            "indicateur": "Intensité électrique Globale",
            "unite": "kWh/pièce",
            "periode": period_key,
            "valeur": ie_melt
        })

# Remove any previous Q3 entries to guarantee clean data
metrics = [m for m in metrics if m.get('periode') not in ['2026-07', '2026-08', '2026-09']]
metrics.extend(new_entries)
added_count = len(new_entries)

with open(metrics_path, 'w', encoding='utf-8') as f:
    json.dump(metrics, f, indent=2, ensure_ascii=False)
print(f"--> Added {added_count} new entries to {metrics_path}")

wb_global.close()
wb_qhse.close()
print("=== Extraction completed successfully ===")
