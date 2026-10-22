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

# Ph_QC in ws_phqc
ws_phqc = wb_global['Ph_QC']
phqc_bu_rows = {
    'SAFETY': {'prod_fqc': 6, 'throughput_fqc': 26, 'qc_pcs': 4, 'worked_hrs': 5},
    'LEATHER': {'prod_fqc': 10, 'throughput_fqc': 30, 'qc_pcs': 8, 'worked_hrs': 9},
    'MEL-T': {'prod_fqc': 18, 'throughput_fqc': 38, 'qc_pcs': 16, 'worked_hrs': 17}
}

empty_phqc_names = {'Productivité FQC', 'Throughput FQC', 'Pièces contrôlées FQC', 'Heures travaillées FQC'}
if 'D|AIR' in prod_json['sites']:
    prod_json['sites']['D|AIR'] = [item for item in prod_json['sites']['D|AIR'] if item.get('nom') not in empty_phqc_names]

for bu, cfg in phqc_bu_rows.items():
    if bu not in prod_json['sites']:
        continue

    def extract_phqc_vals(r):
        if not r:
            return {}
        # In Ph_QC: col 20: 2026-06, col 21: 2026-07, col 22: 2026-08, col 23: 2026-09, col 24: YTD, col 25: TARGET
        return {
            'juin_2026': safe_num(ws_phqc.cell(r, 20).value),
            'juillet_2026': safe_num(ws_phqc.cell(r, 21).value),
            'aout_2026': safe_num(ws_phqc.cell(r, 22).value),
            'septembre_2026': safe_num(ws_phqc.cell(r, 23).value),
            '2026-06': safe_num(ws_phqc.cell(r, 20).value),
            '2026-07': safe_num(ws_phqc.cell(r, 21).value),
            '2026-08': safe_num(ws_phqc.cell(r, 22).value),
            '2026-09': safe_num(ws_phqc.cell(r, 23).value),
            'ytd_2026': safe_num(ws_phqc.cell(r, 24).value),
            'cible': safe_num(ws_phqc.cell(r, 25).value)
        }

    kpi_defs = [
        {
            'nom': 'Productivité FQC',
            'code': 'Productivity_QC_Ph',
            'description': 'Productivité du contrôle qualité final (pièces/heure)',
            'formule': 'QC_Pcs / Worked Hrs',
            'unite': 'pièces / heure',
            'valeurs': extract_phqc_vals(cfg['prod_fqc'])
        },
        {
            'nom': 'Throughput FQC',
            'code': 'Throughput (ph)_QC',
            'description': 'Débit de contrôle final rapporté aux heures 1.0 FTE',
            'formule': 'QC_Pcs / Run_hrs (1.0 FTE)',
            'unite': 'pièces / heure',
            'valeurs': extract_phqc_vals(cfg['throughput_fqc'])
        },
        {
            'nom': 'Pièces contrôlées FQC',
            'code': 'QC_Pcs',
            'description': 'Quantité totale de pièces inspectées en contrôle final',
            'formule': 'Somme des pièces présentées au FQC',
            'unite': 'unités',
            'valeurs': extract_phqc_vals(cfg['qc_pcs'])
        },
        {
            'nom': 'Heures travaillées FQC',
            'code': 'Worked Hrs (QC)',
            'description': 'Heures de présence effectives dédiées au contrôle qualité final',
            'formule': 'Total des heures de contrôle final',
            'unite': 'heures',
            'valeurs': extract_phqc_vals(cfg['worked_hrs'])
        }
    ]
    for kdef in kpi_defs:
        existing = next((x for x in prod_json['sites'][bu] if x['nom'] == kdef['nom']), None)
        if existing:
            existing['valeurs'].update(kdef['valeurs'])
        else:
            prod_json['sites'][bu].append(kdef)

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
maint_overall_vals = {
    'juin_2026': safe_num(ws_costs.cell(24, 15).value),
    'juillet_2026': safe_num(ws_costs.cell(24, 16).value),
    'aout_2026': safe_num(ws_costs.cell(24, 17).value),
    'septembre_2026': safe_num(ws_costs.cell(24, 18).value),
    'ytd_2026': safe_num(ws_costs.cell(24, 19).value),
    'cible': safe_num(ws_costs.cell(24, 20).value)
}
maint_mch_vals = {
    'juin_2026': safe_num(ws_costs.cell(25, 15).value),
    'juillet_2026': safe_num(ws_costs.cell(25, 16).value),
    'aout_2026': safe_num(ws_costs.cell(25, 17).value),
    'septembre_2026': safe_num(ws_costs.cell(25, 18).value),
    'ytd_2026': safe_num(ws_costs.cell(25, 19).value),
    'cible': safe_num(ws_costs.cell(25, 20).value)
}

maintenance_items = {
    'Coût total de maintenance': ('Total Maintenance Cost', 'TND', maint_total_vals),
    'Coût de maintenance par heure directe': ('Maintenance Cost/h DL', 'TND/h', maint_dl_vals),
    'Coût de maintenance par heure globale': ('Maintenance Cost/h overall', 'TND/h', maint_overall_vals),
    'Coût de maintenance par heure MCH': ('Maintenance Cost/h (MCH)', 'TND/h', maint_mch_vals),
}
for name, (code, unit, values) in maintenance_items.items():
    existing = next((item for item in fin_json['sites']['MEL-T'] if item['nom'] == name), None)
    if existing:
        existing['valeurs'] = values
    else:
        fin_json['sites']['MEL-T'].append({
            'nom': name,
            'code': code,
            'description': code,
            'formule': 'Valeur extraite de Costs & Profitability',
            'unite': unit,
            'valeurs': values,
        })

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
# 4. UPDATE PROCESSUS PLANNING & DELIVERY GLOBAL
# ==============================================================================
print("--> Processing Delivery and Warehouse...")
ws_delivery = wb_global['Delivery']
ws_wh = wb_global['Warehouse & Inventory Manag']
log_data_path = 'data/processus_logistique_global.json'
with open(log_data_path, 'r', encoding='utf-8') as f:
    log_json = json.load(f)

log_json['processus'] = 'Planning & Delivery'
log_json['description'] = 'Indicateurs opérationnels issus de Delivery, complétés par POM et Inventory Accuracy depuis Warehouse & Inventory Manag'
log_json['notes'] = [
    'Source opérationnelle: 202609_Factory_Meeting_KPIs - Global.xlsx, feuille Delivery.',
    'Les données Delivery 2026 s’arrêtent à juin; le YTD reste une période distincte et ne remplace jamais septembre.',
    'POM et Inventory Accuracy proviennent uniquement de Warehouse & Inventory Manag.',
    'La feuille Logistic n’est pas utilisée.'
]

months_2026 = [
    ('janvier_2026', 15), ('fevrier_2026', 16), ('mars_2026', 17),
    ('avril_2026', 18), ('mai_2026', 19), ('juin_2026', 20),
]
warehouse_months_2026 = [
    ('janvier_2026', 9), ('fevrier_2026', 10), ('mars_2026', 11),
    ('avril_2026', 12), ('mai_2026', 13), ('juin_2026', 14),
    ('juillet_2026', 15), ('aout_2026', 16), ('septembre_2026', 17),
]

def row_values(ws, row, columns, multiplier=1.0, ytd_col=None, target_col=None):
    values = {}
    for key, col in columns:
        value = safe_num(ws.cell(row, col).value, multiplier)
        if value is not None:
            values[key] = value
    if ytd_col:
        value = safe_num(ws.cell(row, ytd_col).value, multiplier)
        if value is not None:
            values['ytd_2026'] = value
    if target_col:
        value = safe_num(ws.cell(row, target_col).value, multiplier)
        if value is not None:
            values['cible'] = value
    return values

def item(name, code, description, formula, unit, values):
    return {
        'nom': name, 'code': code, 'description': description,
        'formule': formula, 'unite': unit, 'valeurs': values,
    }

delivery_rows = {
    'SAFETY': {'otif': 5, 'mo': 6, 'lt': 21, 'service': 40, 'sold': 51, 'defective': 52, 'crr': 53, 'claims': 54},
    'LEATHER': {'otif': 7, 'mo': 8, 'lt': 24, 'service': 41, 'sold': 55, 'defective': 56, 'crr': 57, 'claims': 58},
    'D|AIR': {'otif': 9, 'mo': 10, 'lt': 27, 'service': 42, 'sold': 59, 'defective': 60, 'crr': 61, 'claims': 62},
    'MEL-T': {'otif': 11, 'mo': 12, 'lt': 30, 'service': 43, 'sold': 63, 'defective': 64, 'crr': 65, 'claims': 68, 'cqr': 66, 'dpmu': 67},
}
# Discard legacy HELMET/Logistic records: Delivery contains only these four reporting scopes.
log_json['sites'] = {bu: [] for bu in delivery_rows}
pom_rows = {'SAFETY': 6, 'LEATHER': 7, 'D|AIR': 8, 'MEL-T': 9}

for bu, rows in delivery_rows.items():
    otif = row_values(ws_delivery, rows['otif'], months_2026, 100.0, 21, 22)
    mo = row_values(ws_delivery, rows['mo'], months_2026, 1.0, 21)
    anticipo = row_values(ws_delivery, rows['lt'], months_2026, 1.0, 21, 22)
    production_lt = row_values(ws_delivery, rows['lt'] + 1, months_2026, 1.0, 21, 22)
    end_to_end = row_values(ws_delivery, rows['lt'] + 2, months_2026, 1.0, 21, 22)
    units_shipped = row_values(ws_delivery, rows['service'], months_2026, 1.0, 21)
    service_level = {}
    ytd_sl = safe_num(ws_delivery.cell(rows['service'], 24).value, 100.0)
    target_sl = safe_num(ws_delivery.cell(rows['service'], 25).value, 100.0)
    if ytd_sl is not None: service_level['ytd_2026'] = ytd_sl
    if target_sl is not None: service_level['cible'] = target_sl

    units_sold = row_values(ws_delivery, rows['sold'], months_2026, 1.0, 21)
    defective = row_values(ws_delivery, rows['defective'], months_2026, 1.0, 21)
    crr = row_values(ws_delivery, rows['crr'], months_2026, 100.0, 21, 22)
    claims = row_values(ws_delivery, rows['claims'], months_2026, 1.0, 21, 22)
    cqr = {}
    dpmu = {}
    for month_key, _ in months_2026:
        if month_key in crr:
            cqr[month_key] = round(100.0 - crr[month_key], 4)
        sold = units_sold.get(month_key)
        defects = defective.get(month_key)
        if sold not in (None, 0) and defects is not None:
            dpmu[month_key] = round(defects / sold * 1_000_000, 4)
    if bu == 'MEL-T':
        cqr = row_values(ws_delivery, rows['cqr'], months_2026, 100.0, 21, 22)
        dpmu = row_values(ws_delivery, rows['dpmu'], months_2026, 1.0, 21, 22)
    else:
        if 'ytd_2026' in crr: cqr['ytd_2026'] = round(100.0 - crr['ytd_2026'], 4)
        if 'cible' in crr: cqr['cible'] = round(100.0 - crr['cible'], 4)
        sold_ytd = units_sold.get('ytd_2026')
        defects_ytd = defective.get('ytd_2026')
        if sold_ytd not in (None, 0) and defects_ytd is not None:
            dpmu['ytd_2026'] = round(defects_ytd / sold_ytd * 1_000_000, 4)

    pom = row_values(ws_wh, pom_rows[bu], warehouse_months_2026, 100.0, 18, 19)
    inv_correct = row_values(ws_wh, 17, warehouse_months_2026, 1.0, 18)
    inv_record = row_values(ws_wh, 18, warehouse_months_2026, 1.0, 18)
    inv_accuracy = row_values(ws_wh, 19, warehouse_months_2026, 100.0, 18, 19)

    log_json['sites'][bu] = [
        item('OTIF', 'OTIF', 'Orders delivered on time and in full at first confirmation.', 'OTIF orders / total orders × 100', '%', otif),
        item('MO count', 'MO Count', 'Manufacturing orders included in OTIF.', 'Count of manufacturing orders', 'MO', mo),
        item('LT Anticipo', 'LT Anticipo', 'Anticipation lead time.', 'Avvio - Partenza materiale', 'jours', anticipo),
        item('Lead Time de production', 'LT Production', 'Internal production lead time.', 'Spedizione - Avvio', 'jours', production_lt),
        item('Lead Time End-to-End', 'LT End-to-End', 'End-to-end delivery lead time.', 'Consegna - Partenza materiale', 'jours', end_to_end),
        item('Niveau de service', 'Service Level', 'YTD units sold versus budget.', 'Units sold / budget × 100', '%', service_level),
        item('Nombre de pièces expédiées', 'Shipped Pieces', 'Units sold/shipped reported in Delivery.', 'Units sold', 'pièces', units_shipped),
        item('# Units sold', 'Units sold', 'Units sold used for customer quality.', 'Count of units sold', 'pièces', units_sold),
        item('# Defective units', 'Defective units', 'Customer defective units.', 'Count of defective units', 'pièces', defective),
        item('CRR%', 'Customer Reject Rate', 'Customer reject rate.', 'Defective units / units sold × 100', '%', crr),
        item('Customer Quality Rate', 'CQR', 'Customer quality rate.', '100 - CRR', '%', cqr),
        item('DPMU (ppm)', 'DPMU', 'Defective units per million units sold.', 'Defective units / units sold × 1,000,000', 'ppm', dpmu),
        item('Customer Claims', 'Customer Claims', 'Customer claims registered in Delivery.', 'Count of claims', 'réclamations', claims),
        item('Fiabilité des livraisons', 'POM', 'Reliability of deliveries from Warehouse.', 'Error-free orders / total orders × 100', '%', pom),
        item('Précision des stocks', 'Inventory Accuracy Rate', 'Inventory accuracy from Warehouse.', 'Correct counted units / units on record × 100', '%', inv_accuracy),
        item('Unités conformes inventoriées', 'Correct Counted Units', 'Correct counted inventory units.', 'Count', 'pièces', inv_correct),
        item('Unités théoriques en stock', 'Units on Record', 'Units recorded in the inventory system.', 'Count', 'pièces', inv_record),
    ]

with open(log_data_path, 'w', encoding='utf-8') as f:
    json.dump(log_json, f, indent=2, ensure_ascii=False)
print("--> Updated", log_data_path)


# ==============================================================================
# 5. UPDATE KPI_METRICS.JSON (QHSE SST, ENERGY & WATER)
# ==============================================================================
print("--> Processing QHSE Metrics...")
metrics_path = 'data/kpi_metrics.json'
with open(metrics_path, 'r', encoding='utf-8') as f:
    metrics = json.load(f)

periods_2026 = [(32 + i, f'2026-{i + 1:02d}') for i in range(9)]
new_entries = []

def add_metric(section, site, indicator, unit, period, value):
    if value is None:
        return
    new_entries.append({
        'section': section,
        'site': site.lower().replace('|', '').replace(' + ', '_').replace(' ', '_'),
        'site_label': site,
        'indicateur': indicator,
        'unite': unit,
        'periode': period,
        'valeur': value,
    })

ws_sst = wb_qhse['KPI - SST']
sst_site_blocks = {'Safety': 14, 'Leather': 20, 'D|AIR': 26, 'Warehouse': 32, 'MEL-T': 39}
sst_indicators = [
    (1, 'Accidents du travail'), (2, 'Accidents avec arrêt'),
    (3, 'Jours perdus'), (4, 'Taux de fréquence (TF)'), (5, 'Taux de gravité (TG)'),
]
for site, base_row in sst_site_blocks.items():
    for offset, indicator in sst_indicators:
        for col, period in periods_2026:
            add_metric('s1_kpi_sst__sant___s_curit__au_travail', site, indicator, None, period, safe_num(ws_sst.cell(base_row + offset, col).value))

ws_nrg = wb_qhse['KPI ENERGY']
energy_sites = {
    'Safety + Leather': {'kwh': 14, 'co2': 16, 'intensity': 19},
    'D|AIR': {'kwh': 20, 'co2': 22, 'intensity': 24},
    'Warehouse': {'kwh': 25, 'co2': 27, 'intensity': None},
    'MEL-T': {'kwh': 29, 'co2': 31, 'intensity': 33},
}
for site, rows in energy_sites.items():
    for col, period in periods_2026:
        kwh = safe_num(ws_nrg.cell(rows['kwh'], col).value)
        # A trailing zero in this workbook means that the meter value was not supplied.
        if kwh not in (None, 0):
            add_metric('s2_kpi_energy___consommation__lectrique', site, "Consommation d'électricité", 'kWh', period, kwh)
            add_metric('s2_kpi_energy___consommation__lectrique', site, "Émission d'électricité", 'tCO₂', period, safe_num(ws_nrg.cell(rows['co2'], col).value))
        if rows['intensity']:
            intensity = safe_num(ws_nrg.cell(rows['intensity'], col).value)
            if intensity not in (None, 0):
                add_metric('s2_kpi_energy___consommation__lectrique', site, 'Intensité électrique (IE)', 'kWh/pièce', period, intensity)

for col, period in periods_2026:
    diesel = safe_num(ws_nrg.cell(48, col).value)
    add_metric('s3_kpi_energy___consommation_gasoil', 'MEL-T', 'Consommation Gasoil sans soufre', 'L', period, diesel)
    add_metric('s3_kpi_energy___consommation_gasoil', 'MEL-T', 'Équivalent énergétique de gasoil', 'TEP', period, safe_num(ws_nrg.cell(49, col).value))
    add_metric('s3_kpi_energy___consommation_gasoil', 'MEL-T', 'Émissions de gasoil', 'tCO₂', period, safe_num(ws_nrg.cell(50, col).value))
    # Global totals are omitted when the electricity component is missing.
    if safe_num(ws_nrg.cell(29, col).value) not in (None, 0):
        add_metric('s4_kpi_energy___bilan__nerg_tique_global', 'MEL-T', 'TEP totale', 'TEP', period, safe_num(ws_nrg.cell(61, col).value))
        add_metric('s4_kpi_energy___bilan__nerg_tique_global', 'MEL-T', 'Émissions totales CO₂', 'tCO₂', period, safe_num(ws_nrg.cell(63, col).value))

water_rows = {
    'Safety': (73, 77), 'Leather': (78, 82), 'D|AIR': (83, 87), 'MEL-T': (89, 93),
}
for site, (consumption_row, intensity_row) in water_rows.items():
    for col, period in periods_2026:
        consumption = safe_num(ws_nrg.cell(consumption_row, col).value)
        # Trailing zeros are placeholders for missing readings, not measured consumption.
        if consumption not in (None, 0):
            add_metric('s5_kpi_energy___consommation_eau', site, "Consommation d'eau", 'm³', period, consumption)
            add_metric('s5_kpi_energy___consommation_eau', site, 'Consommation L/P/J', 'L/P/J', period, safe_num(ws_nrg.cell(intensity_row, col).value))

replaced_indicators = {
    'Accidents du travail', 'Accidents avec arrêt', 'Jours perdus',
    'Taux de fréquence (TF)', 'Taux de gravité (TG)',
    "Consommation d'électricité", "Émission d'électricité", 'Intensité électrique (IE)',
    'Consommation Gasoil', 'Consommation Gasoil sans soufre',
    'Équivalent énergétique de gasoil', 'Émissions de gasoil',
    'TEP', 'TEP totale', 'Émissions totales CO₂',
    "Consommation d'eau", 'Consommation L/P/J',
}
monthly_2026 = {f'2026-{m:02d}' for m in range(1, 10)}
metrics = [m for m in metrics if not (m.get('indicateur') in replaced_indicators and m.get('periode') in monthly_2026)]
metrics.extend(new_entries)
metrics.sort(key=lambda m: (str(m.get('section')), str(m.get('site_label')), str(m.get('indicateur')), str(m.get('periode'))))
with open(metrics_path, 'w', encoding='utf-8') as f:
    json.dump(metrics, f, indent=2, ensure_ascii=False)
print(f"--> Added {len(new_entries)} exact 2026 QHSE entries to {metrics_path}")

wb_global.close()
wb_qhse.close()
print("=== Extraction completed successfully ===")
