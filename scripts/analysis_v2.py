import csv, io, json, statistics, pickle, math
from collections import defaultdict

# Load FIXED university hospital mapping
with open('/home/ubuntu/univ_hospital_mapping_v2.json') as f:
    univ_mapping = json.load(f)
univ_area_codes = set(univ_mapping.keys())

# Read the secondary medical area SCR file
with open('/home/ubuntu/scr_n_kubun.csv', 'rb') as f:
    raw = f.read()
text = raw.decode('shift_jis', errors='replace')
reader = csv.reader(io.StringIO(text))
rows = list(reader)

pref_nums = [x.strip() for x in rows[0][4:]]
pref_names = [x.strip() for x in rows[1][4:]]
area_codes = [x.strip() for x in rows[2][4:]]
area_names = [x.strip() for x in rows[3][4:]]
n_areas = len(area_codes)

area_info = {}
for i in range(n_areas):
    area_info[area_codes[i]] = {
        'pref_num': pref_nums[i],
        'pref_name': pref_names[i],
        'area_name': area_names[i],
        'has_univ': area_codes[i] in univ_area_codes,
        'n_univ': len(univ_mapping.get(area_codes[i], []))
    }

target_codes = {
    ('L', '0', '1'): 'L000_anesthesia_addon',
    ('L', '1', '1'): 'L001_IV_anesthesia',
    ('L', '2', '1'): 'L002_epidural',
    ('L', '3', '1'): 'L003_epidural_continuous',
    ('L', '4', '1'): 'L004_spinal',
    ('L', '5', '1'): 'L005_lower_limb_block',
    ('L', '6', '1'): 'L006_face_neck_block',
    ('L', '8', '1'): 'L008_general_anesthesia',
    ('L', '9', '1'): 'L009_mgmt_fee1',
    ('L', '10', '1'): 'L010_mgmt_fee2',
    ('L', '100', '1'): 'L100_nerve_block_inpt',
    ('L', '100', '2'): 'L100_nerve_block_outpt',
    ('L', '104', '1'): 'L104_trigger_point_inpt',
    ('L', '104', '2'): 'L104_trigger_point_outpt',
}

anesthesia_data = {}
for row in rows[4:]:
    if len(row) < 5:
        continue
    chapter = row[0].strip()
    code_num = row[1].strip()
    inout = row[3].strip()
    key = (chapter, code_num, inout)
    if key in target_codes:
        label = target_codes[key]
        values = {}
        for i in range(4, min(len(row), 4 + n_areas)):
            v = row[i].strip()
            if v:
                try:
                    values[area_codes[i-4]] = float(v)
                except ValueError:
                    pass
        anesthesia_data[label] = values

# --- ANALYSIS 1: Overall comparison ---
print("=" * 80)
print("ANALYSIS 1: University vs Non-University Areas (Overall)")
print("=" * 80)

key_codes_list = [
    'L008_general_anesthesia', 'L009_mgmt_fee1', 'L002_epidural',
    'L004_spinal', 'L005_lower_limb_block', 'L001_IV_anesthesia',
    'L100_nerve_block_inpt', 'L100_nerve_block_outpt', 'L104_trigger_point_outpt',
]

jp_names = {
    'L008_general_anesthesia': 'L008 全身麻酔',
    'L009_mgmt_fee1': 'L009 麻酔管理料1',
    'L002_epidural': 'L002 硬膜外麻酔',
    'L004_spinal': 'L004 脊椎麻酔',
    'L005_lower_limb_block': 'L005 下肢伝達麻酔',
    'L001_IV_anesthesia': 'L001 静脈麻酔',
    'L100_nerve_block_inpt': 'L100 神経ブロック(入院)',
    'L100_nerve_block_outpt': 'L100 神経ブロック(外来)',
    'L104_trigger_point_outpt': 'L104 トリガーポイント(外来)',
    'L010_mgmt_fee2': 'L010 麻酔管理料2',
}

for code_label in key_codes_list:
    if code_label not in anesthesia_data:
        continue
    values = anesthesia_data[code_label]
    univ_vals = [values[ac] for ac in values if ac in univ_area_codes]
    non_univ_vals = [values[ac] for ac in values if ac not in univ_area_codes]
    if len(univ_vals) < 3 or len(non_univ_vals) < 3:
        continue
    u_mean = statistics.mean(univ_vals)
    n_mean = statistics.mean(non_univ_vals)
    u_sd = statistics.stdev(univ_vals)
    n_sd = statistics.stdev(non_univ_vals)
    diff = u_mean - n_mean
    pooled_sd = ((u_sd**2*(len(univ_vals)-1) + n_sd**2*(len(non_univ_vals)-1)) / (len(univ_vals)+len(non_univ_vals)-2))**0.5
    d = diff/pooled_sd if pooled_sd > 0 else 0
    se = math.sqrt(u_sd**2/len(univ_vals) + n_sd**2/len(non_univ_vals))
    t = diff/se if se > 0 else 0
    jn = jp_names.get(code_label, code_label)
    print(f"\n{jn}:")
    print(f"  大学病院圏(n={len(univ_vals):3d}): mean={u_mean:7.1f}, SD={u_sd:7.1f}")
    print(f"  非大学病院圏(n={len(non_univ_vals):3d}): mean={n_mean:7.1f}, SD={n_sd:7.1f}")
    print(f"  差={diff:+.1f}, Cohen's d={d:+.3f}, t={t:.2f}")

# --- ANALYSIS 3: Within-prefecture comparison ---
print("\n" + "=" * 80)
print("ANALYSIS 3: Within-Prefecture Comparison (controls for audit policy)")
print("=" * 80)

pref_areas = defaultdict(list)
for ac in area_codes:
    pref_areas[area_info[ac]['pref_num']].append(ac)

analysis_codes = ['L008_general_anesthesia', 'L009_mgmt_fee1', 'L002_epidural',
                   'L004_spinal', 'L001_IV_anesthesia', 'L100_nerve_block_inpt']

for code_label in analysis_codes:
    if code_label not in anesthesia_data:
        continue
    values = anesthesia_data[code_label]
    within_diffs = []
    for pref_num, areas in sorted(pref_areas.items()):
        uv = [values[ac] for ac in areas if ac in univ_area_codes and ac in values]
        nv = [values[ac] for ac in areas if ac not in univ_area_codes and ac in values]
        if len(uv) >= 1 and len(nv) >= 1:
            within_diffs.append(statistics.mean(uv) - statistics.mean(nv))
    if len(within_diffs) < 3:
        continue
    mean_d = statistics.mean(within_diffs)
    sd_d = statistics.stdev(within_diffs)
    se = sd_d / math.sqrt(len(within_diffs))
    t = mean_d / se if se > 0 else 0
    n_pos = sum(1 for d in within_diffs if d > 0)
    jn = jp_names.get(code_label, code_label)
    print(f"\n{jn}:")
    print(f"  県内差: mean={mean_d:+.1f}, SD={sd_d:.1f}, t={t:.2f} (n={len(within_diffs)}県)")
    print(f"  大学圏>非大学圏: {n_pos}/{len(within_diffs)} ({100*n_pos/len(within_diffs):.0f}%)")

# --- ANALYSIS 6: 3-level variance decomposition ---
print("\n" + "=" * 80)
print("ANALYSIS 6: Three-Level Variance Decomposition")
print("=" * 80)

for code_label in analysis_codes:
    if code_label not in anesthesia_data:
        continue
    values = anesthesia_data[code_label]
    all_vals = [values[ac] for ac in area_codes if ac in values]
    grand_mean = statistics.mean(all_vals)

    pref_areas_d = defaultdict(list)
    for ac in area_codes:
        if ac in values:
            pref_areas_d[area_info[ac]['pref_num']].append(ac)

    ss_total = sum((v - grand_mean)**2 for v in all_vals)
    ss_between_pref = 0
    ss_univ_effect = 0
    ss_residual = 0

    for pref_num, areas in pref_areas_d.items():
        pref_vals = [values[ac] for ac in areas]
        pref_mean = statistics.mean(pref_vals)
        ss_between_pref += len(areas) * (pref_mean - grand_mean)**2

        if len(areas) >= 2:
            uv = [values[ac] for ac in areas if ac in univ_area_codes]
            nv = [values[ac] for ac in areas if ac not in univ_area_codes]
            if len(uv) >= 1 and len(nv) >= 1:
                u_mean = statistics.mean(uv)
                n_mean = statistics.mean(nv)
                ss_u = len(uv)*(u_mean-pref_mean)**2 + len(nv)*(n_mean-pref_mean)**2
                ss_univ_effect += ss_u
                ss_r = sum((values[ac]-u_mean)**2 for ac in areas if ac in univ_area_codes)
                ss_r += sum((values[ac]-n_mean)**2 for ac in areas if ac not in univ_area_codes)
                ss_residual += ss_r
            else:
                ss_residual += sum((v-pref_mean)**2 for v in pref_vals)

    if ss_total > 0:
        pct_p = 100*ss_between_pref/ss_total
        pct_u = 100*ss_univ_effect/ss_total
        pct_r = 100*ss_residual/ss_total
        pct_w = 100*(ss_total-ss_between_pref)/ss_total
        jn = jp_names.get(code_label, code_label)
        print(f"\n{jn}:")
        print(f"  県間 {pct_p:.1f}% + 大学病院効果 {pct_u:.1f}% + 残差 {pct_r:.1f}%")
        if pct_w > 0:
            print(f"  県内分散{pct_w:.1f}%のうち大学病院効果: {100*pct_u/pct_w:.1f}%")

# --- ANALYSIS 7: Proximity gradient ---
print("\n" + "=" * 80)
print("ANALYSIS 7: Proximity Gradient (A=univ area, B=same pref, C=no-univ pref)")
print("=" * 80)

pref_has_univ = set()
for ac in area_codes:
    if ac in univ_area_codes:
        pref_has_univ.add(area_info[ac]['pref_num'])

type_a, type_b, type_c = [], [], []
for ac in area_codes:
    pref = area_info[ac]['pref_num']
    if ac in univ_area_codes:
        type_a.append(ac)
    elif pref in pref_has_univ:
        type_b.append(ac)
    else:
        type_c.append(ac)

print(f"A(univ area): {len(type_a)}, B(same pref): {len(type_b)}, C(no-univ pref): {len(type_c)}")

for code_label in analysis_codes:
    if code_label not in anesthesia_data:
        continue
    values = anesthesia_data[code_label]
    a = [values[ac] for ac in type_a if ac in values]
    b = [values[ac] for ac in type_b if ac in values]
    c = [values[ac] for ac in type_c if ac in values]
    if len(a) < 2 or len(b) < 2:
        continue
    am, bm = statistics.mean(a), statistics.mean(b)
    cm = statistics.mean(c) if len(c) >= 2 else float('nan')
    pattern = ""
    if not math.isnan(cm):
        if am > bm > cm:
            pattern = "A>B>C gradient"
        elif am > bm and bm <= cm:
            pattern = "A>B<=C (direct only)"
        else:
            pattern = "no gradient"
    jn = jp_names.get(code_label, code_label)
    print(f"\n{jn}:")
    print(f"  A={am:.1f}, B={bm:.1f}, C={cm:.1f} => {pattern}")
    if not math.isnan(cm):
        print(f"  Direct(A-B)={am-bm:+.1f}, Spillover(B-C)={bm-cm:+.1f}")

# Save updated data
with open('/home/ubuntu/analysis_data_v2.pkl', 'wb') as f:
    pickle.dump({
        'area_info': area_info,
        'anesthesia_data': anesthesia_data,
        'univ_mapping': univ_mapping,
        'univ_area_codes': univ_area_codes,
        'area_codes': area_codes,
        'pref_nums': pref_nums,
        'pref_names': pref_names,
        'area_names': area_names,
    }, f)

print("\nDone. Data saved to analysis_data_v2.pkl")
