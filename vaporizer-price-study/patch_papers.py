#!/usr/bin/env python3
"""Patch write_papers.py to add literature review citations to Japanese sections."""

import re

with open('/home/ubuntu/vaporizer_research/write_papers.py', 'r') as f:
    lines = f.readlines()

content = ''.join(lines)

# 1. Replace Japanese Introduction "no study" paragraph
# Find the exact block by line content matching
old_block_1 = (
    "    doc.add_paragraph(\n"
    "        '\\u6211\\u3005\\u306e\\u77e5\\u308b\\u9650\\u308a\\u3001\\u74b0\\u5883\\u898f\\u5236\\u304c\\u9ebb\\u9154\\u6a5f\\u5668\\u306e\\u4e2d\\u53e4\\u5e02\\u5834\\u4fa1\\u5024\\u306b\\u4e0e\\u3048\\u308b\\u5f71\\u97ff\\u3092\\u691c\\u8a0e\\u3057\\u305f\\u7814\\u7a76\\u306f\\u306a\\u3044\\u3002'\n"
    "        '\\u6211\\u3005\\u306f\\u3001EU\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u898f\\u5236\\u304c\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u6c17\\u5316\\u5668\\u306e\\u4e2d\\u53e4\\u5e02\\u5834\\u4fa1\\u683c\\u306e\\u6bb5\\u968e\\u7684\\u306a\\u4e0b\\u843d\\u3068\\u95a2\\u9023\\u3057\\u3001'\n"
    "        '\\u30bb\\u30dc\\u30d5\\u30eb\\u30e9\\u30f3\\u304a\\u3088\\u3073\\u30a4\\u30bd\\u30d5\\u30eb\\u30e9\\u30f3\\u6c17\\u5316\\u5668\\u306e\\u4fa1\\u683c\\u306f\\u5b89\\u5b9a\\u3059\\u308b\\u3068\\u4eee\\u8aac\\u3092\\u7acb\\u3066\\u305f\\u3002'\n"
    "        'eBay\\u306e3\\u5e74\\u5206\\u306e\\u843d\\u672d\\u30c7\\u30fc\\u30bf\\uff08Terapeak\\u7d4c\\u7531\\uff09\\u3092\\u7528\\u3044\\u3001\\u6a2a\\u65ad\\u7684\\u6bd4\\u8f03\\u3068\\u6642\\u7cfb\\u5217\\u30c8\\u30ec\\u30f3\\u30c9\\u5206\\u6790\\u306e\\u4e21\\u65b9\\u3067\\u3053\\u306e\\u4eee\\u8aac\\u3092\\u691c\\u8a3c\\u3057\\u305f\\u3002')\n"
)

new_block_1 = (
    "    doc.add_paragraph(\n"
    "        '\\u5148\\u884c\\u7814\\u7a76\\u3067\\u306f\\u3001\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u4f7f\\u7528\\u4e2d\\u6b62\\u306e\\u8ca1\\u52d9\\u7684\\u6839\\u62e0[16]\\u3001'\n"
    "        '\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u5ec3\\u6b62\\u306e\\u81e8\\u5e8a\\u7684\\u30fb\\u653f\\u7b56\\u7684\\u542b\\u610f[17,18]\\u3001'\n"
    "        '\\u304a\\u3088\\u3073\\u65bd\\u8a2d\\u30ec\\u30d9\\u30eb\\u3067\\u306e\\u6c17\\u5316\\u5668\\u64a4\\u53bb\\u30d7\\u30ed\\u30b0\\u30e9\\u30e0\\u306e\\u6709\\u52b9\\u6027[15]\\u304c\\u691c\\u8a0e\\u3055\\u308c\\u3066\\u3044\\u308b\\u3002'\n"
    "        '\\u63ee\\u767a\\u6027\\u9ebb\\u9154\\u85ac\\u6d88\\u8cbb\\u524a\\u6e1b\\u306b\\u3088\\u308b\\u30b3\\u30b9\\u30c8\\u524a\\u6e1b\\u306e\\u7d4c\\u6e08\\u5206\\u6790[9,19]\\u3084\\u3001'\n"
    "        '\\u4ed6\\u306e\\u6a5f\\u5668\\u30ab\\u30c6\\u30b4\\u30ea\\u30fc\\u306b\\u304a\\u3051\\u308b\\u4e2d\\u53e4\\u533b\\u7642\\u6a5f\\u5668\\u5e02\\u5834\\u306e\\u7279\\u5fb4\\u4ed8\\u3051[20]\\u3082\\u884c\\u308f\\u308c\\u3066\\u3044\\u308b\\u3002'\n"
    "        '\\u3057\\u304b\\u3057\\u3001\\u6211\\u3005\\u306e\\u77e5\\u308b\\u9650\\u308a\\u3001\\u74b0\\u5883\\u898f\\u5236\\u304c\\u9ebb\\u9154\\u6a5f\\u5668\\u306e\\u4e2d\\u53e4\\u5e02\\u5834\\u4fa1\\u5024\\u306b\\u4e0e\\u3048\\u308b\\u5f71\\u97ff\\u3092\\u691c\\u8a0e\\u3057\\u305f\\u7814\\u7a76\\u306f\\u306a\\u3044\\u3002'\n"
    "        '\\u6211\\u3005\\u306f\\u3001EU\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u898f\\u5236\\u304c\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u6c17\\u5316\\u5668\\u306e\\u4e2d\\u53e4\\u5e02\\u5834\\u4fa1\\u683c\\u306e\\u6bb5\\u968e\\u7684\\u306a\\u4e0b\\u843d\\u3068\\u95a2\\u9023\\u3057\\u3001'\n"
    "        '\\u30bb\\u30dc\\u30d5\\u30eb\\u30e9\\u30f3\\u304a\\u3088\\u3073\\u30a4\\u30bd\\u30d5\\u30eb\\u30e9\\u30f3\\u6c17\\u5316\\u5668\\u306e\\u4fa1\\u683c\\u306f\\u5b89\\u5b9a\\u3059\\u308b\\u3068\\u4eee\\u8aac\\u3092\\u7acb\\u3066\\u305f\\u3002'\n"
    "        'eBay\\u306e3\\u5e74\\u5206\\u306e\\u843d\\u672d\\u30c7\\u30fc\\u30bf\\uff08Terapeak\\u7d4c\\u7531\\uff09\\u3092\\u7528\\u3044\\u3001\\u6a2a\\u65ad\\u7684\\u6bd4\\u8f03\\u3068\\u6642\\u7cfb\\u5217\\u30c8\\u30ec\\u30f3\\u30c9\\u5206\\u6790\\u306e\\u4e21\\u65b9\\u3067\\u3053\\u306e\\u4eee\\u8aac\\u3092\\u691c\\u8a3c\\u3057\\u305f\\u3002')\n"
)

if old_block_1 in content:
    content = content.replace(old_block_1, new_block_1, 1)
    print("1. Japanese Introduction updated successfully")
else:
    print("ERROR: Japanese intro block not found!")
    # Debug: find the line
    for i, line in enumerate(lines):
        if '\\u6211\\u3005\\u306e\\u77e5\\u308b\\u9650\\u308a' in line:
            print(f"  Found at line {i+1}: {repr(line[:100])}")

# 2. Add "Comparison with other studies" subsection before "Strengths and limitations" in Japanese Discussion
old_block_2 = "    add_heading_styled(doc, '\\u5f37\\u307f\\u3068\\u9650\\u754c', level=2)\n"

jp_comparison_section = (
    "    add_heading_styled(doc, '\\u5148\\u884c\\u7814\\u7a76\\u3068\\u306e\\u6bd4\\u8f03', level=2)\n"
    "    doc.add_paragraph(\n"
    "        '\\u6211\\u3005\\u306e\\u77e5\\u308b\\u9650\\u308a\\u3001\\u9ebb\\u9154\\u6a5f\\u5668\\u306b\\u5bfe\\u3059\\u308b\\u74b0\\u5883\\u898f\\u5236\\u306e\\u4e2d\\u53e4\\u5e02\\u5834\\u3078\\u306e\\u5f71\\u97ff\\u3092\\u691c\\u8a0e\\u3057\\u305f\\u5148\\u884c\\u7814\\u7a76\\u306f\\u306a\\u3044\\u3002'\n"
    "        'Lehmann\\u3089[15]\\u306f\\u3001\\u6559\\u80b2\\u3068\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u6c17\\u5316\\u5668\\u306e\\u7269\\u7406\\u7684\\u64a4\\u53bb\\u3092\\u7d44\\u307f\\u5408\\u308f\\u305b\\u305f\\u65bd\\u8a2d\\u30ec\\u30d9\\u30eb\\u306e\\u4ecb\\u5165\\u306b\\u3088\\u308a\\u3001'\n"
    "        '\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u8d77\\u56e0\\u306eCO\\u2082\\u7b49\\u4fa1\\u6392\\u51fa\\u91cf\\u304c86%\\u524a\\u6e1b\\u3055\\u308c\\u305f\\u3053\\u3068\\u3092\\u5b9f\\u8a3c\\u3057\\u305f\\u304c\\u3001'\n"
    "        '\\u85ac\\u5264\\u6d88\\u8cbb\\u91cf\\u3092\\u6e2c\\u5b9a\\u3057\\u305f\\u3082\\u306e\\u3067\\u3042\\u308a\\u3001\\u6a5f\\u5668\\u306e\\u518d\\u8ca9\\u58f2\\u4fa1\\u5024\\u306f\\u691c\\u8a0e\\u3057\\u3066\\u3044\\u306a\\u3044\\u3002'\n"
    "        'Meyer[16]\\u304a\\u3088\\u3073Mohammed\\u3068Metta[18]\\u306f\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u4e2d\\u6b62\\u306e\\u4e16\\u754c\\u7684\\u30fb\\u8ca1\\u52d9\\u7684\\u6839\\u62e0\\u3092\\u8ad6\\u3058\\u3001'\n"
    "        'Moonesinghe[17]\\u306f\\u5ec3\\u6b62\\u30d7\\u30ed\\u30b0\\u30e9\\u30e0\\u306e\\u5e83\\u7bc4\\u306a\\u542b\\u610f\\u3092\\u8b70\\u8ad6\\u3057\\u305f\\u304c\\u3001'\n"
    "        '\\u3044\\u305a\\u308c\\u3082\\u4e2d\\u53e4\\u6a5f\\u5668\\u5e02\\u5834\\u3078\\u306e\\u4e0b\\u6d41\\u5f71\\u97ff\\u306f\\u691c\\u8a0e\\u3057\\u3066\\u3044\\u306a\\u3044\\u3002'\n"
    "        'Beard\\u3089[19]\\u306f\\u7d42\\u672b\\u547c\\u6c17\\u6fc3\\u5ea6\\u5236\\u5fa1\\u306b\\u3088\\u308b\\u7d4c\\u6e08\\u7684\\u4fbf\\u76ca\\u3092\\u5b9a\\u91cf\\u5316\\u3057\\u305f\\u304c\\u3001\\u6a5f\\u5668\\u306e\\u6e1b\\u4fa1\\u512a\\u5374\\u306f\\u6271\\u3063\\u3066\\u3044\\u306a\\u3044\\u3002'\n"
    "        'BFMV\\u4e2d\\u53e4\\u533b\\u7642\\u6a5f\\u5668\\u4fa1\\u683c\\u30d9\\u30f3\\u30c1\\u30de\\u30fc\\u30af\\u30ec\\u30dd\\u30fc\\u30c8[20]\\u306f\\u7d041,500\\u6a5f\\u7a2e\\u306e\\u4e2d\\u53e4\\u533b\\u7642\\u6a5f\\u5668\\u306e\\u5e74\\u6b21\\u4fa1\\u683c\\u30d9\\u30f3\\u30c1\\u30de\\u30fc\\u30af\\u3092\\u63d0\\u4f9b\\u3057\\u3001'\n"
    "        '\\u591a\\u304f\\u306e\\u6a5f\\u5668\\u30ab\\u30c6\\u30b4\\u30ea\\u30fc\\u306e\\u518d\\u8ca9\\u58f2\\u4fa1\\u5024\\u304c5\\u5e74\\u9593\\u6bd4\\u8f03\\u7684\\u5b89\\u5b9a\\u3067\\u3042\\u308b\\u3053\\u3068\\u3092\\u793a\\u3057\\u3066\\u3044\\u308b\\u3002'\n"
    "        '\\u672c\\u7814\\u7a76\\u306e\\u30a4\\u30bd\\u30d5\\u30eb\\u30e9\\u30f3\\u30fb\\u30bb\\u30dc\\u30d5\\u30eb\\u30e9\\u30f3\\u6c17\\u5316\\u5668\\u4fa1\\u683c\\u306e\\u5b89\\u5b9a\\u6027\\u306f\\u3053\\u308c\\u3068\\u4e00\\u81f4\\u3059\\u308b\\u304c\\u3001'\n"
    "        '\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u306e\\u4e0b\\u843d\\u306f\\u898f\\u5236\\u4ecb\\u5165\\u306b\\u8d77\\u56e0\\u3059\\u308b\\u7279\\u7b46\\u3059\\u3079\\u304d\\u4f8b\\u5916\\u3067\\u3042\\u308b\\u3002')\n"
    "    doc.add_paragraph(\n"
    "        '\\u672c\\u7814\\u7a76\\u306e\\u77e5\\u898b\\u306f\\u3001\\u898f\\u5236\\u306b\\u3088\\u308b\\u9673\\u8150\\u5316[14]\\u306b\\u95a2\\u3059\\u308b\\u5e83\\u7bc4\\u306a\\u7d4c\\u6e08\\u5b66\\u6587\\u732e\\u3068\\u4e00\\u81f4\\u3057\\u3066\\u304a\\u308a\\u3001'\n"
    "        '\\u4e88\\u60f3\\u3055\\u308c\\u308b\\u653f\\u5e9c\\u898f\\u5236\\u304c\\u4e2d\\u53e4\\u5e02\\u5834\\u306b\\u304a\\u3051\\u308b\\u4e88\\u6e2c\\u7684\\u306a\\u4fa1\\u683c\\u4e0b\\u843d\\u3092\\u5f15\\u304d\\u8d77\\u3053\\u3059\\u3002'\n"
    "        '\\u7acb\\u6cd5\\u904e\\u7a0b\\uff082022\\u5e74\\uff5e2024\\u5e74\\uff09\\u306b\\u304a\\u3051\\u308b\\u6bb5\\u968e\\u7684\\u306a\\u4fa1\\u683c\\u4fb5\\u98df\\u3068\\u3001'\n"
    "        '\\u898f\\u5236\\u5f8c\\u306e\\u3088\\u308a\\u9855\\u8457\\u306a\\u4e0b\\u843d\\u3068\\u3044\\u3046\\u30d1\\u30bf\\u30fc\\u30f3\\u306f\\u3001'\n"
    "        '\\u8eca\\u4e21\\u6392\\u51fa\\u898f\\u5236\\u304c\\u4e2d\\u53e4\\u8eca\\u5e02\\u5834\\u306b\\u4e0e\\u3048\\u305f\\u5f71\\u97ff\\u306e\\u7814\\u7a76\\u3068\\u985e\\u4f3c\\u3057\\u3066\\u3044\\u308b\\u3002'\n"
    "        '\\u30c7\\u30b9\\u30d5\\u30eb\\u30e9\\u30f3\\u306e\\u307f\\u306b\\u5f71\\u97ff\\u3057\\u30bb\\u30dc\\u30d5\\u30eb\\u30e9\\u30f3\\u30fb\\u30a4\\u30bd\\u30d5\\u30eb\\u30e9\\u30f3\\u4fa1\\u683c\\u306f\\u5909\\u5316\\u3057\\u306a\\u304b\\u3063\\u305f\\u3068\\u3044\\u3046\\u85ac\\u5264\\u7279\\u7570\\u6027\\u306f\\u3001'\n"
    "        '\\u4e00\\u822c\\u7684\\u306a\\u5e02\\u5834\\u52b9\\u679c\\u3067\\u306f\\u306a\\u304f\\u898f\\u5236\\u52b9\\u679c\\u3067\\u3042\\u308b\\u3053\\u3068\\u306e\\u7279\\u306b\\u5f37\\u3044\\u30a8\\u30d3\\u30c7\\u30f3\\u30b9\\u3092\\u63d0\\u4f9b\\u3059\\u308b\\u3002')\n"
    "\n"
    "    add_heading_styled(doc, '\\u5f37\\u307f\\u3068\\u9650\\u754c', level=2)\n"
)

if old_block_2 in content:
    content = content.replace(old_block_2, jp_comparison_section, 1)
    print("2. Japanese Discussion 'Comparison' subsection added successfully")
else:
    print("ERROR: Japanese strengths heading not found!")
    for i, line in enumerate(lines):
        if '\\u5f37\\u307f\\u3068\\u9650\\u754c' in line:
            print(f"  Found at line {i+1}: {repr(line[:100])}")

# 3. Update Japanese references
old_jp_ref_14 = "        '14. Davis G, Patel N. Regulatory obsolescence and secondary market asset depreciation. J Environ Econ Manag 2019;95:142-60.',\n    ]\n    for ref in references:\n        p = doc.add_paragraph(ref)\n        p.paragraph_format.space_after = Pt(4)\n        for run in p.runs:\n            run.font.size = Pt(10)\n\n    doc.save(outdir + 'vaporizer_paper_japanese.docx')"

new_jp_refs = (
    "        '14. Davis G, Patel N. Regulatory obsolescence and secondary market asset depreciation. J Environ Econ Manag 2019;95:142-60.',\n"
    "        '15. Lehmann H, Werning J, Baschnegger H, et al. Minimising the usage of desflurane only by education and removal of the vaporisers. BMC Anesthesiol 2025;25:108.',\n"
    "        '16. Meyer MJ. Desflurane should des-appear: global and financial rationale. Anesth Analg 2020;131:1317-22.',\n"
    "        '17. Moonesinghe SR. Desflurane decommissioning: more than meets the eye. Anaesthesia 2024;79:237-41.',\n"
    "        '18. Mohammed A, Metta H. Is it time to bid adieu to desflurane? J Anaesthesiol Clin Pharmacol 2025;41:211-2.',\n"
    "        '19. Beard D, Aston W, Black S, et al. Environmental and economic impacts of end-tidal control of volatile anaesthetics. Open Anaesthesia J 2025;19:e18742126.',\n"
    "        '20. Buckhead Fair Market Value. 2025 Benchmark Report on Pre-Owned Medical Equipment Prices. Atlanta, GA: BFMV, 2025.',\n"
    "    ]\n"
    "    for ref in references:\n"
    "        p = doc.add_paragraph(ref)\n"
    "        p.paragraph_format.space_after = Pt(4)\n"
    "        for run in p.runs:\n"
    "            run.font.size = Pt(10)\n"
    "\n"
    "    doc.save(outdir + 'vaporizer_paper_japanese.docx')"
)

if old_jp_ref_14 in content:
    content = content.replace(old_jp_ref_14, new_jp_refs, 1)
    print("3. Japanese References updated successfully")
else:
    print("ERROR: Japanese references block not found!")

# Write back
with open('/home/ubuntu/vaporizer_research/write_papers.py', 'w') as f:
    f.write(content)

print("\nAll Japanese paper modifications saved!")
