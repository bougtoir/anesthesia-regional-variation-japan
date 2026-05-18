"""Generator script: appends the main block to write_papers.py"""

code = '''

# ==========================================
# MAIN
# ==========================================
if __name__ == '__main__':
    print("Generating BMJ-format papers with Spearman/Kendall trend analysis...")
    print(f"Dataset: {total_n} listings ({date_min_all} to {date_max_all})")
    print()
    for agent in ['Desflurane', 'Sevoflurane', 'Isoflurane']:
        tr = trend_results[agent]
        print(f"  {agent}:")
        print(f"    Spearman rho={tr['spearman_rho']:.3f}, P={tr['spearman_p']:.6f}")
        print(f"    Kendall tau={tr['kendall_tau']:.3f}, P={tr['kendall_p']:.6f}")
        print(f"    Quarterly rho={tr['quarterly_rho']:.3f}, P={tr['quarterly_p']:.6f}")
    print()
    write_english_paper()
    write_japanese_paper()
    print()
    print("Both papers generated successfully!")
    print(f"  English: {outdir}vaporizer_paper_english.docx")
    print(f"  Japanese: {outdir}vaporizer_paper_japanese.docx")
'''

with open('/home/ubuntu/vaporizer_research/write_papers.py', 'a') as f:
    f.write(code)
print("Main block appended successfully")
