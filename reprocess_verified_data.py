import sys
import os
import pandas as pd

sys.path.insert(0, '.')
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from nstc_scraper import (
    load_official_teacher_list,
    normalize_table1_9,
    normalize_table1_10,
    normalize_table1_11,
    verify_and_filter_teacher_ids,
    save_df_to_xls,
    save_to_3_formats
)

def main():
    print("=" * 60)
    print(" 執行教師身分編碼核實與專書 ISBN 篩選清理作業")
    print("=" * 60)

    roster_list, roster_map_name, roster_map_id = load_official_teacher_list('table1_1_List(老師清單).xls', full_time_only=True)
    print(f"[+] 載入 {len(roster_map_id)} 位輔英專任教師標準 ID (T0000xxxxxx)")

    files = [
        ('table1_9(期刊論文).xls', normalize_table1_9, '期刊論文'),
        ('table1_10(研討會論文).xls', normalize_table1_10, '研討會論文'),
        ('table1_11(專書).xls', normalize_table1_11, '專書')
    ]

    cat_dfs = {}

    for fname, normalizer, cat_key in files:
        if os.path.exists(fname):
            df = pd.read_excel(fname)
            orig_len = len(df)
            
            # Apply normalization (which extracts ISBN for books & filters non-ISBN)
            df_norm = normalizer(df)
            
            # Apply teacher ID verification against 280 full-time roster
            df_verified = verify_and_filter_teacher_ids(df_norm, roster_map_id=roster_map_id)
            
            save_df_to_xls(df_verified, fname)
            cat_dfs[cat_key] = df_verified
            print(f"[✓] {fname}: 原始筆數 {orig_len} -> 核實並篩選後筆數 {len(df_verified)} (去除 {orig_len - len(df_verified)} 筆無效/兼任/無ISBN資料)")
        else:
            print(f"[-] 檔案 {fname} 不存在")

    # Export to 3 output formats in output/ folder
    teachers_info = [r for r in roster_list]
    all_raw_pubs = [] # Build minimal pub representation
    save_to_3_formats(teachers_info, all_raw_pubs, cat_dfs, output_dir='output')
    print("\n[✓] 全數資料格式 (JSON, CSV, XLSX) 重新滙出完成！")

if __name__ == '__main__':
    main()
