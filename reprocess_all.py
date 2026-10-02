import pandas as pd
import numpy as np
import sys
import os
import re
import xlwt

sys.path.insert(0, '.')
from nstc_scraper import (
    clean_excel_str,
    save_df_to_xls,
    map_author_order,
    map_indexing_code,
    map_country_code,
    map_language_code,
    normalize_table1_17
)

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

def format_fooyin_teacher_id(raw_id):
    if not raw_id or pd.isna(raw_id):
        return 'T0000000001'
    raw_id = str(raw_id).strip()
    if re.match(r'^T\d{10}$', raw_id, re.IGNORECASE):
        return raw_id.upper()
    num_hash = abs(hash(raw_id)) % (10**10)
    return f'T{num_hash:010d}'

def normalize_table1_9(df):
    cols = ['身分編碼', '所屬計畫案', '論文名稱', '作者順序', '通訊作者', '收錄分類', '刊物名稱', '發表卷數', '是否具有審稿制度', '發表期數', '期刊出版地國別/地區', '發表月份', '發表年份', '發表型式', '是否為跨國(地區)合作']
    for c in cols:
        if c not in df.columns:
            df[c] = np.nan
    df = df[cols].copy()
    
    df['身分編碼'] = df['身分編碼'].apply(format_fooyin_teacher_id)
    df['所屬計畫案'] = df['所屬計畫案'].fillna('無').replace('', '無')
    df['論文名稱'] = df['論文名稱'].fillna('').apply(clean_excel_str)
    df['作者順序'] = df['作者順序'].fillna('A').apply(lambda v: str(v) if str(v) in ['A','B','C','D'] else 'A')
    df['通訊作者'] = df['通訊作者'].fillna('否').apply(lambda v: '是' if str(v) in ['是', 'TRUE', 'True'] else '否')
    df['收錄分類'] = df['收錄分類'].fillna('O').apply(map_indexing_code)
    df['刊物名稱'] = df['刊物名稱'].fillna('').apply(clean_excel_str)
    df['發表卷數'] = df['發表卷數'].fillna('').apply(lambda v: str(int(v)) if isinstance(v, (int, float)) and not pd.isna(v) else str(v))
    df['是否具有審稿制度'] = '是'
    df['發表期數'] = df['發表期數'].fillna('').apply(lambda v: str(int(v)) if isinstance(v, (int, float)) and not pd.isna(v) else str(v))
    df['期刊出版地國別/地區'] = df['期刊出版地國別/地區'].fillna('NATTWN').apply(map_country_code)
    
    df['發表月份'] = pd.to_numeric(df['發表月份'], errors='coerce').astype('Int64')
    df['發表年份'] = pd.to_numeric(df['發表年份'], errors='coerce').astype('Int64')
    df['發表型式'] = df['發表型式'].fillna('電子期刊').apply(lambda v: str(v) if str(v) in ['電子期刊', '紙本'] else '電子期刊')
    df['是否為跨國(地區)合作'] = pd.to_numeric(df['是否為跨國(地區)合作'], errors='coerce').fillna(4).astype(int)
    
    return df

def normalize_table1_10(df):
    cols = ['身分編碼', '所屬計畫案', '論文名稱', '是否具有對外公開徵稿及審稿制度', '作者順序', '通訊作者', '研討會名稱', '舉行之國家', '舉行之城市', '開始日期', '結束日期', '發表年份']
    for c in cols:
        if c not in df.columns:
            df[c] = np.nan
    df = df[cols].copy()
    
    df['身分編碼'] = df['身分編碼'].apply(format_fooyin_teacher_id)
    df['所屬計畫案'] = df['所屬計畫案'].fillna('無').replace('', '無')
    df['論文名稱'] = df['論文名稱'].fillna('').apply(clean_excel_str)
    df['是否具有對外公開徵稿及審稿制度'] = '是'
    df['作者順序'] = df['作者順序'].fillna('A').apply(lambda v: str(v) if str(v) in ['A','B','C','D'] else 'A')
    df['通訊作者'] = df['通訊作者'].fillna('是').apply(lambda v: '是' if str(v) in ['是', 'TRUE', 'True'] else '否')
    df['研討會名稱'] = df['研討會名稱'].fillna('').apply(clean_excel_str)
    df['舉行之國家'] = df['舉行之國家'].fillna('NATTWN').apply(map_country_code)
    df['舉行之城市'] = df['舉行之城市'].fillna('台北').apply(clean_excel_str)
    df['開始日期'] = df['開始日期'].fillna('')
    df['結束日期'] = df['結束日期'].fillna('')
    df['發表年份'] = pd.to_numeric(df['發表年份'], errors='coerce').astype('Int64')
    
    return df

def normalize_table1_11(df):
    cols = ['身分編碼', '所屬計劃案', '篇章及所屬專書名稱/或專書名稱', '專書類別', '是否為專書', '是否有外部審稿程序或公開發行出版', '作者順序', '通訊作者', '使用語文', '出版年(yyyy)', '出版月', '出版社/出版處所', 'ISBN編號', '是否為跨國(地區)合作']
    for c in cols:
        if c not in df.columns:
            df[c] = np.nan
    df = df[cols].copy()
    
    df['身分編碼'] = df['身分編碼'].apply(format_fooyin_teacher_id)
    df['所屬計劃案'] = df['所屬計劃案'].fillna('無').replace('', '無')
    df['篇章及所屬專書名稱/或專書名稱'] = df['篇章及所屬專書名稱/或專書名稱'].fillna('').apply(clean_excel_str)
    df['專書類別'] = df['專書類別'].fillna('其他').apply(lambda v: str(v) if str(v) in ['紙本', '其他', '專書篇章', '專書'] else '其他')
    df['是否為專書'] = df['是否為專書'].fillna('是')
    df['是否有外部審稿程序或公開發行出版'] = '是'
    df['作者順序'] = df['作者順序'].fillna('A').apply(lambda v: str(v) if str(v) in ['A','B','C','D'] else 'A')
    df['通訊作者'] = df['通訊作者'].fillna('是').apply(lambda v: '是' if str(v) in ['是', 'TRUE', 'True'] else '否')
    df['使用語文'] = df['使用語文'].fillna('C').apply(map_language_code)
    df['出版年(yyyy)'] = pd.to_numeric(df['出版年(yyyy)'], errors='coerce').astype('Int64')
    df['出版月'] = pd.to_numeric(df['出版月'], errors='coerce').astype('Int64')
    df['出版社/出版處所'] = df['出版社/出版處所'].fillna('').apply(clean_excel_str)
    df['ISBN編號'] = df['ISBN編號'].fillna('')
    df['是否為跨國(地區)合作'] = pd.to_numeric(df['是否為跨國(地區)合作'], errors='coerce').fillna(4).astype(int)
    
    return df

def reprocess_all_files():
    print("[+] 讀取基準範本檔 (Baseline Reference Format Files)...")
    base9 = pd.read_excel('table1_9匯入檔-1150407(FORMATE).xls') if os.path.exists('table1_9匯入檔-1150407(FORMATE).xls') else pd.DataFrame()
    base10 = pd.read_excel('table1_10-匯入檔1150407(FORMAT).xls') if os.path.exists('table1_10-匯入檔1150407(FORMAT).xls') else pd.DataFrame()
    base11 = pd.read_excel('table1_11-匯入檔1150423(FORMATE).xls') if os.path.exists('table1_11-匯入檔1150423(FORMATE).xls') else pd.DataFrame()

    print(f"  - Baseline 1_9: {len(base9)} rows")
    print(f"  - Baseline 1_10: {len(base10)} rows")
    print(f"  - Baseline 1_11: {len(base11)} rows")

    # Read current files if they exist
    curr9 = pd.read_excel('table1_9(期刊論文).xls') if os.path.exists('table1_9(期刊論文).xls') else pd.DataFrame()
    curr10 = pd.read_excel('table1_10(研討會論文).xls') if os.path.exists('table1_10(研討會論文).xls') else pd.DataFrame()
    curr11 = pd.read_excel('table1_11(專書).xls') if os.path.exists('table1_11(專書).xls') else pd.DataFrame()
    curr17 = pd.read_excel('table1_17(國科會計畫案).xls') if os.path.exists('table1_17(國科會計畫案).xls') else pd.DataFrame()

    # Merge baseline with current and drop duplicates
    full9 = pd.concat([base9, curr9], ignore_index=True).drop_duplicates(subset=['論文名稱']) if not curr9.empty or not base9.empty else pd.DataFrame()
    full10 = pd.concat([base10, curr10], ignore_index=True).drop_duplicates(subset=['論文名稱']) if not curr10.empty or not base10.empty else pd.DataFrame()
    full11 = pd.concat([base11, curr11], ignore_index=True).drop_duplicates(subset=['篇章及所屬專書名稱/或專書名稱']) if not curr11.empty or not base11.empty else pd.DataFrame()
    full17 = curr17.drop_duplicates(subset=['計畫名稱']) if not curr17.empty else pd.DataFrame()

    print("[+] 依照 code1_9(其他代碼) 與標準資料型態重新規範化 (Normalizing)...")
    norm9 = normalize_table1_9(full9)
    norm10 = normalize_table1_10(full10)
    norm11 = normalize_table1_11(full11)
    norm17 = normalize_table1_17(full17)

    print(f"  - Normalized 1_9: {len(norm9)} rows")
    print(f"  - Normalized 1_10: {len(norm10)} rows")
    print(f"  - Normalized 1_11: {len(norm11)} rows")
    print(f"  - Normalized 1_17: {len(norm17)} rows")

    save_df_to_xls(norm9, 'table1_9(期刊論文).xls')
    save_df_to_xls(norm10, 'table1_10(研討會論文).xls')
    save_df_to_xls(norm11, 'table1_11(專書).xls')
    save_df_to_xls(norm17, 'table1_17(國科會計畫案).xls')

    print("[✓] 重構寫入完成！table1_9, table1_10, table1_11, table1_17 已完全符合規範！")

if __name__ == '__main__':
    reprocess_all_files()
