import requests
import re
import os
import sys
import json
import time
import datetime
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import urllib3
import xlwt

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# Suppress SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://arspb.nstc.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do"
ALT_BASE_URL = "https://wrs.nstc.gov.tw/modules/talentSearch/talentSearch.do"
DEFAULT_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
}

# --- Code Mapping Tables (code1_9 & Format Specs) ---
CODE_AUTHOR_ORDER = {
    '第一作者': 'A',
    '第二作者': 'B',
    '第三作者': 'C',
    '第四(以上)作者': 'D'
}

CODE_CROSS_NATION = {
    '是(跨國合作)': '1',
    '是(大陸、港澳地區合作)': '2',
    '是(跨國及大陸、港澳地區合作)': '3',
    '否': '4'
}

# Country Code Dictionary
COUNTRY_MAP = {
    'TAIWAN': 'NATTWN',
    'ROC': 'NATTWN',
    'NATTWN': 'NATTWN',
    'USA': 'NATUSA',
    'UNITED STATES': 'NATUSA',
    'NATUSA': 'NATUSA',
    'JAPAN': 'NATJPN',
    'NATJPN': 'NATJPN',
    'KOREA': 'NATKOR',
    'NATKOR': 'NATKOR',
    'UK': 'NATGBR',
    'UNITED KINGDOM': 'NATGBR',
    'NATGBR': 'NATGBR',
    'AUSTRALIA': 'NATAUS',
    'NATAUS': 'NATAUS',
    'INDONESIA': 'NATIDN',
    'NATIDN': 'NATIDN',
    'FINLAND': 'NATFIN',
    'NATFIN': 'NATFIN'
}

def parse_pub_date(date_str):
    """
    Parses publication date string into (year, month).
    e.g. '2019-04' -> ('2019', '04')
    e.g. '2022' -> ('2022', '')
    """
    if not date_str:
        return '', ''
    date_str = date_str.strip()
    m = re.search(r'(\d{4})(?:[-/](\d{1,2}))?', date_str)
    if m:
        year = m.group(1)
        month = str(int(m.group(2))) if m.group(2) else ''
        return year, month
    return '', ''

def parse_journal_source(source_str):
    """
    Extracts journal name, volume, issue from source string.
    """
    if not source_str:
        return '', '', ''
    
    source_str = source_str.strip()
    journal_name = source_str
    volume = ''
    issue = ''
    
    m = re.search(r'^(.*?)(?:\s+(\d+))?(?:\((\d+)\))?(?::|\s*,|\s*$)', source_str)
    if m:
        j_candidate = m.group(1).strip()
        if j_candidate and not j_candidate.isdigit():
            journal_name = j_candidate
        if m.group(2):
            volume = m.group(2)
        if m.group(3):
            issue = m.group(3)
            
    return journal_name, volume, issue

def map_author_order(author_str, teacher_name):
    """
    Maps author order into code A, B, C, or D based on code1_9 specification.
    """
    if not author_str:
        return 'A'
    authors = [a.strip() for a in re.split(r'[,;&]| and ', author_str) if a.strip()]
    if not authors:
        return 'A'
    
    for idx, a in enumerate(authors):
        if teacher_name and teacher_name in a:
            if idx == 0: return 'A'
            elif idx == 1: return 'B'
            elif idx == 2: return 'C'
            else: return 'D'
            
    if len(authors) == 1:
        return 'A'
    return 'D'

def map_indexing_code(source_str):
    """
    Maps journal indexing type according to code1_9:
    SCI, SSCI, AHCI, TSSCI, EI, SCIE, or O (Other)
    """
    if not source_str:
        return 'O'
    s = source_str.upper()
    if 'SCIE' in s:
        return 'SCIE'
    elif 'SCI' in s and 'TSSCI' not in s and 'SSCI' not in s:
        return 'SCI'
    elif 'SSCI' in s:
        return 'SSCI'
    elif 'AHCI' in s or 'A&HCI' in s:
        return 'AHCI'
    elif 'TSSCI' in s:
        return 'TSSCI'
    elif 'EI' in s:
        return 'EI'
    else:
        return 'O'

def map_country_code(source_str):
    """
    Maps country name/string to standard country code (e.g. NATTWN, NATUSA).
    """
    if not source_str:
        return 'NATTWN'
    s = source_str.upper()
    for k, v in COUNTRY_MAP.items():
        if k in s:
            return v
    return 'NATTWN'

def map_language_code(text):
    """
    Maps title/text language to 'C' (Chinese), 'E' (English), or 'J' (Japanese).
    """
    if not text:
        return 'C'
    if re.search(r'[\u4e00-\u9fa5]', text):
        return 'C'
    return 'E'

def search_teachers(keyword="", organ_desc="", page=1, page_size=50, full_time_only=False):
    """
    Searches NSTC talent database by teacher name, keyword, or organization (e.g., 輔英科技大學).
    When full_time_only=True, cross-references against table1_1_List(老師清單).xls (輔英專任教師名冊)
    and filters out non-Fooyin / non-full-time personnel.
    """
    params = {'action': 'initSearchList', 'LANG': 'chi'}
    data = {
        'nameChi': keyword if keyword else '',
        'organDesc': organ_desc if organ_desc else '',
        'academicExpertiseFullSearch': keyword if not keyword and organ_desc else '',
        'isSearch': '1',
        'LANG': 'chi',
        'currentPage': str(page),
        'pageSize': str(page_size)
    }
    
    try:
        resp = requests.post(BASE_URL, params=params, data=data, headers=DEFAULT_HEADERS, verify=False, timeout=15)
        if resp.status_code != 200:
            print(f"Search request failed with status code {resp.status_code}")
            return []
        
        soup = BeautifulSoup(resp.text, 'html.parser')
        teachers = []
        
        for a in soup.find_all('a'):
            href = a.get('href', '')
            if 'action=initBasic' in href and 'rsNo=' in href:
                rs_no_m = re.search(r'rsNo=([a-f0-9]+)', href)
                if rs_no_m:
                    rs_no = rs_no_m.group(1)
                    raw_text = a.get_text(strip=True)
                    
                    m_name = re.match(r'([\u4e00-\u9fa5]+)(.*)', raw_text)
                    name_chi = m_name.group(1) if m_name else raw_text
                    name_eng = m_name.group(2).strip() if m_name else ""
                    
                    if not any(t['rsNo'] == rs_no for t in teachers):
                        teachers.append({
                            'rsNo': rs_no,
                            'name_chi': name_chi,
                            'name_eng': name_eng,
                            'raw_name': raw_text
                        })

        # Load 輔英專任教師名冊 for metadata binding and optional filtering
        roster_list, roster_map_name, roster_map_id = load_official_teacher_list('table1_1_List(老師清單).xls', full_time_only=True)
        if roster_map_name and teachers:
            filtered_teachers = []
            for t in teachers:
                name = t['name_chi']
                if name in roster_map_name:
                    matched_info = roster_map_name[name]
                    t['fy_id'] = matched_info['id']
                    t['dept'] = matched_info['dept']
                    t['job_type'] = matched_info['job_type']
                    filtered_teachers.append(t)
                else:
                    if full_time_only:
                        print(f"[-] 自動去除非輔英專任人員: {name} (非屬 280 位專任教師名冊)")
                    else:
                        t['dept'] = '兼任/非名冊人員'
                        filtered_teachers.append(t)
            if full_time_only:
                teachers = [t for t in filtered_teachers if t.get('job_type') == '專任']
            else:
                teachers = filtered_teachers

        return teachers
    except Exception as e:
        print(f"Error searching teachers: {e}")
        return []

def search_fooyin_teachers(department="", keyword=""):
    """
    Searches faculty members specifically at Fooyin University (輔英科技大學).
    """
    print(f"[+] 正在檢索 輔英科技大學 教師資料 (系所: '{department}', 關鍵字: '{keyword}')...")
    organ_desc = "輔英科技大學"
    if department:
        organ_desc += f" {department}"
    return search_teachers(keyword=keyword, organ_desc=organ_desc, page_size=100)

def scrape_teacher_projects(rs_no, teacher_name=""):
    """
    Scrapes research projects (國科會計畫案) for a given rsNo from initRsm17new.
    """
    params = {'action': 'initRsm17new', 'rsNo': rs_no, 'LANG': 'chi'}
    projects = []
    
    for target_url in [BASE_URL, ALT_BASE_URL]:
        try:
            resp = requests.get(target_url, params=params, headers=DEFAULT_HEADERS, verify=False, timeout=15)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, 'html.parser')
                for table in soup.find_all('table'):
                    rows = table.find_all('tr')
                    if len(rows) <= 1:
                        continue
                    header = [clean_excel_str(re.sub(r'\s+', ' ', c.get_text())) for c in rows[0].find_all(['td', 'th'])]
                    if any('計畫' in h for h in header):
                        for tr in rows[1:]:
                            cells = [clean_excel_str(re.sub(r'\s+', ' ', c.get_text())) for c in tr.find_all(['td', 'th'])]
                            if len(cells) >= 4:
                                year_raw = cells[0]
                                category = cells[1] if len(cells) > 1 else ''
                                discipline = cells[2] if len(cells) > 2 else ''
                                title = cells[3] if len(cells) > 3 else ''
                                role = cells[4] if len(cells) > 4 else ''
                                budget = cells[5] if len(cells) > 5 else ''
                                
                                if title and '計畫名稱' not in title and '補助類別' not in title:
                                    projects.append({
                                        'rsNo': rs_no,
                                        'teacher_name': teacher_name,
                                        'year': year_raw,
                                        'category': category,
                                        'discipline': discipline,
                                        'title': title,
                                        'role': role,
                                        'budget': budget
                                    })
                if projects:
                    break
        except Exception as e:
            print(f"Error fetching projects for {rs_no} from {target_url}: {e}")
            
    return projects

def scrape_teacher_details(rs_no, teacher_name=None, teacher_id=None):
    """
    Scrapes researcher basic info, publication catalog, and research projects for a given rsNo or Teacher ID.
    Returns (teacher_info, publications, projects).
    """
    actual_rs_no = rs_no
    fy_id = teacher_id
    t_name = teacher_name
    
    if not re.match(r'^[a-f0-9]{32}$', str(rs_no), re.I):
        roster_list, roster_map_name, roster_map_id = load_official_teacher_list('table1_1_List(老師清單).xls', full_time_only=False)
        if str(rs_no) in roster_map_id:
            item = roster_map_id[str(rs_no)]
            fy_id = item['id']
            t_name = item['name']
        elif str(rs_no) in roster_map_name:
            item = roster_map_name[str(rs_no)]
            fy_id = item['id']
            t_name = item['name']
            
        if t_name:
            matches = search_teachers(keyword=t_name, page_size=5, full_time_only=False)
            if matches:
                actual_rs_no = matches[0]['rsNo']
            else:
                return {'rsNo': fy_id or rs_no, 'name_chi': t_name or '', 'name_eng': '', 'org': '國科會未查獲', 'title': '', 'fy_id': fy_id}, [], []

    teacher_info = {
        'rsNo': actual_rs_no,
        'fy_id': fy_id,
        'name_chi': t_name or '',
        'name_eng': '',
        'org': '',
        'title': ''
    }
    
    # 1. Fetch Basic Info
    basic_params = {'action': 'initBasic', 'rsNo': actual_rs_no, 'LANG': 'chi'}
    for target_url in [BASE_URL, ALT_BASE_URL]:
        try:
            r_basic = requests.get(target_url, params=basic_params, headers=DEFAULT_HEADERS, verify=False, timeout=15)
            if r_basic.status_code == 200:
                soup_b = BeautifulSoup(r_basic.text, 'html.parser')
                for tr in soup_b.find_all('tr'):
                    tds = [td.get_text(strip=True) for td in tr.find_all('td')]
                    if len(tds) == 2:
                        key, val = tds[0], tds[1]
                        if key == '中文姓名': teacher_info['name_chi'] = val
                        elif key == '英文姓名': teacher_info['name_eng'] = val
                        elif key == '服務機關': teacher_info['org'] = val
                        elif key == '職稱': teacher_info['title'] = val
                if teacher_info['name_chi'] or teacher_info['org']:
                    break
        except Exception as e:
            print(f"Error fetching basic info for {rs_no}: {e}")

    # 2. Fetch Publication Catalog
    pub_params = {'action': 'initRsm05', 'rsNo': rs_no, 'LANG': 'chi'}
    publications = []
    
    for target_url in [BASE_URL, ALT_BASE_URL]:
        try:
            r_pub = requests.get(target_url, params=pub_params, headers=DEFAULT_HEADERS, verify=False, timeout=15)
            if r_pub.status_code == 200:
                soup_p = BeautifulSoup(r_pub.text, 'html.parser')
                table = soup_p.find('table')
                if table:
                    rows = table.find_all('tr')
                    for tr in rows[1:]:
                        tds = [td.get_text(strip=True) for td in tr.find_all('td')]
                        if len(tds) >= 5:
                            date_raw = tds[0]
                            category = tds[1]
                            title = tds[2]
                            authors = tds[3]
                            source = tds[4]
                            
                            year, month = parse_pub_date(date_raw)
                            
                            publications.append({
                                'rsNo': rs_no,
                                'teacher_name': teacher_info['name_chi'] or teacher_info['name_eng'] or rs_no,
                                'pub_date': date_raw,
                                'year': year,
                                'month': month,
                                'category': category,
                                'title': title,
                                'authors': authors,
                                'source': source
                            })
                    if publications:
                        break
        except Exception as e:
            print(f"Error fetching publications for {rs_no}: {e}")
            
    # 3. Fetch Research Projects (initRsm17new)
    projects = scrape_teacher_projects(rs_no, teacher_name=teacher_info['name_chi'] or teacher_info['name_eng'] or rs_no)
        
    return teacher_info, publications, projects

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

def extract_isbn(text):
    """
    Extracts ISBN-10 or ISBN-13 number from title/source text.
    """
    if not text or pd.isna(text):
        return ''
    m = re.search(r'(?:ISBN[:\s]*)?(97[89][-\s]?\d{1,5}[-\s]?\d{1,7}[-\s]?\d{1,7}[-\s]?[\dX])', str(text), re.I)
    return m.group(1).strip() if m else ''

def verify_and_filter_teacher_ids(df, roster_map_id=None):
    """
    Strictly verifies teacher identity code (身分編碼) against 輔英專任教師名冊 (280位專任教師).
    Removes records with non-full-time staff or invalid fake hash IDs.
    """
    if df is None or df.empty or '身分編碼' not in df.columns:
        return df
    if roster_map_id is None:
        _, _, roster_map_id = load_official_teacher_list('table1_1_List(老師清單).xls', full_time_only=True)
    if not roster_map_id:
        return df
    return df[df['身分編碼'].astype(str).isin(roster_map_id)].copy()

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
    
    # Try extracting ISBN from book title or publisher if missing
    for idx, row in df.iterrows():
        curr_isbn = str(row.get('ISBN編號', '')).strip()
        if not curr_isbn or curr_isbn.lower() == 'nan':
            ext = extract_isbn(f"{row.get('篇章及所屬專書名稱/或專書名稱', '')} {row.get('出版社/出版處所', '')}")
            if ext:
                df.at[idx, 'ISBN編號'] = ext

    # Strict Rule: 專書必須要有 ISBN 號碼才能計入 (Filter out rows without ISBN)
    isbn_s = df['ISBN編號'].astype(str).str.strip().str.lower()
    df = df[df['ISBN編號'].notna() & (isbn_s != '') & (isbn_s != 'nan') & (isbn_s != 'none')].copy()

    df['是否為跨國(地區)合作'] = pd.to_numeric(df['是否為跨國(地區)合作'], errors='coerce').fillna(4).astype(int)
    
    return df

def normalize_table1_17(df):
    """
    Normalizes NSTC Research Projects (國科會計畫案) into standard Table 1_17 format.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=['身分編碼', '教師姓名', '年度', '西元年', '補助類別', '學門代碼', '計畫名稱', '工作', '核定經費'])
    
    col_map = {
        'rsNo': '身分編碼',
        'teacher_name': '教師姓名',
        'year': '年度',
        'category': '補助類別',
        'discipline': '學門代碼',
        'title': '計畫名稱',
        'role': '工作',
        'budget': '核定經費'
    }
    df = df.rename(columns=col_map)
    
    cols = ['身分編碼', '教師姓名', '年度', '西元年', '補助類別', '學門代碼', '計畫名稱', '工作', '核定經費']
    for c in cols:
        if c not in df.columns:
            df[c] = np.nan
    df = df[cols].copy()
    
    df['身分編碼'] = df['身分編碼'].apply(format_fooyin_teacher_id)
    df['教師姓名'] = df['教師姓名'].fillna('').apply(clean_excel_str)
    df['年度'] = df['年度'].fillna('').astype(str).str.strip()
    
    def to_ce_year(y_str):
        if not y_str or pd.isna(y_str): return None
        try:
            m = re.search(r'\d+', str(y_str))
            if m:
                y = int(m.group(0))
                return (y + 1911) if y < 1000 else y
            return None
        except:
            return None

    df['西元年'] = df['年度'].apply(to_ce_year)
    df['補助類別'] = df['補助類別'].fillna('').apply(clean_excel_str)
    df['學門代碼'] = df['學門代碼'].fillna('').apply(clean_excel_str)
    df['計畫名稱'] = df['計畫名稱'].fillna('').apply(clean_excel_str)
    df['工作'] = df['工作'].fillna('計畫主持人').apply(clean_excel_str)
    df['核定經費'] = df['核定經費'].fillna('').astype(str).str.strip()
    
    df = df.drop_duplicates(subset=['身分編碼', '年度', '計畫名稱'])
    return df

def categorize_publications(publications, projects=None):
    """
    Categorizes scraped publication dicts and projects into 4 datasets matching exact format specifications:
    - table1_9 (期刊論文)
    - table1_10 (研討會論文)
    - table1_11 (專書)
    - table1_17 (國科會計畫案)
    """
    list_journal = []
    list_conference = []
    list_book = []
    
    df_projects = normalize_table1_17(pd.DataFrame(projects)) if projects else normalize_table1_17(pd.DataFrame())
    
    for pub in publications:
        cat = pub['category']
        rs_no = pub['rsNo']
        teacher_name = pub.get('teacher_name', '')
        
        author_order_code = map_author_order(pub['authors'], teacher_name)
        
        matched_proj_name = '無'
        if not df_projects.empty:
            teacher_projs = df_projects[df_projects['身分編碼'] == format_fooyin_teacher_id(rs_no)]
            if not teacher_projs.empty:
                for _, p_row in teacher_projs.iterrows():
                    p_title = p_row['計畫名稱']
                    if p_title and len(p_title) > 3:
                        if p_title in pub['title'] or pub['title'] in p_title:
                            matched_proj_name = p_title
                            break
        
        if '期刊' in cat:
            j_name, volume, issue = parse_journal_source(pub['source'])
            indexing_code = map_indexing_code(pub['source'])
            country_code = map_country_code(pub['source'])
            
            list_journal.append({
                '身分編碼': rs_no,
                '所屬計畫案': matched_proj_name,
                '論文名稱': pub['title'],
                '作者順序': author_order_code,
                '通訊作者': '否',
                '收錄分類': indexing_code,
                '刊物名稱': j_name,
                '發表卷數': volume if volume else None,
                '是否具有審稿制度': '是',
                '發表期數': issue if issue else None,
                '期刊出版地國別/地區': country_code,
                '發表月份': int(pub['month']) if pub['month'].isdigit() else None,
                '發表年份': int(pub['year']) if pub['year'].isdigit() else None,
                '發表型式': '電子期刊',
                '是否為跨國(地區)合作': 4
            })
        elif '研討會' in cat or '會議' in cat:
            country_code = map_country_code(pub['source'])
            start_date = ""
            end_date = ""
            if pub['year'].isdigit() and pub['month'].isdigit():
                m_str = str(pub['month']).zfill(2)
                start_date = f"{pub['year']}/{m_str}/01"
                end_date = f"{pub['year']}/{m_str}/02"
                
            list_conference.append({
                '身分編碼': rs_no,
                '所屬計畫案': matched_proj_name,
                '論文名稱': pub['title'],
                '是否具有對外公開徵稿及審稿制度': '是',
                '作者順序': author_order_code,
                '通訊作者': '是',
                '研討會名稱': pub['source'],
                '舉行之國家': country_code,
                '舉行之城市': '台北' if country_code == 'NATTWN' else '海外',
                '開始日期': start_date,
                '結束日期': end_date,
                '發表年份': int(pub['year']) if pub['year'].isdigit() else None
            })
        else:
            lang_code = map_language_code(pub['title'])
            extracted_isbn = extract_isbn(f"{pub['title']} {pub['source']}")
            list_book.append({
                '身分編碼': rs_no,
                '所屬計劃案': matched_proj_name,
                '篇章及所屬專書名稱/或專書名稱': pub['title'],
                '專書類別': '紙本' if '紙本' in cat else ('專書篇章' if '篇章' in cat else '其他'),
                '是否為專書': '是' if '篇章' not in cat else '否',
                '是否有外部審稿程序或公開發行出版': '是',
                '作者順序': author_order_code,
                '通訊作者': '是',
                '使用語文': lang_code,
                '出版年(yyyy)': int(pub['year']) if pub['year'].isdigit() else None,
                '出版月': int(pub['month']) if pub['month'].isdigit() else None,
                '出版社/出版處所': pub['source'],
                'ISBN編號': extracted_isbn,
                '是否為跨國(地區)合作': 4
            })
            
    df_journal = normalize_table1_9(pd.DataFrame(list_journal))
    df_conference = normalize_table1_10(pd.DataFrame(list_conference))
    df_book = normalize_table1_11(pd.DataFrame(list_book))

    return {
        '期刊論文': df_journal,
        '研討會論文': df_conference,
        '專書': df_book,
        '國科會計畫案': df_projects
    }

def filter_by_year_range(df, year_col, start_year=None, end_year=None):
    """
    Filters a DataFrame by publication/project year range.
    Supports both 4-digit CE years (e.g. 2024) and 3-digit ROC years (e.g. 113 -> 2024).
    """
    if df is None or df.empty or year_col not in df.columns:
        return df
    
    res = df.copy()
    years = pd.to_numeric(res[year_col], errors='coerce')
    years = years.apply(lambda y: (y + 1911) if (pd.notna(y) and y < 1000) else y)
    
    if start_year is not None and str(start_year).strip() != '':
        try:
            sy = int(start_year)
            if sy < 1000: sy += 1911
            res = res[years >= sy]
            years = years[years >= sy]
        except (ValueError, TypeError):
            pass
            
    if end_year is not None and str(end_year).strip() != '':
        try:
            ey = int(end_year)
            if ey < 1000: ey += 1911
            res = res[years <= ey]
        except (ValueError, TypeError):
            pass
            
    return res

def filter_cat_dfs_by_year(cat_dfs, start_year=None, end_year=None):
    """
    Filters all 4 publication/project category DataFrames by start_year and end_year.
    """
    if not cat_dfs:
        return cat_dfs
        
    res_dfs = {}
    if '期刊論文' in cat_dfs:
        res_dfs['期刊論文'] = filter_by_year_range(cat_dfs['期刊論文'], '發表年份', start_year, end_year)
    if '研討會論文' in cat_dfs:
        res_dfs['研討會論文'] = filter_by_year_range(cat_dfs['研討會論文'], '發表年份', start_year, end_year)
    if '專書' in cat_dfs:
        res_dfs['專書'] = filter_by_year_range(cat_dfs['專書'], '出版年(yyyy)', start_year, end_year)
    if '國科會計畫案' in cat_dfs:
        res_dfs['國科會計畫案'] = filter_by_year_range(cat_dfs['國科會計畫案'], '西元年', start_year, end_year)
        
    return res_dfs

def clean_excel_str(val):
    """
    Removes illegal XML/control characters that cause openpyxl IllegalCharacterError.
    """
    if isinstance(val, str):
        return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]', '', val).strip()
    return val

def save_to_3_formats(teachers_info, all_raw_pubs, cat_dfs, output_dir='output', all_raw_projects=None):
    """
    Saves scraped data into 3 formats:
    1. JSON (scraped_publications.json - including teachers, publications, projects)
    2. CSV (table1_9, table1_10, table1_11, table1_17)
    3. Excel XLSX (scraped_publications.xlsx with 4 sheets)
    """
    os.makedirs(output_dir, exist_ok=True)
    saved_files = []
    
    # Format 1: JSON
    json_path = os.path.join(output_dir, 'scraped_publications.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump({
            'teachers': teachers_info,
            'publications': all_raw_pubs,
            'projects': all_raw_projects or [],
            'updated_at': datetime.datetime.now().isoformat()
        }, f, ensure_ascii=False, indent=2)
    saved_files.append(json_path)

    # Format 2: CSV (4 separate files with UTF-8 BOM)
    csv_paths = {}
    file_tuples = [
        ('期刊論文', 'table1_9(期刊論文).csv'),
        ('研討會論文', 'table1_10(研討會論文).csv'),
        ('專書', 'table1_11(專書).csv'),
        ('國科會計畫案', 'table1_17(國科會計畫案).csv')
    ]
    for key, name in file_tuples:
        if key in cat_dfs:
            cpath = os.path.join(output_dir, name)
            df_clean = cat_dfs[key].map(clean_excel_str)
            df_clean.to_csv(cpath, index=False, encoding='utf-8-sig')
            saved_files.append(cpath)
            csv_paths[key] = cpath

    # Format 3: Excel XLSX
    xlsx_path = os.path.join(output_dir, 'scraped_publications.xlsx')
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
        for key, df in cat_dfs.items():
            df_clean = df.map(clean_excel_str)
            df_clean.to_excel(writer, sheet_name=key, index=False)
    saved_files.append(xlsx_path)

    return saved_files

def save_df_to_xls(df, filename):
    """
    Saves a DataFrame to a BIFF8 .xls file using xlwt directly.
    """
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Sheet1')
    
    # Write headers
    for col_idx, col_name in enumerate(df.columns):
        ws.write(0, col_idx, str(col_name))
        
    # Write data rows
    for row_idx, row in enumerate(df.values, start=1):
        for col_idx, val in enumerate(row):
            ws.write(row_idx, col_idx, '' if pd.isna(val) else str(val))
            
    wb.save(filename)

def update_workspace_excel_files(cat_dfs, target_dir='.'):
    """
    Appends newly scraped publication and project data into the 4 existing Excel files in the target directory:
    - table1_9(期刊論文).xls
    - table1_10(研討會論文).xls
    - table1_11(專書).xls
    - table1_17(國科會計畫案).xls
    Ensures 100% compliance with code1_9 and baseline reference format files.
    """
    file_map = {
        '期刊論文': ('table1_9(期刊論文).xls', normalize_table1_9, '論文名稱'),
        '研討會論文': ('table1_10(研討會論文).xls', normalize_table1_10, '論文名稱'),
        '專書': ('table1_11(專書).xls', normalize_table1_11, '篇章及所屬專書名稱/或專書名稱'),
        '國科會計畫案': ('table1_17(國科會計畫案).xls', normalize_table1_17, '計畫名稱')
    }
    
    updated_info = {}
    
    for key, (filename, normalizer_func, title_col) in file_map.items():
        if key not in cat_dfs:
            continue
        filepath = os.path.join(target_dir, filename)
        new_df = cat_dfs[key]
        
        if os.path.exists(filepath):
            try:
                existing_df = pd.read_excel(filepath)
                combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                combined_df = combined_df.drop_duplicates(subset=[title_col] if title_col in combined_df.columns else None)
            except Exception as e:
                print(f"Error reading existing {filename}: {e}, overwriting with new data.")
                combined_df = new_df
        else:
            combined_df = new_df
            
        # Apply strict normalization and teacher ID verification against 280 full-time roster
        normalized_df = normalizer_func(combined_df)
        normalized_df = verify_and_filter_teacher_ids(normalized_df)
            
        try:
            save_df_to_xls(normalized_df, filepath)
            updated_info[filename] = len(normalized_df)
            print(f"Successfully updated {filepath} (Total rows: {len(normalized_df)})")
        except Exception as e:
            print(f"Error writing to {filepath}: {e}")
            alt_path = os.path.join(target_dir, filename.replace('.xls', '.xlsx'))
            normalized_df.to_excel(alt_path, index=False, engine='openpyxl')
            updated_info[alt_path] = len(normalized_df)

    return updated_info

def scrape_multiple_teachers(rs_nos, max_workers=10):
    """
    Scrapes basic info, publications, and projects for multiple rsNos concurrently using ThreadPoolExecutor.
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed
    
    all_teachers = []
    all_pubs = []
    all_projs = []
    
    def fetch_one(rs_no):
        return scrape_teacher_details(rs_no)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(fetch_one, rs_no): rs_no for rs_no in rs_nos}
        for future in as_completed(futures):
            try:
                t_info, pubs, projs = future.result()
                all_teachers.append(t_info)
                all_pubs.extend(pubs)
                all_projs.extend(projs)
            except Exception as e:
                print(f"Error scraping {futures[future]}: {e}")
                
    return all_teachers, all_pubs, all_projs

def load_official_teacher_list(filepath='table1_1_List(老師清單).xls', full_time_only=True):
    """
    Parses table1_1_List(老師清單).xls file using html.unescape and returns (roster_list, roster_map_name, roster_map_id).
    By default (full_time_only=True), filters strictly for 輔英專任教師 (job_type == '專任', 280 位).
    """
    if not os.path.exists(filepath):
        print(f"Warning: Teacher list file {filepath} not found.")
        return [], {}, {}

    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            raw_text = f.read()
        
        import html
        unescaped_text = html.unescape(raw_text)
        soup = BeautifulSoup(unescaped_text, 'html.parser')
        tr_list = soup.find_all('tr')
        
        rows = []
        for tr in tr_list:
            cells = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
            if cells:
                rows.append(cells)
                
        if len(rows) < 3:
            return [], {}, {}
            
        header = rows[0]
        col_id_idx, col_name_idx, col_dept_idx, col_job_type_idx, col_rank_idx = -1, -1, -1, -1, -1
        
        for idx, col in enumerate(header):
            if '身分編碼' in col or '編碼' in col: col_id_idx = idx
            elif '教師姓名' in col or '姓名' in col: col_name_idx = idx
            elif '主聘系所' in col and '類別' not in col and '代碼' not in col: col_dept_idx = idx
            elif '專兼任' in col: col_job_type_idx = idx
            elif '聘書職級' in col or '職級' in col: col_rank_idx = idx
            
        roster_list = []
        roster_map_name = {}
        roster_map_id = {}
        
        for row in rows[2:]:
            t_id = row[col_id_idx].strip() if col_id_idx >= 0 and col_id_idx < len(row) else ""
            t_name = row[col_name_idx].strip() if col_name_idx >= 0 and col_name_idx < len(row) else ""
            dept = row[col_dept_idx].strip() if col_dept_idx >= 0 and col_dept_idx < len(row) else ""
            job_type = row[col_job_type_idx].strip() if col_job_type_idx >= 0 and col_job_type_idx < len(row) else ""
            rank = row[col_rank_idx].strip() if col_rank_idx >= 0 and col_rank_idx < len(row) else ""
            
            # Filter strictly for 輔英專任教師 if full_time_only is True
            if full_time_only and job_type != '專任':
                continue
                
            if t_id and t_name and t_name.lower() != 'nan':
                item = {
                    'id': t_id,
                    'name': t_name,
                    'dept': dept,
                    'rank': rank,
                    'job_type': job_type
                }
                roster_map_name[t_name] = item
                roster_map_id[t_id] = item
                if not any(r['id'] == t_id for r in roster_list):
                    roster_list.append(item)
                    
        type_str = "輔英專任教師" if full_time_only else "全校教師(含兼任)"
        print(f"[+] 成功讀取 {len(roster_list)} 位{type_str}資料 ({filepath})")
        return roster_list, roster_map_name, roster_map_id
    except Exception as e:
        print(f"Error reading official teacher list {filepath}: {e}")
        return [], {}, {}

def cross_match_with_official_roster(filepath='table1_1_List(老師清單).xls', max_teachers=50, full_time_only=True):
    """
    Cross-matches personnel from NSTC talent search strictly against 輔英專任教師名冊.
    Filters out non-Fooyin personnel and non-full-time staff ("其它機構也需去除").
    Assigns authentic Fooyin Teacher IDs (T-numbers) and imports publications/projects to system.
    """
    roster_list, roster_map_name, roster_map_id = load_official_teacher_list(filepath, full_time_only=full_time_only)
    
    if not roster_list:
        print("[-] 無法讀取老師清單，使用預設輔英專任名冊執行...")
        return cross_match_fooyin_faculty()

    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 啟動與 輔英專任教師名冊 比對機制 (進行前 {min(max_teachers, len(roster_list))} 位專任教師對比，過濾其它機構與兼任人員)...")
    
    match_results = []
    all_teachers = []
    all_pubs = []
    all_projs = []
    
    target_teachers = roster_list[:max_teachers]
    
    for t_item in target_teachers:
        name = t_item['name']
        fy_id = t_item['id']
        dept = t_item['dept']
        
        # Search NSTC for candidate matches
        matches = search_teachers(keyword=name, page_size=10)
        
        # Select best candidate matching Fooyin or roster
        selected_match = None
        if matches:
            fy_matches = [m for m in matches if '輔英' in (m.get('org') or '') or 'Fooyin' in (m.get('org') or '')]
            if fy_matches:
                selected_match = fy_matches[0]
            else:
                selected_match = matches[0]

        if selected_match:
            rs_no = selected_match['rsNo']
            t_info, pubs, projs = scrape_teacher_details(rs_no)
            
            # Bind authentic Fooyin Teacher ID & department
            t_info['fy_id'] = fy_id
            t_info['dept'] = dept
            t_info['rsNo'] = fy_id
            
            for p in pubs:
                p['rsNo'] = fy_id
            for pr in projs:
                pr['rsNo'] = fy_id
                
            all_teachers.append(t_info)
            all_pubs.extend(pubs)
            all_projs.extend(projs)
            
            match_results.append({
                'name': name,
                'fy_id': fy_id,
                'dept': dept,
                'job_type': t_item.get('job_type', '專任'),
                'nstc_org': t_info.get('org', '未填寫'),
                'nstc_title': t_info.get('title', ''),
                'pubs_count': len(pubs),
                'projs_count': len(projs),
                'matched': True
            })
            print(f"  [OK] 專任比對成功: {name} ({fy_id}, {dept}) <-> 國科會登記: '{t_info.get('org')}' (著作: {len(pubs)}, 計畫: {len(projs)})")
        else:
            match_results.append({
                'name': name,
                'fy_id': fy_id,
                'dept': dept,
                'job_type': t_item.get('job_type', '專任'),
                'nstc_org': '國科會未查獲',
                'nstc_title': '',
                'pubs_count': 0,
                'projs_count': 0,
                'matched': False
            })
            print(f"  [-] 國科會未查獲: {name} (ID: {fy_id})")
            
    cat_dfs = categorize_publications(all_pubs, projects=all_projs)
    save_to_3_formats(all_teachers, all_pubs, cat_dfs, output_dir='output', all_raw_projects=all_projs)
    updated = update_workspace_excel_files(cat_dfs, target_dir='.')
    
    return {
        'match_results': match_results,
        'total_teachers': len(all_teachers),
        'total_pubs': len(all_pubs),
        'total_projs': len(all_projs),
        'updated': updated,
        'cat_dfs': cat_dfs
    }

# --- Default Fooyin University Faculty Roster for Cross-Matching ---
DEFAULT_FOOYIN_ROSTER = [
    {'name': '陳立仁', 'id': 'T0000009267', 'dept': '全人教育中心'},
    {'name': '嚴嘉楓', 'id': 'T0000121656', 'dept': '公共衛生系'},
    {'name': '詹道明', 'id': 'T0000054317', 'dept': '藥學系'},
    {'name': '張重義', 'id': 'T0000010059', 'dept': '外科'},
    {'name': '張志麟', 'id': 'T0000038837', 'dept': '牙科學系'},
    {'name': '左星樺', 'id': 'T0000032139', 'dept': '人文與社會學系'},
    {'name': '莊祐中', 'id': 'T0000095727', 'dept': '醫學系'},
    {'name': '林郁伶', 'id': 'T0000036425', 'dept': '環境與生命科學系'},
    {'name': '張家禎', 'id': 'T0000062021', 'dept': '助產與婦幼健康照護系'},
    {'name': '王承舜', 'id': 'T0000001738', 'dept': '資訊科技與管理系'},
    {'name': '謝佳容', 'id': 'T0000084253', 'dept': '護理系'},
    {'name': '張雅鈴', 'id': 'T0000045947', 'dept': '護理系'},
    {'name': '王乃巧', 'id': 'T0000022109', 'dept': '健康美容系'},
    {'name': '高文魁', 'id': 'T0000100320', 'dept': '保健營養系'},
    {'name': '柯瑋妮', 'id': 'T0000095912', 'dept': '護理系'},
    {'name': '徐英九', 'id': 'T0000077888', 'dept': '工業管理系'},
    {'name': '周怡',   'id': 'T0000010908', 'dept': '財經法律系'}
]

def cross_match_fooyin_faculty(roster=None, max_workers=10):
    """
    Cross-matches Fooyin University faculty members against NSTC talent search by name.
    Even if their NSTC profile lists a non-Fooyin organization (e.g., Tzu Chi, NTU, KMU, Mackay),
    retrieves their NSTC publication catalog & research projects, assigns their Fooyin Teacher ID, and imports to system.
    """
    if not roster:
        roster = DEFAULT_FOOYIN_ROSTER

    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 啟動 輔英科技大學跨校/非本校機構比對機制 (共 {len(roster)} 位教師)...")
    
    match_results = []
    all_teachers = []
    all_pubs = []
    all_projs = []
    
    for t_item in roster:
        name = t_item['name']
        fy_id = t_item.get('id', '')
        dept = t_item.get('dept', '')
        
        # Query NSTC by teacher name
        matches = search_teachers(keyword=name, page_size=10)
        if matches:
            selected_match = matches[0]
            rs_no = selected_match['rsNo']
            
            t_info, pubs, projs = scrape_teacher_details(rs_no)
            
            t_info['fy_id'] = fy_id
            t_info['dept'] = dept
            t_info['rsNo'] = fy_id if fy_id else rs_no
            
            for p in pubs:
                p['rsNo'] = fy_id if fy_id else rs_no
            for pr in projs:
                pr['rsNo'] = fy_id if fy_id else rs_no
                
            all_teachers.append(t_info)
            all_pubs.extend(pubs)
            all_projs.extend(projs)
            
            match_results.append({
                'name': name,
                'fy_id': fy_id,
                'dept': dept,
                'nstc_org': t_info.get('org', '未填寫'),
                'nstc_title': t_info.get('title', ''),
                'pubs_count': len(pubs),
                'projs_count': len(projs),
                'matched': True
            })
            print(f"  [✓] 比對成功: {name} (輔英ID: {fy_id}) <-> 國科會登記機構: '{t_info.get('org')}' (著作數: {len(pubs)}, 計畫數: {len(projs)})")
        else:
            match_results.append({
                'name': name,
                'fy_id': fy_id,
                'dept': dept,
                'nstc_org': '國科會未查獲',
                'nstc_title': '',
                'pubs_count': 0,
                'projs_count': 0,
                'matched': False
            })
            print(f"  [-] 未於國科會查獲: {name}")
            
    cat_dfs = categorize_publications(all_pubs, projects=all_projs)
    save_to_3_formats(all_teachers, all_pubs, cat_dfs, output_dir='output', all_raw_projects=all_projs)
    updated = update_workspace_excel_files(cat_dfs, target_dir='.')
    
    return {
        'match_results': match_results,
        'total_teachers': len(all_teachers),
        'total_pubs': len(all_pubs),
        'total_projs': len(all_projs),
        'updated': updated,
        'cat_dfs': cat_dfs
    }

def run_quarterly_update(organ="輔英科技大學", output_dir="output", full_time_only=True):
    """
    Executes a quarterly update process:
    Scrapes 輔英專任教師 (full-time faculty) concurrently, categorizes publications & projects,
    saves 4 data formats, and updates workspace Excel files ("僅帶輔英專任教師，其它機構/兼任人員去除").
    """
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 啟動每季定期抓取與自動更新任務 (對照 輔英專任教師名冊，自動去除其它機構與兼任人員)...")
    res = cross_match_with_official_roster(max_teachers=280, full_time_only=full_time_only)
    print(f"[✓] 每季專任教師著作與國科會計畫案定期更新完成！狀態: {res['updated']}")
    return res['updated'], res.get('all_teachers', []), res.get('all_pubs', []), res.get('all_projs', []), res['cat_dfs']

COLLEGE_MAP = {
    '護理學院': ['護理系', '護理科', '學士後護理系', '助產與婦嬰健康照護系', '高齡及長期照護事業系'],
    '醫學與健康學院': ['醫學檢驗生物技術系', '保健營養系', '物理治療系'],
    '環境與生命學院': ['環境工程與科學系', '生物科技與綠色產業系', '健康美容系', '職業安全衛生系', '應用化學及材料科學系'],
    '人文與管理學院': ['資訊科技與管理系', '健康事業管理系', '休閒與遊憩事業管理系', '幼兒保育暨產業系', '應用外語系', '軍訓教官', '全人教育中心', '共同教育中心']
}

def get_college_by_dept(dept):
    if not dept:
        return '其他單位'
    for c, depts in COLLEGE_MAP.items():
        if any(d in str(dept) for d in depts):
            return c
    return '其他單位'

def compute_college_analytics(cat_dfs=None, target_years=None):
    """
    Computes analytics per College for 112-115年度 (2023-2026) for Journal papers, Conference papers, and NSTC Project budgets (in 萬元).
    """
    if target_years is None:
        target_years = [2023, 2024, 2025, 2026]
        
    target_years = [int(y) for y in target_years]
    
    # Load teacher roster for teacher ID -> College mapping
    roster_list, _, roster_map_id = load_official_teacher_list('table1_1_List(老師清單).xls', full_time_only=False)
    id_to_college = {t_id: get_college_by_dept(info['dept']) for t_id, info in roster_map_id.items()}
    
    colleges = ['護理學院', '醫學與健康學院', '環境與生命學院', '人文與管理學院']
    analytics = {c: {y: {'journal': 0, 'conference': 0, 'budget': 0.0} for y in target_years} for c in colleges}
    
    # Load DataFrames
    df9 = cat_dfs.get('期刊論文') if (cat_dfs and '期刊論文' in cat_dfs) else (pd.read_excel('table1_9(期刊論文).xls') if os.path.exists('table1_9(期刊論文).xls') else pd.DataFrame())
    df10 = cat_dfs.get('研討會論文') if (cat_dfs and '研討會論文' in cat_dfs) else (pd.read_excel('table1_10(研討會論文).xls') if os.path.exists('table1_10(研討會論文).xls') else pd.DataFrame())
    df17 = cat_dfs.get('國科會計畫案') if (cat_dfs and '國科會計畫案' in cat_dfs) else (pd.read_excel('table1_17(國科會計畫案).xls') if os.path.exists('table1_17(國科會計畫案).xls') else pd.DataFrame())
    
    # Process Journals
    if not df9.empty:
        for _, r in df9.iterrows():
            cid = str(r.get('身分編碼', '')).strip()
            col = id_to_college.get(cid)
            y = r.get('發表年份')
            if pd.notna(y):
                try:
                    yi = int(y)
                    yi = (yi + 1911) if yi < 1000 else yi
                    if col in analytics and yi in target_years:
                        analytics[col][yi]['journal'] += 1
                except:
                    pass

    # Process Conferences
    if not df10.empty:
        for _, r in df10.iterrows():
            cid = str(r.get('身分編碼', '')).strip()
            col = id_to_college.get(cid)
            y = r.get('發表年份')
            if pd.notna(y):
                try:
                    yi = int(y)
                    yi = (yi + 1911) if yi < 1000 else yi
                    if col in analytics and yi in target_years:
                        analytics[col][yi]['conference'] += 1
                except:
                    pass

    # Process Projects
    if not df17.empty:
        for _, r in df17.iterrows():
            cid = str(r.get('身分編碼', '')).strip()
            col = id_to_college.get(cid)
            y = r.get('西元年') or r.get('年度')
            if pd.notna(y):
                m = re.search(r'\d+', str(y))
                if m:
                    yi = int(m.group(0))
                    yi = (yi + 1911) if yi < 1000 else yi
                    b_raw = str(r.get('核定經費', '0'))
                    b_clean = re.sub(r'[^\d.]', '', b_raw)
                    try:
                        budget_wan = float(b_clean) / 10000.0 if b_clean else 0.0
                        if col in analytics and yi in target_years:
                            analytics[col][yi]['budget'] += budget_wan
                    except:
                        pass

    # Round budget values
    for c in colleges:
        for y in target_years:
            analytics[c][y]['budget'] = round(analytics[c][y]['budget'], 2)

    return {
        'years': target_years,
        'colleges': colleges,
        'analytics': analytics
    }

if __name__ == '__main__':
    print("Testing Updated NSTC & Fooyin Scraper Module...")
    res = cross_match_fooyin_faculty()
    print(f"Cross-match summary: {res['total_teachers']} teachers, {res['total_pubs']} pubs, {res.get('total_projs', 0)} projs updated.")


