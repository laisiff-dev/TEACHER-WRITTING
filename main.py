import sys
import os
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argparse
import re
from nstc_scraper import (
    search_teachers,
    scrape_teacher_details,
    categorize_publications,
    save_to_3_formats,
    update_workspace_excel_files
)

import threading
import webbrowser

def open_browser(url):
    try:
        webbrowser.open(url)
    except Exception:
        pass

def run_web_server(initial_port=5000):
    from app import app
    ports_to_try = [initial_port, 8080, 8000, 5001, 8888]
    for port in ports_to_try:
        try:
            url = f"http://127.0.0.1:{port}"
            print(f"=" * 60)
            print(f" 網頁系統介面已成功啟動！")
            print(f" 請開啟 Chrome / Edge 瀏覽器輸入網址存取:")
            print(f"   主要網址: {url}")
            print(f"   備用網址: http://localhost:{port}")
            print(f" 觀察黑色主控台視窗中的提示網址： {url}")
            print(f"=" * 60)
            threading.Timer(1.5, open_browser, [url]).start()
            app.run(host='0.0.0.0', port=port, debug=False)
            break
        except OSError as e:
            if "Address already in use" in str(e) or "10048" in str(e) or "only one usage" in str(e):
                print(f"[!] Port {port} 已被其他程式佔用，自動切換至下一個 Port...")
                continue
            else:
                raise e

def main():
    parser = argparse.ArgumentParser(
        description="國科會研究人員著作抓取與匯入系統 (NSTC Talent Scraper & Exporter)"
    )
    parser.add_argument("--search", "-s", type=str, help="搜尋教師姓名或關鍵字 (例如: --search '陳立仁')")
    parser.add_argument("--organ", "-o", type=str, default="", help="服務機關名稱 (例如: --organ '慈濟')")
    parser.add_argument("--url", "-u", type=str, help="國科會研究人員網頁網址或 rsNo")
    parser.add_argument("--web", "-w", action="store_true", help="啟動 Web 圖形化操作介面 (Flask Dashboard)")
    parser.add_argument("--port", "-p", type=int, default=5000, help="Web 介面 Port 號 (預設 5000)")
    
    args = parser.parse_args()

    if args.web:
        run_web_server(args.port)
        return

    if args.url:
        m = re.search(r'rsNo=([a-f0-9]+)', args.url)
        rs_no = m.group(1) if m else args.url.strip()
        print(f"[+] 正在擷取指定研究人員資料 (rsNo: {rs_no})...")
        t_info, pubs, projs = scrape_teacher_details(rs_no)
        process_and_export([t_info], pubs, projs)
        return

    if args.search or args.organ:
        keyword = args.search or ""
        print(f"[+] 正在國科會人才庫搜尋: 關鍵字='{keyword}', 機關='{args.organ}'...")
        teachers = search_teachers(keyword=keyword, organ_desc=args.organ, page_size=20)
        
        if not teachers:
            print("[-] 未搜尋到符合條件的研究人員。")
            return
            
        print(f"[+] 找到 {len(teachers)} 位研究人員:")
        for idx, t in enumerate(teachers, 1):
            print(f"  [{idx}] {t['name_chi']} {t['name_eng']} (rsNo: {t['rsNo']})")

        all_teachers = []
        all_pubs = []
        all_projs = []
        
        for t in teachers:
            print(f" -> 抓取著作與國科會計畫案目錄: {t['name_chi']}...")
            t_info, pubs, projs = scrape_teacher_details(t['rsNo'])
            all_teachers.append(t_info)
            all_pubs.extend(pubs)
            all_projs.extend(projs)

        process_and_export(all_teachers, all_pubs, all_projs)
        return

    # Default if no arguments provided: print usage & run interactive prompt
    print("=" * 60)
    print(" 輔英科技大學國科會教師著作與計畫案檢索系統 ")
    print(" 網頁網址: https://arspb.nstc.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initSearchList&LANG=chi")
    print("=" * 60)
    print("使用方式:")
    print("  1. 啟動 Web UI Dashboard:   python main.py --web")
    print("  2. CLI 搜尋抓取教師著作與計畫: python main.py --search '教師姓名'")
    print("  3. 指定 URL 或 rsNo 抓取:     python main.py --url '網頁連結'")
    print("-" * 60)
    
    choice = input("是否立即搜尋並抓取專任教師對照資料？ (y/n) [預設 y]: ").strip().lower()
    if choice in ['', 'y', 'yes']:
        print("[+] 執行預設 輔英專任教師名冊 對照與匯入任務 (包含國科會計畫案，自動去除其它機構與兼任人員)...")
        res = cross_match_with_official_roster(filepath='table1_1_List(老師清單).xls', max_teachers=10, full_time_only=True)
        print(f"[✓] 任務完成！狀態: {res['updated']}")

def process_and_export(teachers, pubs, projs=None):
    projs = projs or []
    print(f"\n[+] 著作與計畫案擷取完成，共 {len(pubs)} 筆著作資料、{len(projs)} 筆國科會計畫案。")
    cat_dfs = categorize_publications(pubs, projects=projs)
    
    # Save into formats
    output_dir = 'output'
    saved_files = save_to_3_formats(teachers, pubs, cat_dfs, output_dir=output_dir, all_raw_projects=projs)
    print(f"\n[OK] 分別存入資料格式完成 (儲存至 {output_dir}/ 目錄):")
    for sf in saved_files:
        print(f"   - {sf}")

    # Import into the 4 Excel files in workspace folder
    updated = update_workspace_excel_files(cat_dfs, target_dir='.')
    print(f"\n[OK] 統一滙入資料夾之4種 EXCEL 檔案完成:")
    for filename, count in updated.items():
        print(f"   - {filename}: 現有總筆數 {count}")

if __name__ == '__main__':
    main()
