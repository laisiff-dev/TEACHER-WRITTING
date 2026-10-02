import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import json
import zipfile
import io
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd

from nstc_scraper import (
    search_teachers,
    scrape_teacher_details,
    scrape_multiple_teachers,
    categorize_publications,
    save_to_3_formats,
    update_workspace_excel_files,
    run_quarterly_update,
    cross_match_fooyin_faculty,
    load_official_teacher_list,
    cross_match_with_official_roster,
    clean_excel_str,
    filter_cat_dfs_by_year,
    normalize_table1_9,
    normalize_table1_10,
    normalize_table1_11,
    normalize_table1_17
)

app = Flask(__name__, template_folder='templates')

# Cache in-memory active scraped data
LATEST_SCRAPED = {
    'teachers': [],
    'publications': [],
    'projects': [],
    'cat_dfs': None
}

def df_to_clean_dict(df):
    """
    Safely converts a DataFrame to JSON-compliant records by cleaning string control chars
    and replacing NaN/None values with empty strings to prevent invalid 'NaN' JSON tokens.
    """
    if df is None or df.empty:
        return []
    df_clean = df.map(clean_excel_str).fillna('')
    return df_clean.to_dict(orient='records')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/search', methods=['POST'])
def api_search():
    data = request.get_json(force=True, silent=True) or {}
    keyword = data.get('keyword', '').strip()
    organ_desc = data.get('organ_desc', '').strip()
    
    if not organ_desc and not keyword:
        organ_desc = '輔英科技大學'
        
    teachers = search_teachers(keyword=keyword, organ_desc=organ_desc, page_size=100)
    return jsonify({
        'success': True,
        'count': len(teachers),
        'teachers': teachers
    })

@app.route('/api/match_official_roster', methods=['POST'])
def api_match_official_roster():
    """
    Executes cross-matching against table1_1_List(老師清單).xls official roster (輔英專任教師).
    """
    global LATEST_SCRAPED
    data = request.get_json(force=True, silent=True) or {}
    max_teachers = data.get('max_teachers', 280)
    full_time_only = data.get('full_time_only', True)
    
    try:
        res = cross_match_with_official_roster(filepath='table1_1_List(老師清單).xls', max_teachers=max_teachers, full_time_only=full_time_only)
        cat_dfs = res['cat_dfs']
        
        LATEST_SCRAPED['cat_dfs'] = cat_dfs
        
        categorized_dict = {
            'journal': df_to_clean_dict(cat_dfs.get('期刊論文')),
            'conference': df_to_clean_dict(cat_dfs.get('研討會論文')),
            'book': df_to_clean_dict(cat_dfs.get('專書')),
            'project': df_to_clean_dict(cat_dfs.get('國科會計畫案'))
        }
        
        type_desc = "輔英專任教師" if full_time_only else "全校教師(含兼任)"
        return jsonify({
            'success': True,
            'message': f"成功對照 table1_1_List {type_desc}名冊 (其它機構/兼任已過濾)！已完成 {res['total_teachers']} 位專任教師對比，共擷取 {res['total_pubs']} 筆著作及 {res.get('total_projs', 0)} 筆國科會計畫案",
            'match_results': res['match_results'],
            'total_teachers': res['total_teachers'],
            'total_pubs': res['total_pubs'],
            'total_projs': res.get('total_projs', 0),
            'updated': res['updated'],
            'categorized': categorized_dict
        })
    except Exception as e:
        print(f"Error during official roster match: {e}")
        return jsonify({'success': False, 'message': f'官方名冊比對出錯: {str(e)}'}), 500

@app.route('/api/cross_match', methods=['POST'])
def api_cross_match():
    """
    Executes cross-matching between Fooyin faculty roster and NSTC talent search profiles.
    """
    global LATEST_SCRAPED
    data = request.get_json(force=True, silent=True) or {}
    custom_roster = data.get('roster', None)
    
    try:
        res = cross_match_fooyin_faculty(roster=custom_roster)
        cat_dfs = res['cat_dfs']
        
        LATEST_SCRAPED['cat_dfs'] = cat_dfs
        
        categorized_dict = {
            'journal': df_to_clean_dict(cat_dfs.get('期刊論文')),
            'conference': df_to_clean_dict(cat_dfs.get('研討會論文')),
            'book': df_to_clean_dict(cat_dfs.get('專書')),
            'project': df_to_clean_dict(cat_dfs.get('國科會計畫案'))
        }
        
        return jsonify({
            'success': True,
            'message': f"跨校與非本校機構比對完成！成功比對 {res['total_teachers']} 位教師，共擷取 {res['total_pubs']} 筆著作與 {res.get('total_projs', 0)} 筆國科會計畫案",
            'match_results': res['match_results'],
            'total_teachers': res['total_teachers'],
            'total_pubs': res['total_pubs'],
            'total_projs': res.get('total_projs', 0),
            'updated': res['updated'],
            'categorized': categorized_dict
        })
    except Exception as e:
        print(f"Error during cross-matching: {e}")
        return jsonify({'success': False, 'message': f'比對出錯: {str(e)}'}), 500

@app.route('/api/quarterly_sync', methods=['POST'])
def api_quarterly_sync():
    """
    Executes quarterly scheduled sync for 輔英科技大學 faculty members.
    """
    global LATEST_SCRAPED
    try:
        updated, all_teachers, all_pubs, all_projs, cat_dfs = run_quarterly_update(organ="輔英科技大學", output_dir="output")
        
        LATEST_SCRAPED['teachers'] = all_teachers
        LATEST_SCRAPED['publications'] = all_pubs
        LATEST_SCRAPED['projects'] = all_projs
        LATEST_SCRAPED['cat_dfs'] = cat_dfs
        
        categorized_dict = {
            'journal': df_to_clean_dict(cat_dfs.get('期刊論文')),
            'conference': df_to_clean_dict(cat_dfs.get('研討會論文')),
            'book': df_to_clean_dict(cat_dfs.get('專書')),
            'project': df_to_clean_dict(cat_dfs.get('國科會計畫案'))
        }
        
        return jsonify({
            'success': True,
            'message': '每季定期更新執行完成！',
            'updated': updated,
            'teachers_count': len(all_teachers),
            'total_pubs': len(all_pubs),
            'total_projs': len(all_projs),
            'categorized': categorized_dict
        })
    except Exception as e:
        print(f"Error during quarterly sync: {e}")
        return jsonify({'success': False, 'message': f'每季更新出錯: {str(e)}'}), 500

@app.route('/api/scrape', methods=['POST'])
def api_scrape():
    global LATEST_SCRAPED
    data = request.get_json() or {}
    rs_nos = data.get('rs_nos', [])
    
    if not rs_nos:
        return jsonify({'success': False, 'message': '未提供選擇的教師編號'}), 400
        
    all_teachers, all_pubs, all_projs = scrape_multiple_teachers(rs_nos, max_workers=10)
    cat_dfs = categorize_publications(all_pubs, projects=all_projs)
    
    saved_files = save_to_3_formats(all_teachers, all_pubs, cat_dfs, output_dir='output', all_raw_projects=all_projs)
    
    LATEST_SCRAPED['teachers'] = all_teachers
    LATEST_SCRAPED['publications'] = all_pubs
    LATEST_SCRAPED['projects'] = all_projs
    LATEST_SCRAPED['cat_dfs'] = cat_dfs
    
    categorized_dict = {
        'journal': df_to_clean_dict(cat_dfs.get('期刊論文')),
        'conference': df_to_clean_dict(cat_dfs.get('研討會論文')),
        'book': df_to_clean_dict(cat_dfs.get('專書')),
        'project': df_to_clean_dict(cat_dfs.get('國科會計畫案'))
    }
    
    return jsonify({
        'success': True,
        'teachers_count': len(all_teachers),
        'total_pubs': len(all_pubs),
        'total_projs': len(all_projs),
        'categorized': categorized_dict,
        'saved_files': saved_files
    })

@app.route('/api/import_excel', methods=['POST'])
def api_import_excel():
    global LATEST_SCRAPED
    if LATEST_SCRAPED['cat_dfs'] is None:
        try:
            df9 = pd.read_excel('table1_9(期刊論文).xls') if os.path.exists('table1_9(期刊論文).xls') else pd.DataFrame()
            df10 = pd.read_excel('table1_10(研討會論文).xls') if os.path.exists('table1_10(研討會論文).xls') else pd.DataFrame()
            df11 = pd.read_excel('table1_11(專書).xls') if os.path.exists('table1_11(專書).xls') else pd.DataFrame()
            df17 = pd.read_excel('table1_17(國科會計畫案).xls') if os.path.exists('table1_17(國科會計畫案).xls') else pd.DataFrame()
            cat_dfs = {
                '期刊論文': normalize_table1_9(df9),
                '研討會論文': normalize_table1_10(df10),
                '專書': normalize_table1_11(df11),
                '國科會計畫案': normalize_table1_17(df17)
            }
            updated = update_workspace_excel_files(cat_dfs, target_dir='.')
            return jsonify({
                'success': True,
                'message': '從既存 Excel 讀取並重新匯入對齊完成',
                'updated': updated
            })
        except Exception as e:
            return jsonify({'success': False, 'message': f'匯入失敗: {str(e)}'}), 500
        
    updated = update_workspace_excel_files(LATEST_SCRAPED['cat_dfs'], target_dir='.')
    return jsonify({
        'success': True,
        'updated': updated
    })

def ensure_output_files_exist():
    """
    Ensures output/ directory contains scraped_publications.json, scraped_publications.xlsx,
    and 4 CSV files by reading existing workspace Excel files if output files do not exist yet.
    """
    os.makedirs('output', exist_ok=True)
    json_path = os.path.join('output', 'scraped_publications.json')
    xlsx_path = os.path.join('output', 'scraped_publications.xlsx')
    
    if not os.path.exists(json_path) or not os.path.exists(xlsx_path):
        try:
            df9 = pd.read_excel('table1_9(期刊論文).xls') if os.path.exists('table1_9(期刊論文).xls') else pd.DataFrame()
            df10 = pd.read_excel('table1_10(研討會論文).xls') if os.path.exists('table1_10(研討會論文).xls') else pd.DataFrame()
            df11 = pd.read_excel('table1_11(專書).xls') if os.path.exists('table1_11(專書).xls') else pd.DataFrame()
            df17 = pd.read_excel('table1_17(國科會計畫案).xls') if os.path.exists('table1_17(國科會計畫案).xls') else pd.DataFrame()
            
            cat_dfs = {
                '期刊論文': normalize_table1_9(df9),
                '研討會論文': normalize_table1_10(df10),
                '專書': normalize_table1_11(df11),
                '國科會計畫案': normalize_table1_17(df17)
            }
            save_to_3_formats([], [], cat_dfs, output_dir='output')
        except Exception as e:
            print(f"Error auto-generating output files: {e}")

@app.route('/download/json')
def download_json():
    ensure_output_files_exist()
    json_path = os.path.join('output', 'scraped_publications.json')
    if os.path.exists(json_path):
        return send_file(json_path, as_attachment=True, download_name='scraped_publications.json')
    return "JSON 檔案尚未生成", 404

@app.route('/download/xlsx')
def download_xlsx():
    ensure_output_files_exist()
    xlsx_path = os.path.join('output', 'scraped_publications.xlsx')
    if os.path.exists(xlsx_path):
        return send_file(xlsx_path, as_attachment=True, download_name='scraped_publications.xlsx')
    return "Excel 檔案尚未生成", 404

@app.route('/download/csv')
def download_csv():
    ensure_output_files_exist()
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname in ['table1_9(期刊論文).csv', 'table1_10(研討會論文).csv', 'table1_11(專書).csv', 'table1_17(國科會計畫案).csv']:
            fpath = os.path.join('output', fname)
            if os.path.exists(fpath):
                zf.write(fpath, fname)
    memory_file.seek(0)
    return send_file(memory_file, mimetype='application/zip', as_attachment=True, download_name='scraped_csv_files.zip')

@app.route('/download/table1_9')
def download_table1_9():
    if os.path.exists('table1_9(期刊論文).xls'):
        return send_file('table1_9(期刊論文).xls', as_attachment=True, download_name='table1_9(期刊論文).xls')
    return "table1_9 檔案不存在", 404

@app.route('/download/table1_10')
def download_table1_10():
    if os.path.exists('table1_10(研討會論文).xls'):
        return send_file('table1_10(研討會論文).xls', as_attachment=True, download_name='table1_10(研討會論文).xls')
    return "table1_10 檔案不存在", 404

@app.route('/download/table1_11')
def download_table1_11():
    if os.path.exists('table1_11(專書).xls'):
        return send_file('table1_11(專書).xls', as_attachment=True, download_name='table1_11(專書).xls')
    return "table1_11 檔案不存在", 404

@app.route('/download/table1_17')
def download_table1_17():
    if os.path.exists('table1_17(國科會計畫案).xls'):
        return send_file('table1_17(國科會計畫案).xls', as_attachment=True, download_name='table1_17(國科會計畫案).xls')
    return "table1_17 檔案不存在", 404

@app.route('/download/all_xls')
def download_all_xls():
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname in ['table1_9(期刊論文).xls', 'table1_10(研討會論文).xls', 'table1_11(專書).xls', 'table1_17(國科會計畫案).xls']:
            if os.path.exists(fname):
                zf.write(fname, fname)
    memory_file.seek(0)
    return send_file(memory_file, mimetype='application/zip', as_attachment=True, download_name='all_table_xls_files.zip')
@app.errorhandler(500)
def internal_server_error(e):
    print(f"Server 500 Error: {e}")
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'message': f'伺服器內部處理錯誤: {str(e)}'}), 500
    return "伺服器內部錯誤 (500)", 500

@app.errorhandler(404)
def not_found_error(e):
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'message': '請求 API 路徑不存在 (404)'}), 404
    return "找不到頁面 (404)", 404

if __name__ == '__main__':
    print("Starting NSTC Talent Search Web Server on http://127.0.0.1:5000 ...")
    app.run(host='127.0.0.1', port=5000, debug=False)
