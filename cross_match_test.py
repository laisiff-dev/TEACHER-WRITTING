import sys
import os
sys.path.insert(0, '.')
from nstc_scraper import search_teachers, scrape_teacher_details

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

sample_names = ['嚴嘉楓', '張家禎', '林雅晴', '陳立仁', '王承舜', '高文魁', '張雅鈴', '詹道明', '張重義', '王乃巧', '張志麟', '左星樺', '柯瑋妮', '徐英九', '邢智田', '顏澤文', '莊祐中', '周怡', '林郁伶']

print("Cross-matching Fooyin faculty names against NSTC Talent Database...")

for name in sample_names:
    res = search_teachers(keyword=name)
    if res:
        print(f"\n[+] Name: {name} (Found {len(res)} matches on NSTC):")
        for r in res[:2]:
            t_info, pubs = scrape_teacher_details(r['rsNo'])
            print(f"   - {r['name_chi']} {r['name_eng']} | Org: '{t_info.get('org')}' | Title: '{t_info.get('title')}' | Publications: {len(pubs)}")
    else:
        print(f"\n[-] Name: {name} (Not found on NSTC search)")
