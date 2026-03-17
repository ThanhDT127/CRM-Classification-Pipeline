"""
Script trích xuất keywords từ Excel ra JSON
Cấu trúc Excel:
- Cột 0: Nhóm (MAJOR) - merged cells
- Cột 1: Tên Cột (MINOR) - merged cells  
- Cột 2: Nguồn Dữ Liệu (BỎ QUA)
- Cột 3: Trạng Thái Cột (SUB)
- Cột 4: Từ Khóa Liên Quan (KEYWORDS)
"""

import pandas as pd
import json
from pathlib import Path

def extract_keywords_to_json(excel_file, output_json):
    """
    Trích xuất keywords từ Excel theo cấu trúc:
    MAJOR → MINOR → SUB → [keywords]
    """
    
    # Đọc Excel (không có header)
    df = pd.read_excel(excel_file, header=None)
    
    print(f"📊 Đọc {len(df)} dòng từ Excel")
    
    # Đặt tên cột (có thêm cột "Từ khóa bổ sung")
    df.columns = ['major', 'minor', 'source', 'sub', 'keywords', 'keywords_extra'] + [f'col_{i}' for i in range(6, len(df.columns))]
    
    # Bỏ dòng header
    df = df.iloc[1:].reset_index(drop=True)
    
    # Forward fill merged cells
    df['major'] = df['major'].ffill()
    df['minor'] = df['minor'].ffill()
    
    # Build structure
    result = {}
    
    for idx, row in df.iterrows():
        major = str(row['major']).strip()
        minor = str(row['minor']).strip()
        sub = str(row['sub']).strip() if pd.notna(row['sub']) else ""
        keywords_text = str(row['keywords']).strip() if pd.notna(row['keywords']) else ""
        keywords_extra_text = str(row['keywords_extra']).strip() if pd.notna(row['keywords_extra']) else ""
        
        # Skip nan
        if major == 'nan' or not sub or sub == 'nan':
            continue
        
        # Init major
        if major not in result:
            result[major] = {}
        
        # Init minor  
        if minor not in result[major]:
            result[major][minor] = {}
        
        # Parse keywords (merge gốc + bổ sung)
        keyword_list = []
        
        # Keywords gốc
        if keywords_text and keywords_text != 'nan':
            kw_list = [k.strip() for k in keywords_text.replace('\n', ',').split(',') if k.strip()]
            keyword_list.extend(kw_list)
        
        # Keywords bổ sung
        if keywords_extra_text and keywords_extra_text != 'nan':
            kw_extra_list = [k.strip() for k in keywords_extra_text.replace('\n', ',').split(',') if k.strip()]
            keyword_list.extend(kw_extra_list)
        
        # Remove duplicates, keep order
        seen = set()
        final_keywords = []
        for kw in keyword_list:
            if kw not in seen:
                final_keywords.append(kw)
                seen.add(kw)
        
        # Add to structure
        result[major][minor][sub] = final_keywords
    
    # Save JSON
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    # Thống kê
    print(f"\n✅ Đã xuất: {output_json}")
    print(f"\n📊 THỐNG KÊ:")
    for major, minors in result.items():
        print(f"\n🔷 {major}")
        for minor, subs in minors.items():
            total_kw = sum(len(kw_list) for kw_list in subs.values())
            print(f"   └─ {minor}: {len(subs)} subs, {total_kw} keywords")
    
    return result


if __name__ == "__main__":
    excel_file = Path("D:/Works/CRM/tu_khoa_cho_cac_cot_cong_trinh.xlsx")
    output_json = Path("D:/Works/CRM/keywords_fixed.json")
    
    print("🔄 BẮT ĐẦU TRÍCH XUẤT...\n")
    keywords = extract_keywords_to_json(excel_file, output_json)
    print("\n✅ HOÀN TẤT!")
