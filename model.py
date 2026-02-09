# model.py - 课程表解析核心
import pandas as pd
import re

def extract_class_weeks(course_text):
    """
    适配复杂课程文本：识别多段"x-y周(单/双)"格式，合并所有上课周数
    支持格式：
    - x-y周(单/双) → 如 1-7周(单)、2-16周(双)
    - x周(单/双) → 如 8周(双)
    - x-y周 → 如 9-15周（全周）
    - x周 → 如 16周
    - (xx节) x-y周 → 如 (3-4节) 1-16周
    """
    # 空值/空格直接返回空列表（无上课周）
    if pd.isna(course_text):
        return []
    
    # 转换为字符串并去除所有空白字符（包括空格、换行、制表符等）
    course_str = str(course_text).strip()
    if not course_str:  # 空字符串或仅空白字符 → 全周空闲
        return []
    
    class_weeks = set()  # 用集合去重，避免重复周数
    
    # 模式1：匹配 "x-y周(单/双)" 格式，如 "1-7周(单)"、"2-16周(双)"
    pattern1 = r'(\d+)-(\d+)周\s*[（(]\s*(单|双)\s*[）)]'
    matches1 = re.findall(pattern1, course_str)
    for match in matches1:
        try:
            start = int(match[0])
            end = int(match[1])
            week_type = match[2]
            if week_type == "单":
                # 单周：只保留奇数周（1, 3, 5, 7...）
                class_weeks.update(range(start if start%2==1 else start+1, end+1, 2))
            elif week_type == "双":
                # 双周：只保留偶数周（2, 4, 6, 8...）
                class_weeks.update(range(start if start%2==0 else start+1, end+1, 2))
        except ValueError:
            continue
    
    # 模式2：匹配 "x周(单/双)" 格式，如 "8周(双)"
    pattern2 = r'(\d+)周\s*[（(]\s*(单|双)\s*[）)]'
    matches2 = re.findall(pattern2, course_str)
    for match in matches2:
        try:
            week = int(match[0])
            week_type = match[1]
            if (week_type == "单" and week%2==1) or (week_type == "双" and week%2==0):
                class_weeks.add(week)
        except ValueError:
            continue
    
    # 模式3：匹配带括号的格式，如 "(3-4节) 1-16周" 或 "(3-4节) 1-16周(单)"
    pattern3 = r'\([^)]*\)\s*(\d+)-(\d+)\s*周\s*[（(]?\s*(单|双)?\s*[）)]?'
    matches3 = re.findall(pattern3, course_str)
    for match in matches3:
        try:
            start = int(match[0])
            end = int(match[1])
            week_type = match[2] if match[2] else ""
            if week_type == "单":
                class_weeks.update(range(start if start%2==1 else start+1, end+1, 2))
            elif week_type == "双":
                class_weeks.update(range(start if start%2==0 else start+1, end+1, 2))
            else:
                # 全周：所有周次
                class_weeks.update(range(start, end+1))
        except ValueError:
            continue
    
    # 模式4：匹配标准格式 "x-y周"（无单双周标识，且不在括号内）
    # 使用负向前瞻排除已经被模式1匹配的内容
    pattern4 = r'(?<![\d-])(\d+)-(\d+)周(?!\s*[（(]\s*(单|双))'
    matches4 = re.findall(pattern4, course_str)
    for match in matches4:
        try:
            start = int(match[0])
            end = int(match[1])
            # 全周：所有周次
            class_weeks.update(range(start, end+1))
        except ValueError:
            continue
    
    # 模式5：匹配单个周数 "x周"（无单双周标识，且不在范围中）
    # 使用负向前瞻排除已经被模式2匹配的内容
    pattern5 = r'(?<!\d)(\d+)周(?!\s*[（(]\s*(单|双))'
    matches5 = re.findall(pattern5, course_str)
    for match in matches5:
        try:
            week = int(match[0])
            class_weeks.add(week)
        except ValueError:
            continue
    
    # 去重并排序返回
    return sorted(list(class_weeks))

def merge_consecutive_weeks(weeks):
    """将连续周数合并为"x-y"格式（如[1,2,3,5]→"1-3,5"）"""
    if not weeks:
        return ""
    
    if len(weeks) == 1:
        return str(weeks[0])
    
    merged = []
    start = weeks[0]
    end = weeks[0]
    
    for i in range(1, len(weeks)):
        if weeks[i] == weeks[i-1] + 1:
            # 连续，继续扩展范围
            end = weeks[i]
        else:
            # 不连续，保存当前范围
            if start == end:
                merged.append(str(start))
            else:
                merged.append(f"{start}-{end}")
            start = weeks[i]
            end = weeks[i]
    
    # 保存最后一个范围
    if start == end:
        merged.append(str(start))
    else:
        merged.append(f"{start}-{end}")
    
    return ",".join(merged)

def calculate_free_schedule(df, total_weeks=16):
    """计算空闲时间：空格→全周空闲，连续周数合并"""
    # 课表结构映射（节次+星期）
    section_map = {1:"1-2", 2:"1-2", 3:"3-4", 4:"3-4", 5:"5-6", 6:"5-6",
                  7:"7-8", 8:"7-8", 9:"9-10", 10:"9-10", 11:"11-12", 12:"11-12"}
    
    free_schedule = []
    processed_sections = set()
    
    # 遍历课表行（跳过前2行表头，从第3行开始，索引2）
    for row_idx in range(2, df.shape[0]):
        section_cell = df.iloc[row_idx, 1]
        
        # 检查是否包含"其他课程"（忽略该行及之后的行）
        first_col = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ""
        if "其他课程" in first_col:
            break
        
        # 检查节次列是否为数字
        if pd.isna(section_cell) or not str(section_cell).strip().isdigit():
            continue
        
        section_num = int(section_cell)
        if section_num not in section_map:
            continue
        
        section = section_map[section_num]
        if section in processed_sections:
            continue
        
        # 🔴 核心修正：适配星期四=F+G列的映射逻辑（精准匹配Excel列索引）
        # 星期列映射：星期一=C列(2), 星期二=D列(3), 星期三=E列(4), 
        # 星期四=F列(5)+G列(6)合并, 星期五=H列(7), 星期六=I列(8), 星期日=J列(9)
        weekday_columns = [
            ("星期一", 2),    # C列（代码索引2）
            ("星期二", 3),    # D列（代码索引3）
            ("星期三", 4),    # E列（代码索引4）
            ("星期四", 5),    # F列（星期四第1部分）
            ("星期四", 6),    # G列（星期四第2部分，需合并）
            ("星期五", 7),    # H列（原G列被占用后，星期五顺延到H列）
            ("星期六", 8),    # I列
            ("星期日", 9)     # J列
        ]
        
        # 处理星期四：合并F列（索引5）和G列（索引6）的内容
        thursday_cols = [5, 6]  # F列和G列
        thursday_texts = []
        for col_idx in thursday_cols:
            if col_idx < df.shape[1]:
                course_text = df.iloc[row_idx, col_idx]
                if not (pd.isna(course_text) or str(course_text).strip() == ""):
                    thursday_texts.append(str(course_text).strip())
        
        # 合并星期四的F列和G列内容
        thursday_combined = " ".join(thursday_texts) if thursday_texts else ""
        
        for weekday, col_idx in weekday_columns:
            # 跳过超出DataFrame列范围的情况
            if col_idx >= df.shape[1]:
                continue
            
            # 🔴 关键：星期四合并F+G列内容（只解析F列，跳过G列重复）
            if weekday == "星期四":
                if col_idx == 6:  # G列是星期四的补充，不重复解析
                    continue
                course_text = thursday_combined
            else:
                # 其他星期正常读取对应列
                course_text = df.iloc[row_idx, col_idx]
            
            # 调试：打印关键单元格内容
            if row_idx == 2 and col_idx == 2:  # 第一行数据，星期一列（索引2）
                print(f"[DEBUG] 星期一第{section}节单元格内容：'{course_text}'")
            
            # 计算空闲周（空格=全周空闲）
            class_weeks = extract_class_weeks(course_text)
            
            # 调试：打印提取的上课周数
            if row_idx == 2 and col_idx == 2 and class_weeks:
                print(f"[DEBUG] 提取到的上课周数：{class_weeks}")
            
            all_weeks = set(range(1, total_weeks + 1))
            free_weeks = sorted(list(all_weeks - set(class_weeks)))
            
            # 合并连续周数
            if free_weeks:
                free_desc = merge_consecutive_weeks(free_weeks)
                free_schedule.append({
                    "weekday": weekday,
                    "section": section,
                    "free_desc": free_desc
                })
        
        processed_sections.add(section)
        if len(processed_sections) >= 6:  # 仅保留6个节次组合
            break
    
    # 按"星期→节次"排序
    weekday_order = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    section_order = ["1-2", "3-4", "5-6", "7-8", "9-10", "11-12"]
    free_schedule.sort(key=lambda x: (weekday_order.index(x["weekday"]), section_order.index(x["section"])))
    
    return free_schedule
