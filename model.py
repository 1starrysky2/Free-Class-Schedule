"""课程表解析核心逻辑。"""

import re

import pandas as pd

SECTION_MAP = {
    1: "1-2",
    2: "1-2",
    3: "3-4",
    4: "3-4",
    5: "5-6",
    6: "5-6",
    7: "7-8",
    8: "7-8",
    9: "9-10",
    10: "9-10",
    11: "11-12",
    12: "11-12",
}

WEEKDAY_COLUMNS = [
    ("星期一", 2),
    ("星期二", 3),
    ("星期三", 4),
    ("星期四", 5),
    ("星期五", 6),
    ("星期六", 7),
    ("星期日", 8),
]

WEEK_PATTERNS = {
    "range_with_type": r"(\d+)-(\d+)周\s*[（(]\s*(单|双)\s*[）)]",
    "single_with_type": r"(\d+)周\s*[（(]\s*(单|双)\s*[）)]",
    "bracketed_range": r"\([^)]*\)\s*(\d+)-(\d+)\s*周\s*[（(]?\s*(单|双)?\s*[）)]?",
    "normal_range": r"(?<![\d-])(\d+)-(\d+)周(?!\s*[（(]\s*(单|双))",
    "normal_single": r"(?<!\d)(\d+)周(?!\s*[（(]\s*(单|双))",
}

WEEKDAY_ORDER = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
SECTION_ORDER = ["1-2", "3-4", "5-6", "7-8", "9-10", "11-12"]


def _odd_weeks(start: int, end: int) -> range:
    return range(start if start % 2 == 1 else start + 1, end + 1, 2)


def _even_weeks(start: int, end: int) -> range:
    return range(start if start % 2 == 0 else start + 1, end + 1, 2)


def _match_week_pattern(pattern: str, text: str, handler) -> None:
    """通用正则匹配处理器。"""
    matches = re.findall(pattern, text)
    for match in matches:
        try:
            handler(match)
        except (ValueError, TypeError):
            continue

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

    def handle_range_with_type(match):
        start = int(match[0])
        end = int(match[1])
        week_type = match[2]
        if week_type == "单":
            class_weeks.update(_odd_weeks(start, end))
        elif week_type == "双":
            class_weeks.update(_even_weeks(start, end))

    def handle_single_with_type(match):
        week = int(match[0])
        week_type = match[1]
        if (week_type == "单" and week % 2 == 1) or (week_type == "双" and week % 2 == 0):
            class_weeks.add(week)

    def handle_bracketed_range(match):
        start = int(match[0])
        end = int(match[1])
        week_type = match[2] if match[2] else ""
        if week_type == "单":
            class_weeks.update(_odd_weeks(start, end))
        elif week_type == "双":
            class_weeks.update(_even_weeks(start, end))
        else:
            class_weeks.update(range(start, end + 1))

    def handle_normal_range(match):
        start = int(match[0])
        end = int(match[1])
        class_weeks.update(range(start, end + 1))

    def handle_normal_single(match):
        class_weeks.add(int(match[0]))

    _match_week_pattern(WEEK_PATTERNS["range_with_type"], course_str, handle_range_with_type)
    _match_week_pattern(WEEK_PATTERNS["single_with_type"], course_str, handle_single_with_type)
    _match_week_pattern(WEEK_PATTERNS["bracketed_range"], course_str, handle_bracketed_range)
    _match_week_pattern(WEEK_PATTERNS["normal_range"], course_str, handle_normal_range)
    _match_week_pattern(WEEK_PATTERNS["normal_single"], course_str, handle_normal_single)
    
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
    try:
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
            if section_num not in SECTION_MAP:
                continue

            section = SECTION_MAP[section_num]
            if section in processed_sections:
                continue

            for weekday, col_idx in WEEKDAY_COLUMNS:
                # 跳过超出DataFrame列范围的情况
                if col_idx >= df.shape[1]:
                    continue

                course_text = df.iloc[row_idx, col_idx]
                class_weeks = extract_class_weeks(course_text)
                all_weeks = set(range(1, total_weeks + 1))
                free_weeks = sorted(list(all_weeks - set(class_weeks)))

                if free_weeks:
                    free_desc = merge_consecutive_weeks(free_weeks)
                    free_schedule.append({"weekday": weekday, "section": section, "free_desc": free_desc})

            processed_sections.add(section)
            if len(processed_sections) >= 6:
                break

        free_schedule.sort(key=lambda x: (WEEKDAY_ORDER.index(x["weekday"]), SECTION_ORDER.index(x["section"])))
        return free_schedule
    except IndexError as exc:
        raise ValueError(f"课表列数不足，无法匹配星期列（需要至少9列）：{exc}") from exc
    except KeyError as exc:
        raise ValueError(f"节次映射错误，无效的节次编号：{exc}") from exc
    except Exception as exc:
        raise ValueError(
            "课表解析失败："
            f"{exc}。请检查课表格式是否为教务处标准导出格式，是否包含节次、星期、周数信息"
        ) from exc
