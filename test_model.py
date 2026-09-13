"""model.py 的最小自检。

运行方式：python test_model.py（也可用 pytest 运行）。
覆盖 extract_class_weeks 支持的全部周次格式、周次合并与空闲时间计算。
"""

import pandas as pd

from model import calculate_free_schedule, extract_class_weeks, merge_consecutive_weeks


def test_extract_week_formats():
    # 空值 / 空白 → 无课
    assert extract_class_weeks(None) == []
    assert extract_class_weeks("") == []
    assert extract_class_weeks("   \n ") == []

    # x-y周
    assert extract_class_weeks("1-16周") == list(range(1, 17))
    assert extract_class_weeks("9-15周") == list(range(9, 16))

    # x-y周(单/双)，全角/半角括号
    assert extract_class_weeks("1-7周(单)") == [1, 3, 5, 7]
    assert extract_class_weeks("1-7周（单）") == [1, 3, 5, 7]
    assert extract_class_weeks("2-16周(双)") == [2, 4, 6, 8, 10, 12, 14, 16]

    # 单周次 x周(单/双)、x周
    assert extract_class_weeks("8周(双)") == [8]
    assert extract_class_weeks("16周") == [16]

    # (xx节) 前缀格式
    assert extract_class_weeks("(3-4节) 1-16周") == list(range(1, 17))
    assert extract_class_weeks("（3-4节） 1-16周") == list(range(1, 17))

    # 一个单元格里多段课程（换行 / 逗号分隔）
    assert extract_class_weeks("1-7周(单)\n9-15周") == [1, 3, 5, 7] + list(range(9, 16))
    assert extract_class_weeks("1-7周(单),8周(双)") == [1, 3, 5, 7, 8]

    # 与周次无关的文本不误报
    assert extract_class_weeks("无") == []
    assert extract_class_weeks("备注：教室待定") == []


def test_merge_consecutive_weeks():
    assert merge_consecutive_weeks([]) == ""
    assert merge_consecutive_weeks([7]) == "7"
    assert merge_consecutive_weeks([1, 2, 3, 5]) == "1-3,5"
    assert merge_consecutive_weeks(list(range(1, 17))) == "1-16"
    assert merge_consecutive_weeks([2, 4, 6, 8]) == "2,4,6,8"


def _schedule_df():
    return pd.DataFrame([
        ["", "", "", "", "", "", "", "", ""],
        ["", "节次", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"],
        ["上午", "1", "1-16周", "", "1-7周(单)", "", "", "", ""],
        ["", "2", "1-16周", "", "1-7周(单)", "", "", "", ""],
        ["上午", "3", "9-10周", "", "", "", "", "", ""],
        ["", "4", "9-10周", "", "", "", "", "", ""],
        ["其他课程", "5", "", "", "", "", "", "", ""],
    ])


def test_calculate_free_schedule():
    result = calculate_free_schedule(_schedule_df(), 16)

    # 1-2 节和 3-4 节各产出 6+7 条（星期一 1-2 节整周有课，不产出）
    assert len(result) == 13
    assert result[0] == {"weekday": "星期一", "section": "3-4", "free_desc": "1-8,11-16"}
    assert result[1] == {"weekday": "星期二", "section": "1-2", "free_desc": "1-16"}
    assert result[2] == {"weekday": "星期二", "section": "3-4", "free_desc": "1-16"}
    assert result[3] == {"weekday": "星期三", "section": "1-2", "free_desc": "2,4,6,8-16"}

    # “其他课程”之后的行不再参与统计
    assert {item["section"] for item in result} == {"1-2", "3-4"}


def test_calculate_free_schedule_bad_input():
    try:
        calculate_free_schedule(pd.DataFrame([["a"], ["b"], ["上午"]]), 16)
        assert False, "列数不足时应抛出 ValueError"
    except ValueError as exc:
        assert "列数不足" in str(exc)


if __name__ == "__main__":
    test_extract_week_formats()
    test_merge_consecutive_weeks()
    test_calculate_free_schedule()
    test_calculate_free_schedule_bad_input()
    print("model.py 自检全部通过")
