import streamlit as st  # Streamlit：用来把 Python 程序做成网页
import pandas as pd  # pandas：负责读取和处理 Excel 表格
import re
import mammoth
from datetime import datetime, timedelta, time, date
from pandas import Series
from typing import cast
from streamlit_local_storage import LocalStorage

FILE_PATH = "wz.xlsx"


# 判断是否“空值”
def has_value(value: object) -> bool:
    return bool(pd.notna(value)) and str(value).strip() != ""


# 把各种可能的时间格式，统一整理成 HH:MM 的文本
def format_time_hm(value: object) -> str:
    if not has_value(value):
        return ""

    if isinstance(value, time):
        return value.strftime("%H:%M")

    if isinstance(value, (str, int, float, date, datetime)):
        try:
            ts = pd.Timestamp(value)
            return ts.strftime("%H:%M")
        except (ValueError, TypeError):
            pass

    text = str(value).strip()

    if len(text) == 8 and text.count(":") == 2:
        return text[:5]

    if len(text) >= 16:
        return text[11:16]

    return text


# 读取Excel文件，预处理
def load_data(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)

    date_series = pd.to_datetime(df["日期"], errors="coerce")
    df["日期"] = pd.Series(date_series, index=df.index).dt.date

    df = df.sort_values("日期").reset_index(drop=True)
    return df


# 给某个日期打上“今天 / 过去几天 / 未来几天”的标签
def format_day_label(target_date, center_date):
    diff = (target_date - center_date).days
    if diff == 0:
        return "今天"
    elif diff < 0:
        return f"过去 {abs(diff)} 天"
    else:
        return f"未来 {diff} 天"


# 在指定的几列里，查找“距离某一天最近的上一条记录”和“最近的下一条记录”
def get_prev_next_rows(df: pd.DataFrame, center_date, cols: list[str]):
    mask = df[cols].map(has_value).any(axis=1)
    candidates = df[mask].copy()

    prev_rows = candidates[candidates["日期"] < center_date]
    next_rows = candidates[candidates["日期"] > center_date]

    prev_row = prev_rows.iloc[-1] if not prev_rows.empty else None
    next_row = next_rows.iloc[0] if not next_rows.empty else None

    return prev_row, next_row


# 把“前一条/后一条事件记录”格式化成一行
def build_prev_next_line(row, center_date, prefix_type: str, value_cols: list[str], time_col: str | None = None):
    if row is None:
        return ""

    diff_days = abs((row["日期"] - center_date).days)
    prefix = f"      {diff_days}日{'前' if prefix_type == 'prev' else '后'}："

    parts = [prefix + format_md_week(row)]

    for col in value_cols:
        if has_value(row.get(col, "")):
            parts.append(str(row[col]))

    if time_col:
        time_text = format_time_hm(row.get(time_col, ""))
        if time_text:
            parts.append(time_text)

    return " ".join(parts)


# 把某一天的一整行数据，整理成“当天详情页”展示用的多行文本
def row_to_lines(row, center_date, df: pd.DataFrame, anchor_date_a=None):
    label = format_day_label(row["日期"], center_date)
    lines = []

    # 1. 第一行：[今天] 2026-04-15（星期三） 17048
    week = row.get("星期", "")
    first_line = f"[{label}] {row['日期']}（星期{week}）"
    user_day_no = calc_user_day_number(row["日期"], anchor_date_a)
    if user_day_no != "":
        first_line += f" {user_day_no}"
    lines.append("=" * 48)
    lines.append(first_line)

    # 2. 农历：二月小廿八
    lunar_month = str(row.get("农历月", "")).strip()
    big_small = str(row.get("大小月", "")).strip()
    lunar_day = str(row.get("农历日", "")).strip()
    lunar_text = f"{lunar_month}{big_small}{lunar_day}".strip()
    if lunar_text:
        lines.append(f"农历：{lunar_text}")

    # 3. 黄历：黑道朱雀 平日 壁宿
    almanac_parts = []
    if has_value(row.get("黄道黑道", "")):
        almanac_parts.append(str(row["黄道黑道"]))
    if has_value(row.get("十二建日", "")):
        almanac_parts.append(f"{row['十二建日']}日")
    if has_value(row.get("星宿", "")):
        almanac_parts.append(f"{row['星宿']}宿")
    if almanac_parts:
        lines.append(f"黄历：{' '.join(almanac_parts)}")
    else:
        lines.append("黄历：（无）")

    # 4. 干支：丙午年 壬辰月 己未日
    ganzhi_parts = []
    if has_value(row.get("年柱", "")):
        ganzhi_parts.append(f"{row['年柱']}年")
    if has_value(row.get("月柱", "")):
        ganzhi_parts.append(f"{row['月柱']}月")
    if has_value(row.get("日柱", "")):
        ganzhi_parts.append(f"{row['日柱']}日")
    if ganzhi_parts:
        lines.append(f"干支：{' '.join(ganzhi_parts)}")

    # 5. 时辰：时柱1 ~ 时柱12
    shichen_values_a = []
    for i in range(1, 7):
        col = f"时柱{i}"
        if has_value(row.get(col, "")):
            shichen_values_a.append(str(row[col]))
    if shichen_values_a:
        lines.append(f"时辰：{' '.join(shichen_values_a)}")

    shichen_values_b = []
    for j in range(7, 13):
        col = f"时柱{j}"
        if has_value(row.get(col, "")):
            shichen_values_b.append(str(row[col]))
    if shichen_values_b:
        lines.append(f"               {' '.join(shichen_values_b)}")

    # 6. 节气：清明 3 虹始见 07:00
    jieqi_parts = []
    if has_value(row.get("节气", "")):
        jieqi_parts.append(str(row["节气"]))
    if has_value(row.get("物候", "")):
        jieqi_parts.append(str(row["物候"]))
    time_text = format_time_hm(row.get("时间点", ""))
    if time_text:
        jieqi_parts.append(time_text)

    if jieqi_parts:
        lines.append(f"节气：{' '.join(jieqi_parts)}")
    else:
        lines.append("节气：")

    prev_jieqi, next_jieqi = get_prev_next_rows(df, row["日期"], ["节气", "物候", "时间点"])
    prev_jieqi_line = build_prev_next_line(prev_jieqi, row["日期"], "prev", ["节气", "物候"], "时间点")
    next_jieqi_line = build_prev_next_line(next_jieqi, row["日期"], "next", ["节气", "物候"], "时间点")
    if prev_jieqi_line:
        lines.append(prev_jieqi_line)
    if next_jieqi_line:
        lines.append(next_jieqi_line)

    # 7. 月相：有则显示当天，没有则只显示前后月相
    moon_parts = []
    if has_value(row.get("月相", "")):
        moon_parts.append(str(row["月相"]))
    moon_time = format_time_hm(row.get("月相时间", ""))
    if moon_time:
        moon_parts.append(moon_time)

    if moon_parts:
        lines.append(f"月相：{' '.join(moon_parts)}")
    else:
        lines.append("月相：")

    prev_moon, next_moon = get_prev_next_rows(df, row["日期"], ["月相", "月相时间"])
    prev_moon_line = build_prev_next_line(prev_moon, row["日期"], "prev", ["月相"], "月相时间")
    next_moon_line = build_prev_next_line(next_moon, row["日期"], "next", ["月相"], "月相时间")
    if prev_moon_line:
        lines.append(prev_moon_line)
    if next_moon_line:
        lines.append(next_moon_line)

    # 8. 天象：逐行归集
    astro_map = [
        ("地日", "日相时间"),
        ("紫孛", "紫孛相时间"),
        ("月交", "月交相时间"),
        ("水星", "水相时间"),
        ("金星", "金相时间"),
        ("火星", "火相时间"),
        ("木星", "木相时间"),
        ("土星", "土相时间"),
        ("天王星", "天相时间"),
        ("海王星", "海相时间"),
        ("冥王星", "冥相时间"),
    ]

    astro_lines = []
    for label_name, time_col in astro_map:
        value = row.get(label_name, "")
        time_text = format_time_hm(row.get(time_col, ""))

        if has_value(value) or time_text:
            text = f"{label_name}{str(value).strip() if has_value(value) else ''}"
            if time_text:
                text += f" {time_text}"
            astro_lines.append(f"  {text}")

    if astro_lines:
        lines.append("天象：")
        lines.extend(astro_lines)
    else:
        lines.append("天象：（无）")

    # 9. 行星逆行
    retrograde_map = [
        ("水星逆行", "水星"),
        ("金星逆行", "金星"),
        ("火星逆行", "火星"),
        ("木星逆行", "木星"),
        ("土星逆行", "土星"),
        ("天王逆行", "天王星"),
        ("海王逆行", "海王星"),
        ("冥王逆行", "冥王星"),
    ]

    retrograde_list = []
    for field, display_name in retrograde_map:
        if has_value(row.get(field, "")):
            retrograde_list.append(display_name)

    if retrograde_list:
        lines.append(f"行星逆行：{' '.join(retrograde_list)}")
    else:
        lines.append("行星逆行：（无）")

    # 10. 卦象
    gua_parts_a = []
    if has_value(row.get("十年卦", "")):
        gua_parts_a.append(f"【十年】{row['十年卦']}")
    if has_value(row.get("年卦", "")):
        gua_parts_a.append(f"【年】{row['年卦']}")
    if gua_parts_a:
        lines.append(f"卦象：{'      '.join(gua_parts_a)}")

    gua_parts_b = []
    if has_value(row.get("月卦", "")):
        gua_parts_b.append(f"【月】{row['月卦']}")
    if has_value(row.get("旬卦", "")):
        gua_parts_b.append(f"【旬】{row['旬卦']}")
    if has_value(row.get("日卦", "")):
        gua_parts_b.append(f"【日】{row['日卦']}")
    if gua_parts_b:
        lines.append(f"        {' '.join(gua_parts_b)}")

    return "\n".join(lines)


# 七天播报页会用到的函数
def relative_box_label(target_date, center_date):
    diff = (target_date - center_date).days
    if diff == 0:
        return "当天"
    if diff < 0:
        return f"前{abs(diff)}天"
    return f"后{diff}天"


# 用于事件集合部分的相对标签
def relative_event_label(target_date, center_date):
    diff = (target_date - center_date).days
    if diff == 0:
        return "当天"
    if diff < 0:
        return f"{abs(diff)}日前"
    return f"{diff}日后"


# 把某一行的日期和星期格式化
def format_md_week(row) -> str:
    d = row["日期"]
    week = row.get("星期", "")
    return f"{d.month:02d}-{d.day:02d}（{week}）"


# 拼接农历文本
def build_lunar_text(row) -> str:
    parts = [
        str(row.get("农历月", "")).strip(),
        str(row.get("大小月", "")).strip(),
        str(row.get("农历日", "")).strip(),
    ]
    return "".join(parts).strip()


# 拼接干支文本
def build_ganzhi_text(row) -> str:
    parts = []
    if has_value(row.get("年柱", "")):
        parts.append(f"{row['年柱']}年")
    if has_value(row.get("月柱", "")):
        parts.append(f"{row['月柱']}月")
    if has_value(row.get("日柱", "")):
        parts.append(f"{row['日柱']}日")
    if has_value(row.get("时柱1", "")):
        parts.append(f"{row['时柱1']}时")
    return " ".join(parts)


# 拼接黄历文本
def build_huangli_text(row) -> str:
    parts = []
    if has_value(row.get("黄道黑道", "")):
        parts.append(str(row["黄道黑道"]))
    if has_value(row.get("十二建日", "")):
        parts.append(f"{row['十二建日']}日")
    if has_value(row.get("星宿", "")):
        parts.append(f"{row['星宿']}宿")
    return " ".join(parts)


# 拼接卦象文本
def build_gua_text(row) -> str:
    parts = []
    if has_value(row.get("年卦", "")):
        parts.append(f"[年]{row['年卦']}")
    if has_value(row.get("月卦", "")):
        parts.append(f"[月]{row['月卦']}")
    if has_value(row.get("旬卦", "")):
        parts.append(f"[旬]{row['旬卦']}")
    if has_value(row.get("日卦", "")):
        parts.append(f"[日]{row['日卦']}")
    return " ".join(parts)


# 七天主视图中的第一行
def build_window_day_line1(row, center_date, anchor_date_a=None) -> str:
    label = relative_box_label(row["日期"], center_date)
    parts = [f"【{label}】{row['日期']}（星期{row.get('星期', '')}）"]

    lunar_text = build_lunar_text(row)
    if lunar_text:
        parts.append(lunar_text)

    user_day_no = calc_user_day_number(row["日期"], anchor_date_a)
    if user_day_no != "":
        parts.append(str(user_day_no))

    return " ".join(parts)


# 七天主视图中的第二行
def build_window_day_line2(row) -> str:
    huangli_text = build_huangli_text(row)
    ganzhi_text = build_ganzhi_text(row)

    left = f"{huangli_text}" if huangli_text else "黄历：（无）"
    right = f"{ganzhi_text}" if ganzhi_text else "干支：（无）"

    return f"{left}   {right}"


# 七天主视图中的第三行
def build_window_day_line3(row) -> str:
    gua_text = build_gua_text(row)

    return gua_text


# 生成一个布尔筛选条件
def build_event_mask(df: pd.DataFrame, value_cols: list[str], time_col: str | None = None):
    cols = [col for col in value_cols if col in df.columns]
    if time_col and time_col in df.columns:
        cols.append(time_col)

    if not cols:
        return pd.Series([False] * len(df), index=df.index)

    return df[cols].map(has_value).any(axis=1)


# 把一条事件记录格式化
def build_event_line(row: Series, center_date, value_cols: list[str], time_col: str | None = None) -> str:
    prefix = relative_event_label(row["日期"], center_date)
    parts = [f"  {prefix}：{format_md_week(row)}"]

    for col in value_cols:
        if has_value(row.get(col, "")):
            parts.append(str(row[col]).strip())

    if time_col:
        time_text = format_time_hm(row.get(time_col, ""))
        if time_text:
            parts.append(time_text)

    return " ".join(parts)


# 通用事件集合构造器
def build_event_section(
        df: pd.DataFrame,
        center_date,
        range_start,
        range_end,
        value_cols: list[str],
        time_col: str | None = None,
):
    mask = build_event_mask(df, value_cols, time_col)

    in_range = df[mask & (df["日期"] >= range_start) & (df["日期"] <= range_end)].copy()
    before = df[mask & (df["日期"] < range_start)].copy()
    after = df[mask & (df["日期"] > range_end)].copy()

    prev_row = cast(Series | None, before.iloc[-1] if not before.empty else None)
    next_row = cast(Series | None, after.iloc[0] if not after.empty else None)

    lines = []

    if prev_row is not None:
        lines.append(build_event_line(prev_row, center_date, value_cols, time_col))

    if not in_range.empty:
        for _, row in in_range.iterrows():
            lines.append(build_event_line(row, center_date, value_cols, time_col))

    if next_row is not None:
        lines.append(build_event_line(next_row, center_date, value_cols, time_col))

    if not lines:
        return [f"（无）"]

    return lines


# 构造“天象”集合
def build_astro_section(df: pd.DataFrame, center_date):
    astro_map = [
        ("地日", "日相时间"),
        ("紫孛", "紫孛相时间"),
        ("月交", "月交相时间"),
        ("水星", "水相时间"),
        ("金星", "金相时间"),
        ("火星", "火相时间"),
        ("木星", "木相时间"),
        ("土星", "土相时间"),
        ("天王星", "天相时间"),
        ("海王星", "海相时间"),
        ("冥王星", "冥相时间"),
    ]

    start_date = center_date - timedelta(days=7)
    end_date = center_date + timedelta(days=7)
    window_df = df[(df["日期"] >= start_date) & (df["日期"] <= end_date)].copy()

    lines = []
    found = False

    for _, row in window_df.iterrows():
        for label_name, time_col in astro_map:
            value = row.get(label_name, "")
            time_text = format_time_hm(row.get(time_col, ""))

            if has_value(value) or time_text:
                prefix = relative_event_label(row["日期"], center_date)
                detail = f"{label_name}{str(value).strip() if has_value(value) else ''}".strip()

                parts = [f"  {prefix}：{format_md_week(row)}"]
                if detail:
                    parts.append(detail)
                if time_text:
                    parts.append(time_text)

                lines.append(" ".join(parts))
                found = True

    if not found:
        return ["（无）"]

    return lines


# 根据某个“逆行字段”，把连续的逆行日期合并成区间
def build_retrograde_intervals(df: pd.DataFrame, field: str):
    if field not in df.columns:
        return []

    mask = df[field].map(has_value)
    retro_dates = list(df.loc[mask, "日期"])

    if not retro_dates:
        return []

    intervals = []
    start = retro_dates[0]
    end = retro_dates[0]

    for d in retro_dates[1:]:
        if d - end == timedelta(days=1):
            end = d
        else:
            intervals.append((start, end))
            start = d
            end = d

    intervals.append((start, end))
    return intervals


# 构造“行星逆行”集合说明
def build_retrograde_section(df: pd.DataFrame, center_date):
    retrograde_map = [
        ("水星逆行", "水星"),
        ("金星逆行", "金星"),
        ("火星逆行", "火星"),
        ("木星逆行", "木星"),
        ("土星逆行", "土星"),
        ("天王逆行", "天王星"),
        ("海王逆行", "海王星"),
        ("冥王逆行", "冥王星"),
    ]

    lines = []

    for field, display_name in retrograde_map:
        intervals = build_retrograde_intervals(df, field)

        current_interval = None
        for start, end in intervals:
            if start <= center_date <= end:
                current_interval = (start, end)
                break

        if current_interval is not None:
            start, end = current_interval
            display_end = end + timedelta(days=1)
            lines.append(
                f"{display_name}（逆行中），本次逆行开始于{start}，结束于{display_end}"
            )
        else:
            next_interval = None
            for start, end in intervals:
                if start > center_date:
                    next_interval = (start, end)
                    break

            if next_interval is not None:
                start, _ = next_interval
                lines.append(
                    f"{display_name}（无逆行），下次逆行开始于{start}"
                )
            else:
                lines.append(
                    f"{display_name}（无逆行），后续数据中未找到下一次逆行"
                )

    return lines


# 把“单项查询”结果格式化成一行文本
def format_keyword_event_line(row, keyword: str) -> str:
    date_text = f"{row['日期']}（星期{row.get('星期', '')}）"

    if keyword == "节气":
        parts = []
        if has_value(row.get("节气", "")):
            parts.append(str(row["节气"]))
        if has_value(row.get("物候", "")):
            parts.append(str(row["物候"]))
        time_text = format_time_hm(row.get("时间点", ""))
        if time_text:
            parts.append(time_text)
        return f"{date_text} " + " ".join(parts)

    if keyword == "月相":
        parts = []
        if has_value(row.get("月相", "")):
            parts.append(str(row["月相"]))
        time_text = format_time_hm(row.get("月相时间", ""))
        if time_text:
            parts.append(time_text)
        return f"{date_text} " + " ".join(parts)

    keyword_map = {
        "地日": ("地日", "日相时间"),
        "紫孛": ("紫孛", "紫孛相时间"),
        "水星": ("水星", "水相时间"),
        "金星": ("金星", "金相时间"),
        "火星": ("火星", "火相时间"),
        "木星": ("木星", "木相时间"),
        "土星": ("土星", "土相时间"),
        "天王星": ("天王星", "天相时间"),
        "海王星": ("海王星", "海相时间"),
        "冥王星": ("冥王星", "冥相时间"),
    }

    value_col, time_col = keyword_map[keyword]

    parts = []
    if has_value(row.get(value_col, "")):
        parts.append(str(row[value_col]))
    time_text = format_time_hm(row.get(time_col, ""))
    if time_text:
        parts.append(time_text)

    return f"{date_text} " + " ".join(parts)

# 用 Streamlit 缓存读表结果，减少重复读取 Excel 的开销
@st.cache_data
def get_data():
    return load_data(FILE_PATH)


@st.cache_data
def get_hexagrams():
    with open("yiv08.docx", "rb") as f:
        result = mammoth.convert_to_html(f)
    html = result.value

    pattern = r'id="_([^"]+[䷀-䷿][^"]*)"'
    anchors = list(re.finditer(pattern, html))

    gua_anchors = []
    seen = set()
    for m in anchors:
        name = m.group(1)
        base = re.sub(r'_\d+$', '', name)
        if base not in seen and re.search(r'[䷀-䷿]', base):
            seen.add(base)
            gua_anchors.append((base, m.start()))

    hexagram_html = {}
    for i, (name, pos) in enumerate(gua_anchors):
        end = gua_anchors[i + 1][1] if i + 1 < len(gua_anchors) else len(html)
        tag_start = html.rfind('<', 0, pos)
        chunk = html[tag_start:end]
        gua_name = re.sub(r'[䷀-䷿].*', '', name).strip()
        hexagram_html[gua_name] = chunk

    return hexagram_html


# 决定网页默认打开时显示哪一天
def get_default_date(df):
    min_date = df["日期"].min()
    max_date = df["日期"].max()
    today = datetime.today().date()
    if min_date <= today <= max_date:
        return today
    return max_date

# 从浏览器cookie中读取用户之前保存
def get_anchor_date_from_cookie(fallback: date) -> date | None:
    value = local_storage.getItem("anchor_date")

    if not value:
        return None

    try:
        ts = pd.Timestamp(value)
    except (ValueError, TypeError):
        return fallback

    if pd.isna(ts):
        return fallback

    return cast(date, ts.to_pydatetime().date())


# 把当前用户设定日期写入浏览器cookie
def save_anchor_date_to_cookie(anchor_date_a: date | None):
    if anchor_date_a is None:
        local_storage.deleteItem("anchor_date")
    else:
        local_storage.setItem("anchor_date", anchor_date_a.isoformat())


# 把各种可能的值，统一整理成真正的 date
def normalize_date_value(value, fallback: date) -> date:
    if value is None:
        return fallback

    try:
        if pd.isna(value):
            return fallback
    except TypeError:
        pass

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, str):
        try:
            ts = pd.Timestamp(value)
        except (ValueError, TypeError):
            return fallback

        if pd.isna(ts):
            return fallback

        return cast(date, ts.to_pydatetime().date())

    return fallback


# 天数转化
def resolve_date_input(value, fallback):
    if isinstance(value, tuple):
        return value[0] if value else fallback
    return value if value is not None else fallback


# 天数转化
def calc_user_day_number(row_date, anchor_date_a):
    if anchor_date_a is None:
        return ""

    delta = (row_date - anchor_date_a).days
    return delta + 1 if delta >= 0 else delta


# 实例化（约在所有函数定义之后，页面逻辑之前）
local_storage = LocalStorage()


# Streamlit 页面设置
st.set_page_config(page_title="我的日历本", layout="wide")
st.title("我的日历本")

dfr = get_data()
hexagram_data = get_hexagrams()
default_date = get_default_date(dfr)

# 用户自定义起始日（会话 + Cookie）
cookie_anchor_date = get_anchor_date_from_cookie(default_date)

if "use_anchor_date" not in st.session_state:
    st.session_state.use_anchor_date = cookie_anchor_date is not None

if "anchor_date" not in st.session_state:
    st.session_state.anchor_date = cookie_anchor_date if cookie_anchor_date is not None else default_date

if "last_use_anchor_date" not in st.session_state:
    st.session_state.last_use_anchor_date = st.session_state.use_anchor_date

with st.sidebar:
    st.subheader("我的天数设置")

    use_anchor_date = st.toggle(
        "启用自定义关键日",
        key="use_anchor_date"
    )

    if use_anchor_date:
        anchor_date_default = normalize_date_value(
            st.session_state.get("anchor_date"),
            default_date
        )

        anchor_date_input = st.date_input(
            "选择关键日",
            value=anchor_date_default,
            min_value=date(1900, 1, 1),
            key="anchor_date"
        )

        anchor_date = resolve_date_input(anchor_date_input, default_date)

        # 写入 cookie
        save_anchor_date_to_cookie(anchor_date)
    else:
        anchor_date = None
        # 只有从“开”切到“关”时才删 cookie
        if st.session_state.last_use_anchor_date:
            save_anchor_date_to_cookie(None)

    st.session_state.last_use_anchor_date = use_anchor_date


tab1, tab2, tab3 = st.tabs(["一天详情", "七天播报（±3）", "单项查询"])

# 页面一：当天详情
with tab1:
    # 初始化 session_state 里的日期
    if "tab1_date" not in st.session_state:
        st.session_state.tab1_date = default_date

    query_date_input = st.date_input(
        "选择日期",
        value=st.session_state.tab1_date,
        key="tab1_date_input"
    )
    query_date = resolve_date_input(query_date_input, default_date)
    st.session_state.tab1_date = query_date

    # 三个快捷按钮（强制横排）
    st.markdown("""
            <style>
            div[data-testid="stHorizontalBlock"] > div {
                min-width: 0 !important;
            }
            </style>
        """, unsafe_allow_html=True)

    col_prev, col_today, col_next = st.columns(3)
    with col_prev:
        if st.button("←上一天", use_container_width=True):
            st.session_state.tab1_date = query_date - timedelta(days=1)
            st.rerun()
    with col_today:
        if st.button("⊙今天", use_container_width=True):
            st.session_state.tab1_date = datetime.today().date()
            st.rerun()
    with col_next:
        if st.button("下一天→", use_container_width=True):
            st.session_state.tab1_date = query_date + timedelta(days=1)
            st.rerun()

    row_df = dfr[dfr["日期"] == query_date]

    if row_df.empty:
        st.warning(f"{query_date} 没有记录。")
    else:
        row_data = row_df.iloc[0]
        text_data = row_to_lines(row_data, query_date, dfr, anchor_date)
        st.text(text_data)

        # 卦象解析（移到 else 里面）
        gua_cols = ["十年卦", "年卦", "月卦", "旬卦", "日卦"]
        gua_labels = {"十年卦": "十年", "年卦": "年", "月卦": "月", "旬卦": "旬", "日卦": "日"}

        for column in gua_cols:
            gua_val = str(row_data.get(column, "")).strip()
            if not has_value(gua_val):
                continue
            gua_name = re.sub(r'^\d+\.', '', gua_val).strip()
            if gua_name in hexagram_data:
                with st.expander(f"【{gua_labels[column]}卦】{gua_name} — 点击展开解析"):
                    st.html(f"""
                        <style>
                        img {{ max-width: 100%; height: auto; }}
                        </style>
                        {hexagram_data[gua_name]}
                    """)
# 页面二：七天播报（±3）
with tab2:
    center_date_input = st.date_input("选择中心日期", value=default_date, key="center_date")
    cen_date = resolve_date_input(center_date_input, default_date)

    start_date_a = cen_date - timedelta(days=3)
    end_date_a = cen_date + timedelta(days=3)

    window_dfr = dfr[(dfr["日期"] >= start_date_a) & (dfr["日期"] <= end_date_a)].copy()

    st.subheader("七天主视图")
    for _, row_a in window_dfr.iterrows():
        day_text = "\n".join([
            "-" * 60,
            build_window_day_line1(row_a, cen_date, anchor_date),
            build_window_day_line2(row_a),
            build_window_day_line3(row_a)
        ])
        st.text(day_text)

    st.divider()

    st.subheader("节气")
    for line in build_event_section(dfr, cen_date, start_date_a, end_date_a, ["节气", "物候"], "时间点"):
        st.text(line)

    st.divider()

    st.subheader("月相")
    for line in build_event_section(dfr, cen_date, start_date_a, end_date_a, ["月相"], "月相时间"):
        st.text(line)

    st.divider()

    st.subheader("天象")
    for line in build_astro_section(dfr, cen_date):
        st.text(line)

    st.divider()

    st.subheader("行星逆行")
    for line in build_retrograde_section(dfr, cen_date):
        st.text(line)

# 页面三：单项查询
with tab3:
    supported_keywords = [
        "节气", "月相", "地日", "紫孛",
        "水星", "金星", "火星", "木星", "土星",
        "天王星", "海王星", "冥王星"
    ]

    keyword_data = st.selectbox("选择关键词", supported_keywords)
    start_date_kw_input = st.date_input("起始日期", value=default_date, key="kw_date")
    start_date_kw = resolve_date_input(start_date_kw_input, default_date)

    if keyword_data == "节气":
        mask_a = (
                (dfr["日期"] >= start_date_kw) &
                (dfr["节气"].map(has_value) | dfr["物候"].map(has_value) | dfr["时间点"].map(has_value))
        )
    elif keyword_data == "月相":
        mask_a = (
                (dfr["日期"] >= start_date_kw) &
                (dfr["月相"].map(has_value) | dfr["月相时间"].map(has_value))
        )
    else:
        keyword_map_data = {
            "地日": ("地日", "日相时间"),
            "紫孛": ("紫孛", "紫孛相时间"),
            "水星": ("水星", "水相时间"),
            "金星": ("金星", "金相时间"),
            "火星": ("火星", "火相时间"),
            "木星": ("木星", "木相时间"),
            "土星": ("土星", "土相时间"),
            "天王星": ("天王星", "天相时间"),
            "海王星": ("海王星", "海相时间"),
            "冥王星": ("冥王星", "冥相时间"),
        }
        value_column, time_column = keyword_map_data[keyword_data]
        mask_a = (
                (dfr["日期"] >= start_date_kw) &
                (dfr[value_column].map(has_value) | dfr[time_column].map(has_value))
        )

    result = dfr[mask_a].copy().head(12)

    if result.empty:
        st.info(f"从 {start_date_kw} 开始，未找到“{keyword_data}”的后续有效记录。")
    else:
        for _, row_b in result.iterrows():
            st.text(format_keyword_event_line(row_b, keyword_data))
