import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time, date
from pandas import Series
from typing import cast

FILE_PATH = "wz.xlsx"


def has_value(value: object) -> bool:
    return bool(pd.notna(value)) and str(value).strip() != ""


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


def load_data(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)

    date_series = pd.to_datetime(df["日期"], errors="coerce")
    df["日期"] = pd.Series(date_series, index=df.index).dt.date

    df = df.sort_values("日期").reset_index(drop=True)
    return df


def format_day_label(target_date, center_date):
    diff = (target_date - center_date).days
    if diff == 0:
        return "今天"
    elif diff < 0:
        return f"过去 {abs(diff)} 天"
    else:
        return f"未来 {diff} 天"


def get_prev_next_rows(df: pd.DataFrame, center_date, cols: list[str]):
    mask = df[cols].map(has_value).any(axis=1)
    candidates = df[mask].copy()

    prev_rows = candidates[candidates["日期"] < center_date]
    next_rows = candidates[candidates["日期"] > center_date]

    prev_row = prev_rows.iloc[-1] if not prev_rows.empty else None
    next_row = next_rows.iloc[0] if not next_rows.empty else None

    return prev_row, next_row


def build_event_detail(row, value_col: str, time_col: str, include_label=False, label_name=""):
    if row is None:
        return ""

    value = row.get(value_col, "")
    time_text = format_time_hm(row.get(time_col, ""))

    parts = []

    if include_label and label_name:
        if has_value(value):
            parts.append(f"{label_name}{value}")
        else:
            parts.append(label_name)
    else:
        if has_value(value):
            parts.append(str(value))

    if time_text:
        parts.append(time_text)

    return " ".join(parts).strip()


def build_prev_next_line(row, center_date, prefix_type: str, value_cols: list[str], time_col: str | None = None):
    if row is None:
        return ""

    diff_days = abs((row["日期"] - center_date).days)
    prefix = f"  {diff_days}日{'前' if prefix_type == 'prev' else '后'}："

    parts = [prefix + format_md_week(row)]

    for col in value_cols:
        if has_value(row.get(col, "")):
            parts.append(str(row[col]))

    if time_col:
        time_text = format_time_hm(row.get(time_col, ""))
        if time_text:
            parts.append(time_text)

    return " ".join(parts)


def row_to_lines(row, center_date, df: pd.DataFrame):
    label = format_day_label(row["日期"], center_date)
    lines = []

    # 1. 第一行：[今天] 2026-04-15（星期三） 17048
    week = row.get("星期", "")
    my_days = row.get("我的天数", "")
    first_line = f"[{label}] {row['日期']}（星期{week}）"
    if has_value(my_days):
        first_line += f" {my_days}"
    lines.append("=" * 60)
    lines.append(first_line)

    # 2. 农历：二月小廿八
    lunar_month = str(row.get("农历月", "")).strip()
    big_small = str(row.get("大小月", "")).strip()
    lunar_day = str(row.get("农历日", "")).strip()
    lunar_text = f"{lunar_month}{big_small}{lunar_day}".strip()
    if lunar_text:
        lines.append(f"农历：{lunar_text}")

    # 3. 干支：丙午年 壬辰月 己未日
    ganzhi_parts = []
    if has_value(row.get("年柱", "")):
        ganzhi_parts.append(f"{row['年柱']}年")
    if has_value(row.get("月柱", "")):
        ganzhi_parts.append(f"{row['月柱']}月")
    if has_value(row.get("日柱", "")):
        ganzhi_parts.append(f"{row['日柱']}日")
    if ganzhi_parts:
        lines.append(f"干支：{' '.join(ganzhi_parts)}")

    # 4. 时辰：时柱1 ~ 时柱12
    shichen_values = []
    for i in range(1, 13):
        col = f"时柱{i}"
        if has_value(row.get(col, "")):
            shichen_values.append(str(row[col]))
    if shichen_values:
        lines.append(f"时辰：{' '.join(shichen_values)}")

    # 5. 节气：清明 3 虹始见 07:00
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

    # 6. 月相：有则显示当天，没有则只显示前后月相
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

    # 7. 黄历：黑道朱雀 平日 壁宿
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
    gua_parts = []
    if has_value(row.get("十年卦", "")):
        gua_parts.append(f"【十年】{row['十年卦']}")
    if has_value(row.get("年卦", "")):
        gua_parts.append(f"【年】{row['年卦']}")
    if has_value(row.get("月卦", "")):
        gua_parts.append(f"【月】{row['月卦']}")
    if has_value(row.get("旬卦", "")):
        gua_parts.append(f"【日】{row['旬卦']}")
    if has_value(row.get("日卦", "")):
        gua_parts.append(f"【时】{row['日卦']}")
    if gua_parts:
        lines.append(f"卦象：{' '.join(gua_parts)}")

    return "\n".join(lines)


def relative_box_label(target_date, center_date):
    diff = (target_date - center_date).days
    if diff == 0:
        return "当天"
    if diff < 0:
        return f"前{abs(diff)}天"
    return f"后{diff}天"


def relative_event_label(target_date, center_date):
    diff = (target_date - center_date).days
    if diff == 0:
        return "当天"
    if diff < 0:
        return f"{abs(diff)}日前"
    return f"{diff}日后"


def format_md_week(row) -> str:
    d = row["日期"]
    week = row.get("星期", "")
    return f"{d.month:02d}-{d.day:02d}（{week}）"


def build_lunar_text(row) -> str:
    parts = [
        str(row.get("农历月", "")).strip(),
        str(row.get("大小月", "")).strip(),
        str(row.get("农历日", "")).strip(),
    ]
    return "".join(parts).strip()


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


def build_huangli_text(row) -> str:
    parts = []
    if has_value(row.get("黄道黑道", "")):
        parts.append(str(row["黄道黑道"]))
    if has_value(row.get("十二建日", "")):
        parts.append(f"{row['十二建日']}日")
    if has_value(row.get("星宿", "")):
        parts.append(f"{row['星宿']}宿")
    return " ".join(parts)


def build_gua_text(row) -> str:
    parts = []
    if has_value(row.get("十年卦", "")):
        parts.append(f"【十年】{row['十年卦']}")
    if has_value(row.get("年卦", "")):
        parts.append(f"【年】{row['年卦']}")
    if has_value(row.get("月卦", "")):
        parts.append(f"【月】{row['月卦']}")
    if has_value(row.get("旬卦", "")):
        parts.append(f"【日】{row['旬卦']}")
    if has_value(row.get("日卦", "")):
        parts.append(f"【时】{row['日卦']}")
    return " ".join(parts)


def build_window_day_line1(row, center_date) -> str:
    label = relative_box_label(row["日期"], center_date)
    parts = [f"【{label}】{row['日期']}（星期{row.get('星期', '')}）"]

    if has_value(row.get("我的天数", "")):
        parts.append(str(row["我的天数"]))

    lunar_text = build_lunar_text(row)
    if lunar_text:
        parts.append(lunar_text)

    ganzhi_text = build_ganzhi_text(row)
    if ganzhi_text:
        parts.append(ganzhi_text)

    return " ".join(parts)


def build_window_day_line2(row) -> str:
    huangli_text = build_huangli_text(row)
    gua_text = build_gua_text(row)

    left = f"黄历：{huangli_text}" if huangli_text else "黄历：（无）"
    right = f"卦象：{gua_text}" if gua_text else "卦象：（无）"

    return f"{left}    {right}"


def build_event_mask(df: pd.DataFrame, value_cols: list[str], time_col: str | None = None):
    cols = [col for col in value_cols if col in df.columns]
    if time_col and time_col in df.columns:
        cols.append(time_col)

    if not cols:
        return pd.Series([False] * len(df), index=df.index)

    return df[cols].map(has_value).any(axis=1)


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


def build_event_section(
    df: pd.DataFrame,
    center_date,
    range_start,
    range_end,
    title: str,
    value_cols: list[str],
    time_col: str | None = None,
):
    mask = build_event_mask(df, value_cols, time_col)

    in_range = df[mask & (df["日期"] >= range_start) & (df["日期"] <= range_end)].copy()
    before = df[mask & (df["日期"] < range_start)].copy()
    after = df[mask & (df["日期"] > range_end)].copy()

    prev_row = cast(Series | None, before.iloc[-1] if not before.empty else None)
    next_row = cast(Series | None, after.iloc[0] if not after.empty else None)

    lines = [f"{title}："]

    if prev_row is not None:
        lines.append(build_event_line(prev_row, center_date, value_cols, time_col))

    if not in_range.empty:
        for _, row in in_range.iterrows():
            lines.append(build_event_line(row, center_date, value_cols, time_col))

    if next_row is not None:
        lines.append(build_event_line(next_row, center_date, value_cols, time_col))

    if len(lines) == 1:
        return [f"{title}：（无）"]

    return lines


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

    lines = ["天象："]
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
        return ["天象：（无）"]

    return lines


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

    lines = ["行星逆行："]

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


@st.cache_data
def get_data():
    return load_data(FILE_PATH)


def get_default_date(df):
    min_date = df["日期"].min()
    max_date = df["日期"].max()
    today = datetime.today().date()
    if min_date <= today <= max_date:
        return today
    return max_date


st.set_page_config(page_title="我的日历本", layout="wide")
st.title("我的日历本")

df = get_data()
default_date = get_default_date(df)

tab1, tab2, tab3 = st.tabs(["当天/某天日历", "前后3天播报", "关键词查询"])

with tab1:
    query_date_input = st.date_input("选择日期", value=default_date)

    if isinstance(query_date_input, tuple):
        query_date = query_date_input[0] if len(query_date_input) > 0 else default_date
    elif query_date_input is None:
        query_date = default_date
    else:
        query_date = query_date_input
    row_df = df[df["日期"] == query_date]

    if row_df.empty:
        st.warning(f"{query_date} 没有记录。")
    else:
        row = row_df.iloc[0]
        text = row_to_lines(row, query_date, df)
        st.text(text)

with tab2:
    center_date_input = st.date_input("选择中心日期", value=default_date, key="center_date")

    if isinstance(center_date_input, tuple):
        center_date = center_date_input[0] if len(center_date_input) > 0 else default_date
    elif center_date_input is None:
        center_date = default_date
    else:
        center_date = center_date_input

    start_date = center_date - timedelta(days=3)
    end_date = center_date + timedelta(days=3)

    window_df = df[(df["日期"] >= start_date) & (df["日期"] <= end_date)].copy()

    st.subheader("7天主视图")
    for _, row in window_df.iterrows():
        st.text(build_window_day_line1(row, center_date))
        st.text(build_window_day_line2(row))
        st.divider()

    st.subheader("节气")
    for line in build_event_section(df, center_date, start_date, end_date, "节气", ["节气", "物候"], "时间点"):
        st.text(line)

    st.subheader("月相")
    for line in build_event_section(df, center_date, start_date, end_date, "月相", ["月相"], "月相时间"):
        st.text(line)

    st.subheader("天象")
    for line in build_astro_section(df, center_date):
        st.text(line)

    st.subheader("行星逆行")
    for line in build_retrograde_section(df, center_date):
        st.text(line)

with tab3:
    supported_keywords = [
        "节气", "月相", "地日", "紫孛",
        "水星", "金星", "火星", "木星", "土星",
        "天王星", "海王星", "冥王星"
    ]

    keyword = st.selectbox("选择关键词", supported_keywords)
    start_date_kw_input = st.date_input("起始日期", value=default_date, key="kw_date")

    if isinstance(start_date_kw_input, tuple):
        start_date_kw = start_date_kw_input[0] if len(start_date_kw_input) > 0 else default_date
    elif start_date_kw_input is None:
        start_date_kw = default_date
    else:
        start_date_kw = start_date_kw_input

    if keyword == "节气":
        mask = (
            (df["日期"] >= start_date_kw) &
            (df["节气"].map(has_value) | df["物候"].map(has_value) | df["时间点"].map(has_value))
        )
    elif keyword == "月相":
        mask = (
            (df["日期"] >= start_date_kw) &
            (df["月相"].map(has_value) | df["月相时间"].map(has_value))
        )
    else:
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
        mask = (
            (df["日期"] >= start_date_kw) &
            (df[value_col].map(has_value) | df[time_col].map(has_value))
        )

    result = df[mask].copy().head(10)

    if result.empty:
        st.info(f"从 {start_date_kw} 开始，未找到“{keyword}”的后续有效记录。")
    else:
        for _, row in result.iterrows():
            st.text(format_keyword_event_line(row, keyword))

