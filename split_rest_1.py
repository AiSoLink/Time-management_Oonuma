"""
分割休息①
ランキング・休息採用・運行間休息・休憩合算の処理
"""

from datetime import datetime, timedelta, date, time
from typing import Any, List, Optional

EXCEL_EPOCH = datetime(1899, 12, 30)


def _try_to_date(v: Any) -> Optional[datetime]:
    """日付として解釈可能なら datetime を返す"""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v
    try:
        if isinstance(v, date) and not isinstance(v, datetime):
            return datetime.combine(v, time.min)
    except Exception:
        pass
    if isinstance(v, str):
        v = v.strip()
        for fmt in ("%Y/%m/%d %H:%M:%S.%f", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M",
                    "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
                    "%Y/%m/%d", "%Y-%m-%d"):
            try:
                return datetime.strptime(v, fmt)
            except ValueError:
                continue
    try:
        n = float(v)
        if n > 0:
            return EXCEL_EPOCH + timedelta(days=n)
    except (TypeError, ValueError):
        pass
    return None


def _time_to_seconds(v: Any) -> float:
    """時刻を秒数に変換。Excelシリアル or "HH:MM" or "HH:MM:SS" """
    if v is None or v == "":
        return 0.0
    try:
        if isinstance(v, (int, float)):
            frac = float(v) - int(v)
            if frac < 0:
                frac += 1
            return frac * 86400.0
    except (TypeError, ValueError):
        pass
    if isinstance(v, str):
        parts = v.strip().split(":")
        if len(parts) >= 2:
            h, m = float(parts[0]), float(parts[1])
            s = float(parts[2]) if len(parts) >= 3 else 0
            return h * 3600 + m * 60 + s
    return 0.0


def _seconds_to_excel_serial(sec: float) -> float:
    """秒数をExcel時刻シリアル値に"""
    return sec / 86400.0


def _find_col(header_row: List, name: str) -> int:
    """ヘッダー行から列名のインデックスを取得（0-based）"""
    for i, h in enumerate(header_row):
        if str(h).strip() == name:
            return i
    return -1


def _ensure_len(row: List, n: int) -> None:
    while len(row) <= n:
        row.append(None)


def process_split_rest_step1(rows: List[List]) -> None:
    """
    分割休息①：ランキング・休息採用・運行間休息・休憩合算
    rows を in-place で更新。ヘッダー行(0) + データ行(1以降)
    """
    if len(rows) < 2:
        return

    header = rows[0]
    idx = {}
    for name in ["乗務員名", "運行日", "出庫日時", "帰庫日時", "休憩", "休息",
                 "休息採用", "運行間休息", "ランキング"]:
        idx[name] = _find_col(header, name)
        if idx[name] < 0:
            return

    last_row = len(rows) - 1
    r = 1

    while r <= last_row:
        name_cur = str(rows[r][idx["乗務員名"]]).strip() if idx["乗務員名"] < len(rows[r]) else ""

        g_start = r
        g_end = r
        while g_end <= last_row:
            nm = str(rows[g_end][idx["乗務員名"]]).strip() if idx["乗務員名"] < len(rows[g_end]) else ""
            if nm != name_cur:
                break
            g_end += 1
        g_end -= 1

        # 付け足し行除外：運行日の日が21以上の行を末尾から除外
        effective_end = g_end
        while effective_end >= g_start:
            dt = _try_to_date(rows[effective_end][idx["運行日"]] if idx["運行日"] < len(rows[effective_end]) else None)
            if dt is None:
                break
            if dt.day >= 21:
                effective_end -= 1
            else:
                break

        work_days = max(0, effective_end - g_start + 1)

        # クリア
        for i in range(g_start, g_end + 1):
            _ensure_len(rows[i], max(idx["休息採用"], idx["運行間休息"], idx["ランキング"]) + 1)
            rows[i][idx["休息採用"]] = ""
            rows[i][idx["運行間休息"]] = ""
            rows[i][idx["ランキング"]] = ""

        # ランキング対象収集
        rank_rows = []
        rank_secs = []

        for i in range(g_start, effective_end + 1):
            row = rows[i]
            _ensure_len(row, max(idx["帰庫日時"], idx["出庫日時"], idx["休息"]) + 1)

            inter_min = None
            if i < g_end:
                dt_in = _try_to_date(row[idx["帰庫日時"]])
                dt_out_next = _try_to_date(rows[i + 1][idx["出庫日時"]] if idx["出庫日時"] < len(rows[i + 1]) else None)
                if dt_in and dt_out_next:
                    delta = dt_out_next - dt_in
                    inter_min = int(delta.total_seconds() / 60)
                    if inter_min < 0:
                        inter_min = None

            if inter_min is not None:
                row[idx["運行間休息"]] = inter_min
            else:
                row[idx["運行間休息"]] = ""

            # ランキング対象判定
            if inter_min is not None:
                dt1 = _try_to_date(row[idx["運行日"]])
                dt2 = _try_to_date(rows[i + 1][idx["運行日"]] if i + 1 <= last_row and idx["運行日"] < len(rows[i + 1]) else None)
                day_diff = 1
                if dt1 and dt2:
                    d = (dt2.date() - dt1.date()).days
                    day_diff = max(1, d)

                baseline = 540 + (day_diff - 1) * 1440
                rest_sec = _time_to_seconds(row[idx["休息"]])

                if inter_min < baseline and rest_sec > 0:
                    rank_rows.append(i)
                    rank_secs.append(rest_sec)

        # 休息秒降順、同値は行番号昇順でソート
        if len(rank_rows) > 1:
            pairs = list(zip(rank_rows, rank_secs))
            pairs.sort(key=lambda x: (-x[1], x[0]))
            rank_rows = [p[0] for p in pairs]
            rank_secs = [p[1] for p in pairs]

        # ランキング出力
        for k, ri in enumerate(rank_rows, 1):
            rows[ri][idx["ランキング"]] = k

        # 採用数 = 出勤日数の半分（切り捨て）
        adopt_max = min(work_days // 2, len(rank_rows))

        for k in range(adopt_max):
            ri = rank_rows[k]
            rows[ri][idx["休息採用"]] = rows[ri][idx["休息"]]

        # 非採用休息→休憩合算（付け足し行は触らない）
        for rr in range(g_start, effective_end + 1):
            row = rows[rr]
            _ensure_len(row, max(idx["休息採用"], idx["休憩"], idx["休息"]) + 1)
            adopt_val = row[idx["休息採用"]]
            if adopt_val is None or str(adopt_val).strip() == "":
                rest_sec = _time_to_seconds(row[idx["休息"]])
                if rest_sec > 0:
                    break_sec = _time_to_seconds(row[idx["休憩"]])
                    row[idx["休憩"]] = _seconds_to_excel_serial(break_sec + rest_sec)
                    row[idx["休息"]] = 0

        r = g_end + 1
