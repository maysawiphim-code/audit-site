"""ตรวจไฟล์บันทึกของหัวหน้าทีม (ตารางรายชั่วโมงของแต่ละผลัด)

สิ่งที่ทำ
1. อ่านยอด "ผลต่างรายชั่วโมง" ของคนและรถ แยกวันที่ 1 / วันที่ 2 เรียงตามชั่วโมง (เช้า → บ่าย → ดึก)
2. หา "ค่าดิป" คือชั่วโมงที่ยอดตกจากชั่วโมงก่อนหน้าเกินเกณฑ์ (ค่าเริ่มต้น 20%)
3. เทียบกับตารางรายชั่วโมงในไฟล์รายงาน ว่าดิปตกที่ชั่วโมงเดียวกันหรือไม่ และยอดรายชั่วโมงต่างกันกี่ %
"""
from __future__ import annotations

import io
import re

DIP = 0.20      # ยอดตกเกิน 20% ถือเป็นดิป
GAP = 0.20      # ยอดรายชั่วโมงต่างกันเกิน 20% ถือว่าไม่ใกล้เคียง


def _load(data: bytes, filename: str):
    if filename.lower().endswith(".xls"):
        import xlrd
        book = xlrd.open_workbook(file_contents=data)
        sh = book.sheet_by_index(0)
        return [[sh.cell(r, c).value if sh.cell(r, c).ctype != 0 else None
                 for c in range(sh.ncols)] for r in range(sh.nrows)]
    import openpyxl.styles.fonts as _f
    _f.Font.family.max = 100            # ไฟล์บางไฟล์ใส่ค่า font family เกินมาตรฐาน
    from openpyxl import load_workbook
    ws = load_workbook(io.BytesIO(data), data_only=True).worksheets[0]
    return [[c.value for c in row] for row in ws.iter_rows()]


def _num(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace(",", "").strip()
    return float(t) if re.fullmatch(r"-?\d*\.?\d+", t) else None


def parse_team_sheet(data: bytes, filename: str) -> dict:
    """คืนค่า {'days': [{'people': [...], 'cars': [...]}, ...], 'shifts': [...]}"""
    grid = _load(data, filename)
    starts, shifts = [], []
    for r, row in enumerate(grid):
        cols = [c for c, v in enumerate(row) if str(v).strip() == "ชม."]
        if not cols:
            continue
        name = ""
        if r > 0:
            for v in grid[r - 1]:
                if str(v).strip() in ("เช้า", "บ่าย", "ดึก"):
                    name = str(v).strip()
                    break
        starts.append((r, cols))
        shifts.append(name or f"ผลัดที่ {len(starts)}")

    days = []
    for r, cols in starts:
        for d, c0 in enumerate(cols):
            while len(days) <= d:
                days.append({"people": [], "cars": [], "hours": []})
            rr = r + 1
            while rr < len(grid):
                h = _num(grid[rr][c0]) if c0 < len(grid[rr]) else None
                if h is None or h != int(h) or not (1 <= h <= 24):
                    break
                get = lambda off: (_num(grid[rr][c0 + off]) if c0 + off < len(grid[rr]) else None) or 0.0
                days[d]["people"].append(get(4))
                days[d]["cars"].append(get(8))
                days[d]["hours"].append(f"{shifts[starts.index((r, cols))]} ชม.{int(h)}")
                rr += 1
    return {"days": days, "shifts": shifts}


def report_hourly(rep) -> dict:
    """ดึงตารางรายชั่วโมง (ค่าเฉลี่ย/วัน) จากชีตกราฟของไฟล์รายงาน"""
    for sheet in rep.sheets:
        grid = sheet.grid
        hr = next((r for r, row in enumerate(grid)
                   if any(str(v or "").strip() == "ช่วงเวลา" for v in row)), None)
        if hr is None:
            continue
        labcol = next(c for c, v in enumerate(grid[hr]) if str(v or "").strip() == "ช่วงเวลา")
        pcol = next((c for c, v in enumerate(grid[hr]) if "เฉลี่ย" in str(v or "")), labcol + 1)
        ccol = next((c for c, v in enumerate(grid[hr + 1]) if "จำนวนรถ" in str(v or "")), None)
        if ccol is None:
            ccol = next((c for c, v in enumerate(grid[hr]) if "พาหนะ" in str(v or "")), None)
        hours, people, cars = [], [], []
        for r in range(hr + 1, len(grid)):
            lab = str(grid[r][labcol] or "").strip() if labcol < len(grid[r]) else ""
            if not re.match(r"^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}", lab):
                continue
            hours.append(lab)
            people.append(_num(grid[r][pcol]) or 0.0)
            cars.append((_num(grid[r][ccol]) if ccol is not None and ccol < len(grid[r]) else None) or 0.0)
        if hours:
            return {"sheet": sheet.name, "hours": hours, "people": people, "cars": cars}
    return {}


def find_dips(series: list, threshold: float = DIP) -> list:
    """ชั่วโมงที่ยอดตกจากชั่วโมงก่อนหน้าเกินเกณฑ์"""
    out = []
    for i in range(1, len(series)):
        prev, cur = series[i - 1], series[i]
        if prev <= 0:
            continue
        change = (cur - prev) / prev
        if change <= -threshold:
            out.append({"i": i, "prev": prev, "cur": cur, "drop": change})
    return out


def compare(team: dict, report: dict, threshold: float = DIP, gap: float = GAP) -> dict:
    """เทียบค่าดิปและยอดรายชั่วโมงของไฟล์หัวหน้าทีมกับไฟล์รายงาน"""
    days = team.get("days") or []
    if not days or not report:
        return {}
    n = min([len(d["people"]) for d in days] + [len(report["people"])])
    avg = {k: [sum(d[k][i] for d in days) / len(days) for i in range(n)] for k in ("people", "cars")}
    rep_series = {k: report[k][:n] for k in ("people", "cars")}
    hours = report["hours"][:n]

    out = {"hours": hours, "kinds": {}}
    for kind, name in (("people", "คน"), ("cars", "รถ")):
        t, r = avg[kind], rep_series[kind]
        t_dips = {d["i"] for d in find_dips(t, threshold)}
        r_dips = {d["i"] for d in find_dips(r, threshold)}
        rows = []
        for i in range(n):
            diff = t[i] - r[i]
            pct = diff / r[i] if r[i] else (0.0 if not diff else None)
            rows.append({
                "ชั่วโมง": hours[i],
                f"หัวหน้าทีม (เฉลี่ย/วัน)": round(t[i], 2),
                "ไฟล์รายงาน": round(r[i], 2),
                "ต่าง": round(diff, 2),
                "ต่าง %": None if pct is None else round(pct * 100, 2),
                "ดิปในไฟล์หัวหน้าทีม": "✓" if i in t_dips else "",
                "ดิปในไฟล์รายงาน": "✓" if i in r_dips else "",
                "ตรงกัน": "" if i not in (t_dips | r_dips) else ("ตรงกัน" if i in (t_dips & r_dips) else "ไม่ตรงกัน"),
            })
        both, only_t, only_r = t_dips & r_dips, t_dips - r_dips, r_dips - t_dips
        big = [x for x in rows if x["ต่าง %"] is not None and abs(x["ต่าง %"]) > gap * 100]
        out["kinds"][name] = {"rows": rows, "both": sorted(both), "only_team": sorted(only_t),
                              "only_report": sorted(only_r), "big_gap": big,
                              "total_team": round(sum(t), 2), "total_report": round(sum(r), 2)}
    return out
