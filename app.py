"""เว็บแอปตรวจไฟล์รายงานจับตัวเลข — รันด้วย:  streamlit run app.py"""
import io
import re

import pandas as pd
import streamlit as st

import checker
import teamsheet

st.set_page_config(page_title="ตรวจรายงานจับตัวเลข", page_icon="📋", layout="wide")

st.markdown("""
<style>
.block-container{padding-top:2.2rem;max-width:1200px}
h1,h2,h3{font-family:"Bai Jamjuree","IBM Plex Sans Thai",sans-serif}
.small{color:#8A9AA8;font-size:.86rem}
</style>
""", unsafe_allow_html=True)

st.title("ตรวจไฟล์รายงานจับตัวเลขก่อนส่ง")
st.markdown(
    '<p class="small">ตรวจเฉพาะชีตที่มองเห็น (ข้ามชีตที่ซ่อนและชีต “สรุป (2)”) และเฉพาะช่องที่ใส่สีไว้ในไฟล์ '
    'เน้นความถูกต้องของยอดรวมคนและรถ ยอมให้คลาดเคลื่อนได้ ±0.5 — ไฟล์ถูกประมวลผลในหน่วยความจำ ไม่มีการเก็บไว้</p>',
    unsafe_allow_html=True)

deep = st.checkbox("ตรวจกราฟและรูปในไฟล์ด้วย (ต้องแปลงไฟล์ก่อน ใช้เวลาเพิ่มไฟล์ละไม่กี่วินาที)", value=True)

files = st.file_uploader("เลือกไฟล์รายงาน (.xls / .xlsx) ได้หลายไฟล์",
                         type=["xls", "xlsx", "xlsm"], accept_multiple_files=True)
if not files:
    st.info("อัปโหลดไฟล์เพื่อเริ่มตรวจ")
    st.stop()

tabs = st.tabs([f.name for f in files])

for tab, up in zip(tabs, files):
    with tab:
        raw = up.getvalue()
        try:
            rep = checker.audit(raw, up.name)
        except Exception as e:                      # noqa: BLE001
            st.error(f"อ่านไฟล์นี้ไม่ได้: {e} — ลองเปิดใน Excel แล้วบันทึกใหม่เป็น .xlsx")
            continue

        charts, shapes = [], []
        if deep:
            with st.spinner("กำลังตรวจกราฟและกล่องข้อความบนชีตแผนที่…"):
                try:
                    charts = checker.check_charts(raw, up.name, rep)
                    shapes = checker.check_shapes(raw, up.name, rep)
                except Exception as e:              # noqa: BLE001
                    st.warning(f"ตรวจกราฟ/กล่องข้อความไม่สำเร็จ: {e}")

        st.subheader(rep.site or up.name)
        st.caption(f"ชีตข้อมูลหลัก: {rep.data_name or '—'} · ตรวจ {len(rep.sheets)} ชีต · "
                   f"ข้ามชีต {len(rep.skipped)} ชีต" + (f" · ปี {rep.year}" if rep.year else ""))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("ต้องแก้", len(rep.bads))
        c2.metric("ควรตรวจซ้ำ", len(rep.warns))
        c3.metric("รายการที่ผ่าน", len(rep.passed))
        c4.metric("ชีตที่ข้าม", len(rep.skipped))
        if rep.skipped:
            st.caption("ชีตที่ข้าม: " + " · ".join(f"{n} ({why})" for n, why in rep.skipped))

        # ---------- โครงตารางที่อ่านได้ ----------
        with st.expander("โครงตารางที่ระบบอ่านได้จากชีต data (แถว/คอลัมน์ต่างกันตามไซต์)"):
            if rep.structure:
                st.dataframe(pd.DataFrame(rep.structure), use_container_width=True, hide_index=True)
                st.caption("ระบบไม่ได้ยึดตำแหน่งแถว/คอลัมน์ตายตัว แต่หาโครงจากหัวตารางและป้ายกำกับในไฟล์ "
                           "ถ้าตารางนี้ตรงกับไฟล์จริง แปลว่าอ่านโครงถูกต้อง")
            else:
                st.warning("อ่านโครงตารางไม่ได้ — ตรวจไฟล์นี้ด้วยตาไปก่อน")

        # ---------- แดชบอร์ดเทียบวันที่ 1 กับวันที่ 2 (เมนูเสริม) ----------
        if rep.day_rows and len(rep.day_names) >= 2:
            with st.expander("📊 แดชบอร์ดเทียบวันที่ 1 กับวันที่ 2 (สำหรับสรุปนำเสนอ)"):
                if st.button("แสดงแดชบอร์ด", key=f"dash-{up.name}"):
                    st.session_state[f"dash-on-{up.name}"] = True
                if st.session_state.get(f"dash-on-{up.name}"):
                    d1, d2 = rep.day_names[0], rep.day_names[1]
                    flag = st.slider("ธงเตือนเมื่อผลต่างเกิน (%)", 10, 100, 30, 5,
                                     key=f"flag-{up.name}")
                    totals = [x for x in rep.day_rows if x["เป็นยอดรวม"] and x["ช่วง"] == "รวมทั้งวัน"]
                    cols = st.columns(max(len(totals), 1))
                    for col, x in zip(cols, totals):
                        name_ = (x["รายการ"].replace("จำนวนคน (ยอดรวม)", "คน")
                                 .replace("จำนวนรถ (ยอดรวม)", "รถ"))
                        over = x["ต่าง %"] is not None and abs(x["ต่าง %"]) > flag
                        col.metric(("🚩 " if over else "") + name_, f'{x["values"][1]:,.0f}',
                                   f'{x["ต่าง"]:+,.0f} ({x["ต่าง %"]:+.1f}%)' if x["ต่าง %"] is not None else None)
                    st.caption(f"ตัวเลขใหญ่คือ {d2} · ตัวเลขเล็กคือผลต่างจาก {d1} · 🚩 = เกินเกณฑ์ {flag}%")

                    flagged = [x for x in rep.day_rows
                               if x["เป็นยอดรวม"] and x["ต่าง %"] is not None
                               and abs(x["ต่าง %"]) > flag and max(x["values"]) >= 10]
                    if flagged:
                        st.warning("**รายการที่ต่างกันเกินเกณฑ์ ควรอธิบายสาเหตุก่อนนำเสนอ**\n\n"
                                   + "\n".join(f'- {x["รายการ"]} · {x["ช่วง"]} : '
                                                f'{x["values"][0]:,.0f} → {x["values"][1]:,.0f} '
                                                f'({x["ต่าง %"]:+.1f}%)' for x in flagged))
                    else:
                        st.success(f"ยอดรวมของสองวันต่างกันไม่เกิน {flag}% ทุกรายการ")

                    shifts = [x for x in rep.day_rows if x["เป็นยอดรวม"] and x["ช่วง"] != "รวมทั้งวัน"]
                    if shifts:
                        chart = pd.DataFrame([{"รายการ": f'{x["รายการ"]} · {x["ช่วง"]}',
                                               d1: x["values"][0], d2: x["values"][1]}
                                              for x in shifts]).set_index("รายการ")
                        st.bar_chart(chart, height=320)

                    view = st.radio("ดูรายการ", ["เฉพาะยอดรวม", "ทุกรายการ"], horizontal=True,
                                    key=f"dv-{up.name}")
                    rows_ = rep.day_rows if view == "ทุกรายการ" else [x for x in rep.day_rows if x["เป็นยอดรวม"]]
                    table = pd.DataFrame([{
                        "รายการ": x["รายการ"], "ประเภท": x["ประเภท"], "ช่วง": x["ช่วง"],
                        d1: x["values"][0], d2: x["values"][1], "ต่าง": x["ต่าง"], "ต่าง %": x["ต่าง %"],
                        "ธง": "🚩" if (x["ต่าง %"] is not None and abs(x["ต่าง %"]) > flag
                                      and max(x["values"]) >= 10) else "",
                    } for x in rows_])
                    st.dataframe(table, use_container_width=True, hide_index=True)

                    lines = [f"# สรุปเปรียบเทียบรายวัน — {rep.site or up.name}", "",
                             f"- วันที่ 1: {d1}", f"- วันที่ 2: {d2}", ""]
                    for x in totals:
                        lines.append(f'- {x["รายการ"]}: {x["values"][0]:,.0f} → {x["values"][1]:,.0f} '
                                     f'({x["ต่าง %"]:+.1f}%)' if x["ต่าง %"] is not None else
                                     f'- {x["รายการ"]}: {x["values"][0]:,.0f} → {x["values"][1]:,.0f}')
                    if flagged:
                        lines += ["", f"ต้องอธิบายสาเหตุ (ต่างเกิน {flag}%):"]
                        lines += [f'- {x["รายการ"]} · {x["ช่วง"]} ({x["ต่าง %"]:+.1f}%)' for x in flagged]
                    c1_, c2_ = st.columns(2)
                    c1_.download_button("บันทึกตารางเป็น CSV",
                                        table.to_csv(index=False).encode("utf-8-sig"),
                                        file_name=f"เทียบรายวัน-{up.name.rsplit('.', 1)[0]}.csv",
                                        mime="text/csv", key=f"dlday-{up.name}")
                    c2_.download_button("บันทึกสรุปสำหรับนำเสนอ (.md)",
                                        "\n".join(lines).encode("utf-8"),
                                        file_name=f"สรุปนำเสนอ-{up.name.rsplit('.', 1)[0]}.md",
                                        mime="text/markdown", key=f"dlmd-{up.name}")

        # ---------- ยอดรวมคนและรถ ----------
        if rep.key_totals:
            st.markdown("#### ยอดรวมคนและรถที่ตรวจได้จากชีต data")
            st.dataframe(pd.DataFrame([{"รายการ": k["label"], "กลุ่ม": k["group"],
                                        "ยอดรวมทั้งวัน": k["value"], "เซลล์": k["cell"]}
                                       for k in rep.key_totals]),
                         use_container_width=True, hide_index=True)

        # ---------- กราฟทีละจุด ----------
        if charts:
            st.markdown("#### กราฟเส้นในชีตสรุป (3) และรายชั่วโมง")
            st.caption("แต่ละกราฟถูกไล่ตรวจทีละจุดว่าตรงกับเซลล์ต้นทางในชีตกราฟ "
                       "และผลรวมของทุกจุดต้องตรงกับยอดรวมในชีต data")
            st.dataframe(pd.DataFrame([{
                "กราฟอยู่ในชีต": c["sheet"],
                "หัวข้อ": c["title"] or c["file"],
                "ดึงข้อมูลจากชีต": c["src_sheet"],
                "จำนวนจุด": len(c["points"]),
                "จุดที่ไม่ตรง": len(c["mismatch"]),
                "ผลรวมทุกจุด": c["total"],
                "ตรงกับยอดรวม": c["match"]["label"] if c["match"] else "— ไม่ตรงกับยอดใด —",
            } for c in charts]), use_container_width=True, hide_index=True)

        # ---------- กล่องข้อความบนชีตแผนที่ ----------
        if shapes:
            wrong = [s_ for s_ in shapes if s_["status"] != "ตรง"]
            st.markdown(f"#### ตัวเลขในกล่องข้อความบนชีตแผนที่และชีตอื่น — ไม่ตรง {len(wrong)} จุด")
            st.caption("อ่านข้อความที่พิมพ์กำกับไว้บนรูปโดยตรง (ไม่ได้ใช้ OCR) แล้วเทียบกับยอดรวมในชีต data")
            st.dataframe(pd.DataFrame([{
                "ชีต": s_["sheet"], "ข้อความในกล่อง": s_["text"],
                "ตัวเลขในกล่อง": s_["value"],
                "ยอดจริงในชีต data": s_["expect"]["value"] if s_["expect"] else None,
                "รายการที่เทียบ": s_["expect"]["label"] if s_["expect"] else "—",
                "เซลล์ต้นทาง": s_["expect"]["cell"] if s_["expect"] else "—",
                "ผล": s_["status"],
            } for s_ in shapes]), use_container_width=True, hide_index=True)

        # ---------- ปัญหา ----------
        st.markdown("#### ผลตรวจ")
        if not rep.issues:
            st.success("ไม่พบข้อผิดพลาด — ไฟล์นี้ผ่านทุกเกณฑ์ที่ตรวจ")
        else:
            groups = {}
            for i in rep.issues:
                groups.setdefault(i.title, []).append(i)
            for title in sorted(groups, key=lambda t: (groups[t][0].sev != "bad", -len(groups[t]))):
                lst = groups[title]
                icon = "🔴" if lst[0].sev == "bad" else "🟡"
                with st.expander(f"{icon} {title} — {len(lst)} จุด "
                                 f"({', '.join(sorted({i.sheet for i in lst}))})",
                                 expanded=lst[0].sev == "bad"):
                    st.dataframe(pd.DataFrame([{"ชีต": i.sheet, "เซลล์": i.cell, "รายละเอียด": i.detail}
                                               for i in lst]),
                                 use_container_width=True, hide_index=True)

        if rep.passed:
            with st.expander(f"✅ รายการที่ตรวจแล้วถูกต้อง — {len(rep.passed)} รายการ"):
                st.dataframe(pd.DataFrame(rep.passed, columns=["ชีต", "รายการ"]),
                             use_container_width=True, hide_index=True)

        # ---------- ดาวน์โหลดผลตรวจ ----------
        rows = [{"ไฟล์": up.name, "ชีต": i.sheet, "เซลล์": i.cell,
                 "ระดับ": "ต้องแก้" if i.sev == "bad" else "ควรตรวจซ้ำ",
                 "ประเภท": i.title, "รายละเอียด": i.detail} for i in rep.issues]
        rows += [{"ไฟล์": up.name, "ชีต": s, "เซลล์": "", "ระดับ": "ผ่าน", "ประเภท": t, "รายละเอียด": ""}
                 for s, t in rep.passed]
        st.download_button("บันทึกผลตรวจเป็น CSV",
                           pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"ผลตรวจ-{up.name.rsplit('.', 1)[0]}.csv", mime="text/csv",
                           key=f"dl-{up.name}")

        # ---------- รูปในไฟล์ ----------
        with st.expander("ดูรูปที่ฝังอยู่ในไฟล์ (แผนที่ / กราฟ)"):
            if st.button("โหลดรูป", key=f"img-{up.name}"):
                for r_ in checker.check_images(raw, up.name, rep):
                    st.image(r_["image"], caption=f"รูปที่ {r_['index']}", use_container_width=True)

        # ---------- ไฟล์หัวหน้าทีม ----------
        st.markdown("#### ตรวจไฟล์บันทึกของหัวหน้าทีม (ค่าดิปรายชั่วโมง)")
        st.caption("ค่าดิป = ผลต่างระหว่างวันที่ 1 กับวันที่ 2 ของชั่วโมงนั้น "
                   "ระบบจะหาชั่วโมงที่ดิปเกินเกณฑ์ ตรวจว่าค่าดิปในไฟล์คำนวณถูกไหม "
                   "และเทียบยอดเฉลี่ย/วันกับตารางรายชั่วโมงของไฟล์รายงาน")
        col_a, col_b = st.columns([3, 1])
        team_file = col_a.file_uploader("ไฟล์หัวหน้าทีม (.xls / .xlsx)", type=["xls", "xlsx", "xlsm"],
                                        key=f"team-{up.name}")
        thr = col_b.slider("เกณฑ์ดิป %", 5, 50, 20, 5, key=f"thr-{up.name}") / 100
        if team_file:
            try:
                team = teamsheet.parse_team_sheet(team_file.getvalue(), team_file.name)
                res = teamsheet.analyse(team, teamsheet.report_hourly(rep), threshold=thr)
            except Exception as e:                # noqa: BLE001
                res = {}
                st.error(f"อ่านไฟล์หัวหน้าทีมไม่ได้: {e}")
            if not res:
                st.warning("อ่านตารางรายชั่วโมงไม่ได้ — ตรวจว่าไฟล์มีแถวหัวตารางคำว่า \"ชม.\" "
                           "และคอลัมน์ \"รายชั่วโมง\" ครบทุกผลัด")
            else:
                st.caption(f'อ่านได้ {res["n_hours"]} ชั่วโมง × {team["n_days"]} วัน '
                           f'({", ".join(team["shifts"])})')
                m1, m2, m3 = st.columns(3)
                m1.metric(f"ชั่วโมงที่ดิปเกิน {int(thr*100)}%", len(res["dips"]))
                m2.metric("ค่าดิปในไฟล์คำนวณผิด", len(res["formula_bad"]))
                m3.metric("ยอดต่างจากไฟล์รายงานเกิน 5%", len(res["gap"]))

                if res["dips"]:
                    st.warning("**ชั่วโมงที่ค่าดิปเกินเกณฑ์**\n\n" + "\n".join(
                        f'- {x["ผลัด"]} ชม.{x["ชม."]} {x["ช่วงเวลา"]} · {x["ประเภท"]} : '
                        f'{x["วันที่ 1"]:,.0f} → {x["วันที่ 2"]:,.0f} ({x["ค่าดิป %"]:.1f}%)'
                        for x in res["dips"]))
                else:
                    st.success(f'ไม่มีชั่วโมงที่ดิปเกิน {int(thr*100)}%')

                if res["formula_bad"]:
                    st.error("**ค่า %diff ในไฟล์ไม่ตรงกับที่คำนวณจากยอดสองวัน**\n\n" + "\n".join(
                        f'- {x["ผลัด"]} ชม.{x["ชม."]} · {x["ประเภท"]} : ในไฟล์ {x["ดิปในไฟล์ %"]}% '
                        f'แต่คำนวณได้ {x["ค่าดิป %"]}%' for x in res["formula_bad"]))
                if res["gap"]:
                    st.error("**ยอดในไฟล์หัวหน้าทีมไม่ตรงกับไฟล์รายงาน**\n\n" + "\n".join(
                        f'- {x["ผลัด"]} ชม.{x["ชม."]} {x["ช่วงเวลา"]} · {x["ประเภท"]} : '
                        f'เฉลี่ย/วัน {x["เฉลี่ย/วัน"]:,.1f} แต่ไฟล์รายงานใส่ {x["ไฟล์รายงาน"]:,.1f} '
                        f'(ต่าง {x["ต่างจากรายงาน"]:+,.1f})' for x in res["gap"]))
                elif res["has_report"]:
                    st.success("ยอดรายชั่วโมงของสองไฟล์สอดคล้องกันทุกชั่วโมง")

                only_dip = st.checkbox("แสดงเฉพาะชั่วโมงที่เกินเกณฑ์", key=f"od-{up.name}")
                table = pd.DataFrame(res["dips"] if only_dip else res["rows"])
                st.dataframe(table, use_container_width=True, hide_index=True)
                st.download_button("บันทึกผลตรวจค่าดิปเป็น CSV",
                                   table.to_csv(index=False).encode("utf-8-sig"),
                                   file_name=f"ค่าดิป-{up.name.rsplit('.', 1)[0]}.csv",
                                   mime="text/csv", key=f"dldip-{up.name}")

        # ---------- ดูข้อมูลในชีต ----------
        st.markdown("#### ดูข้อมูลในชีต")
        names = [s.name for s in rep.sheets]
        pick = st.selectbox("เลือกชีต", names, key=f"sel-{up.name}")
        sh = next(s for s in rep.sheets if s.name == pick)
        flagged = {i.cell: i.sev for i in rep.issues if i.sheet == pick and i.cell}
        width = max((len(r) for r in sh.grid), default=0)
        df = pd.DataFrame([[("" if v is None else v) for v in (row + [None] * (width - len(row)))]
                           for row in sh.grid[:300]],
                          columns=[checker.col_name(c) for c in range(width)])
        df.index = range(1, len(df) + 1)

        def paint(_):
            style = pd.DataFrame("", index=df.index, columns=df.columns)
            for cell, sev in flagged.items():
                m = re.fullmatch(r"([A-Z]+)(\d+)", str(cell).strip())
                if not m:                      # ปัญหาที่ไม่ผูกกับเซลล์ เช่น กล่องข้อความ/กราฟ
                    continue
                col, row = m.group(1), int(m.group(2))
                if col in style.columns and row in style.index:
                    style.loc[row, col] = ("background-color:#FF6B5A33;color:#FF6B5A"
                                           if sev == "bad" else "background-color:#FFC53D33;color:#B98600")
            return style

        if df.empty:
            st.info("ชีตนี้ไม่มีข้อมูลในตาราง")
        else:
            st.dataframe(df.style.apply(paint, axis=None), use_container_width=True, height=420)
        if len(sh.grid) > 300:
            st.caption(f"แสดง 300 แถวแรกจากทั้งหมด {len(sh.grid)} แถว")

st.divider()
st.caption("เกณฑ์ที่ใช้ตรวจ (เฉพาะช่องที่ใส่สี คลาดเคลื่อนได้ ±0.5): ยอดรวมคน/รถทั้งหมด = ผลรวมยอดรวมของทุกจุด · "
           "ผลรวมกลุ่มย่อย = ยอดรวม · ผลัดเช้า+บ่าย+ดึก = รวมทั้งวัน · คอลัมน์เฉลี่ยของหลายวัน · %สัดส่วน · "
           "ตารางรายชั่วโมงรวมยอดตรง · ตัวเลขในชีตสรุปต้องมาจากชีต data · ชื่อทำเล วันที่ ปี พ.ศ. ตรงกันทุกชีต · "
           "เซลล์สูตรผิดพลาด · ตัวเลขในกล่องข้อความบนชีตแผนที่")
