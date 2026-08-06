"""เว็บแอปตรวจไฟล์รายงานจับตัวเลข — รันด้วย:  streamlit run app.py"""
import io

import pandas as pd
import streamlit as st

import checker

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

        charts = []
        if deep:
            with st.spinner("กำลังตรวจกราฟในชีตสรุป (3) และรายชั่วโมง…"):
                try:
                    charts = checker.check_charts(raw, up.name, rep)
                except Exception as e:              # noqa: BLE001
                    st.warning(f"ตรวจกราฟไม่สำเร็จ: {e}")

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

        # ---------- ชีตแผนที่ / รูปภาพ ----------
        st.markdown("#### ชีตแผนที่และรูปกราฟ")
        st.caption("อ่านตัวเลขในรูปด้วย OCR แล้วเทียบกับยอดรวมในชีต data — OCR อ่านผิดได้ "
                   "จึงแสดงรูปคู่กันไว้ให้ยืนยันด้วยตาเสมอ")
        if st.button("ตรวจรูปในไฟล์นี้", key=f"img-{up.name}"):
            with st.spinner("กำลังดึงรูปและอ่านตัวเลข…"):
                try:
                    imgs = checker.check_images(raw, up.name, rep)
                except Exception as e:              # noqa: BLE001
                    imgs, err = [], str(e)
                    st.warning(f"ดึงรูปไม่สำเร็จ: {err}")
            if not imgs:
                st.info("ไม่พบรูปในไฟล์ หรือเครื่องนี้ไม่มี LibreOffice/Tesseract สำหรับแปลงไฟล์ .xls และอ่านตัวเลข")
            for r in imgs:
                st.image(r["image"], caption=f"รูปที่ {r['index']}", use_container_width=True)
                a, b = st.columns(2)
                a.success("ยอดรวมที่พบในรูป: " +
                          (", ".join(f'{k["label"]} = {checker.fmt(k["value"])}' for k in r["found"])
                           or "— ไม่พบ —"))
                b.warning("ยอดรวมที่ยังไม่พบในรูป (ตรวจด้วยตาอีกครั้ง): " +
                          (", ".join(checker.fmt(k["value"]) for k in r["missing"]) or "— ไม่มี —"))
                if r["odd"]:
                    st.error("ตัวเลขหลักพันในรูปที่ไม่ตรงกับชีต data: " +
                             ", ".join(checker.fmt(v) for v in r["odd"]))

        # ---------- ดูข้อมูลในชีต ----------
        st.markdown("#### ดูข้อมูลในชีต")
        names = [s.name for s in rep.sheets]
        pick = st.selectbox("เลือกชีต", names, key=f"sel-{up.name}")
        sh = next(s for s in rep.sheets if s.name == pick)
        flagged = {i.cell: i.sev for i in rep.issues if i.sheet == pick}
        width = max((len(r) for r in sh.grid), default=0)
        df = pd.DataFrame([[("" if v is None else v) for v in (row + [None] * (width - len(row)))]
                           for row in sh.grid[:300]],
                          columns=[checker.col_name(c) for c in range(width)])
        df.index = range(1, len(df) + 1)

        def paint(_):
            style = pd.DataFrame("", index=df.index, columns=df.columns)
            for cell, sev in flagged.items():
                col = "".join(ch for ch in cell if ch.isalpha())
                row = int("".join(ch for ch in cell if ch.isdigit()))
                if col in style.columns and row in style.index:
                    style.loc[row, col] = ("background-color:#FF6B5A33;color:#FF6B5A"
                                           if sev == "bad" else "background-color:#FFC53D33;color:#B98600")
            return style

        st.dataframe(df.style.apply(paint, axis=None), use_container_width=True, height=420)
        if len(sh.grid) > 300:
            st.caption(f"แสดง 300 แถวแรกจากทั้งหมด {len(sh.grid)} แถว")

st.divider()
st.caption("เกณฑ์ที่ใช้ตรวจ (เฉพาะช่องที่ใส่สี คลาดเคลื่อนได้ ±0.5): ยอดรวมคน/รถทั้งหมด = ผลรวมยอดรวมของทุกจุด · "
           "ผลรวมกลุ่มย่อย = ยอดรวม · ผลัดเช้า+บ่าย+ดึก = รวมทั้งวัน · คอลัมน์เฉลี่ยของหลายวัน · %สัดส่วน · "
           "ตารางรายชั่วโมงรวมยอดตรง · ตัวเลขในชีตสรุปต้องมาจากชีต data · ชื่อทำเล วันที่ ปี พ.ศ. ตรงกันทุกชีต · "
           "เซลล์สูตรผิดพลาด · ตัวเลขในรูปเทียบกับยอดรวม (OCR)")
