import re
from io import BytesIO
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st
import openpyxl

st.set_page_config(page_title="Timetable → ERP + CBS Outputs", layout="wide")
st.title("📅 Timetable CSV → ERP Upload + CBS Template Output (Full System)")

# -----------------------------
# ERP Week output schema (your Week1 format)
# -----------------------------
ERP_OUTPUT_COLS = [
    "Activity Id", "Day", "Hour", "Students Sets", "Group Name", "Subject",
    "Teachers", "Teacher1", "Activity Tags", "Room", "Comments"
]

# CBS Template headers (from your template EntrySheet)
CBS_COLS = [
    "Cal Id", "Course", "Course Variant", "Section", "Room", "Faculty",
    "Day", "From Time Slot", "To Time Slot", "AcademyLocationID", "isAllFaculties"
]

# -----------------------------
# Helpers
# -----------------------------
def norm(s) -> str:
    if pd.isna(s) or s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def split_teacher_codes(raw_teachers: str):
    # "BALA+SHK" -> ["BALA","SHK"]
    if pd.isna(raw_teachers) or raw_teachers is None:
        return []
    return [t.strip() for t in str(raw_teachers).split("+") if t.strip()]

def build_code_to_username_map(lec_df: pd.DataFrame) -> dict:
    lec_df = lec_df.copy()
    lec_df.columns = [c.strip() for c in lec_df.columns]
    required = {"Code", "User name"}
    missing = [c for c in required if c not in lec_df.columns]
    if missing:
        raise ValueError(f"Lec Codes.xlsx missing columns: {missing} (need Code + User name)")

    lec_df["Code"] = lec_df["Code"].astype(str).map(norm).str.upper()
    lec_df["User name"] = lec_df["User name"].astype(str).map(norm)

    lec_df = lec_df.dropna(subset=["Code", "User name"]).drop_duplicates(subset=["Code"], keep="first")
    return dict(zip(lec_df["Code"], lec_df["User name"]))

def map_teachers(raw_teachers: str, code_map: dict):
    """
    Returns:
      Teachers(main), Teacher1(second), unknown(list), extras(list beyond 2)
    Unknown kept as-is (e.g., TBA 2)
    """
    codes = split_teacher_codes(raw_teachers)
    mapped, unknown = [], []
    for c in codes:
        key = c.upper()
        if key in code_map:
            mapped.append(code_map[key])
        else:
            unknown.append(c)
            mapped.append(c)

    main = mapped[0] if len(mapped) >= 1 else ""
    t1 = mapped[1] if len(mapped) >= 2 else ""
    extras = mapped[2:] if len(mapped) > 2 else []
    return main, t1, unknown, extras

# ----- Group Name relationship (Students Sets -> PDF style group name)
def extract_groups(students_sets: str):
    """
    From: 'L5 CS -G1+L5 CS -G2+L5 CS -G21+L5 SE -G10'
    To: {'CS':[1,2,21], 'SE':[10]}
    """
    out = {}
    s = norm(students_sets)
    if not s:
        return out

    for part in s.split("+"):
        part = part.strip()
        m = re.search(r"\bL5\s+(CS|SE)\s*-\s*G(\d+)\b", part, flags=re.IGNORECASE)
        if not m:
            continue
        prog = m.group(1).upper()
        g = int(m.group(2))
        out.setdefault(prog, []).append(g)

    for prog in out:
        out[prog] = sorted(set(out[prog]))
    return out

def build_group_name(students_sets: str, intake_label: str = "L5 Jan 26"):
    """
    PDF style:
      - L5 Jan 26 SE- 1,2,4
      - L5 Jan 26 SE- 10 / CS- 1,2,21
      - L5 Jan 26 CS- 22,23,24,4
    """
    groups = extract_groups(students_sets)
    if not groups:
        return ""

    def fmt(prog):
        nums = ",".join(str(n) for n in groups[prog])
        return f"{intake_label} {prog}- {nums}"

    has_se = "SE" in groups
    has_cs = "CS" in groups
    if has_se and has_cs:
        return f"{fmt('SE')} / {fmt('CS')}"
    if has_se:
        return fmt("SE")
    return fmt("CS")

# ----- Course Variant from Activity Tags
def infer_course_variant(activity_tags: str) -> str:
    s = norm(activity_tags).upper()
    # You can extend rules here
    if "LAB" in s:
        return "Lab"
    if "TUT" in s:
        return "Tutorial"
    if "LEC" in s:
        return "Lecture"
    return ""

# ----- Build session blocks (From/To) from hourly rows
def parse_time_hhmm(t: str) -> datetime:
    # expects "08:30"
    return datetime.strptime(t.strip(), "%H:%M")

def fmt_time_hhmmss(dt: datetime) -> str:
    return dt.strftime("%H:%M:%S")

def build_sessions_for_cbs(df_erp_like: pd.DataFrame):
    """
    Your Week-style data has one row per hour.
    CBS template needs From/To timeslots per session.
    Group by session keys, merge consecutive hours, compute end = last_hour + 1 hour.
    """
    df = df_erp_like.copy()
    df["Hour"] = df["Hour"].map(norm)

    session_keys = ["Activity Id", "Day", "Students Sets", "Subject", "Room", "Teachers", "Teacher1", "Activity Tags"]
    for c in session_keys:
        if c not in df.columns:
            df[c] = ""

    sessions = []
    for _, g in df.groupby(session_keys, dropna=False):
        hours = sorted([parse_time_hhmm(x) for x in g["Hour"].dropna().unique() if x])
        if not hours:
            continue
        start = hours[0]
        end = hours[-1] + timedelta(hours=1)

        row = g.iloc[0].to_dict()
        row["From Time Slot"] = fmt_time_hhmmss(start)
        row["To Time Slot"] = fmt_time_hhmmss(end)
        sessions.append(row)

    return pd.DataFrame(sessions)

def df_to_excel_bytes_simple(df: pd.DataFrame, sheet_name="ERP") -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return bio.getvalue()

def write_cbs_template(template_bytes: bytes, cbs_df: pd.DataFrame) -> bytes:
    """
    Fill template EntrySheet starting row 2 (row 1 = headers).
    Keep template formatting & validations.
    """
    bio_in = BytesIO(template_bytes)
    wb = openpyxl.load_workbook(bio_in)
    ws = wb["EntrySheet"]

    # Clear existing rows from row 2 downward (safe)
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    # Write rows
    for i, rec in enumerate(cbs_df.to_dict("records"), start=2):
        for j, col in enumerate(CBS_COLS, start=1):
            ws.cell(row=i, column=j).value = rec.get(col, "")

    out_bio = BytesIO()
    wb.save(out_bio)
    return out_bio.getvalue()

# -----------------------------
# UI
# -----------------------------
c1, c2, c3 = st.columns(3)
with c1:
    timetable_file = st.file_uploader("1) Timetable CSV (software export)", type=["csv"])
with c2:
    lec_codes_file = st.file_uploader("2) Lec Codes.xlsx", type=["xlsx"])
with c3:
    cbs_template_file = st.file_uploader("3) CBS Template Excel (your file)", type=["xlsx"])

st.subheader("Settings")
intake_label = st.text_input("Intake label for Group Name", value="L5 Jan 26")
put_extras_in_comments = st.checkbox("If >2 lecturers, append extras into Comments", value=True)
duplicate_rows_for_teacher1_in_cbs = st.checkbox("CBS: If Teacher1 exists, create another CBS row", value=True)

st.divider()

# -----------------------------
# Run
# -----------------------------
if timetable_file and lec_codes_file:
    # Load inputs
    df_raw = pd.read_csv(timetable_file)
    lec_df = pd.read_excel(lec_codes_file)
    code_map = build_code_to_username_map(lec_df)

    # Validate timetable columns
    required_cols = ["Activity Id", "Day", "Hour", "Students Sets", "Subject", "Teachers", "Activity Tags", "Room"]
    missing = [c for c in required_cols if c not in df_raw.columns]
    if missing:
        st.error(f"Timetable CSV missing columns: {missing}")
        st.stop()

    out = df_raw.copy()
    out["Students Sets"] = out["Students Sets"].map(norm)
    out["Comments"] = out["Comments"] if "Comments" in out.columns else ""

    # Lecturer mapping
    unknown_all = set()
    extra_total = 0
    t_main, t1_list, comm_list = [], [], []

    for _, row in out.iterrows():
        main, t1, unknown, extras = map_teachers(row["Teachers"], code_map)
        t_main.append(main)
        t1_list.append(t1)
        for u in unknown:
            unknown_all.add(u)

        comment = norm(row.get("Comments", ""))
        if extras:
            extra_total += len(extras)
            if put_extras_in_comments:
                extra_txt = "Extra lecturers: " + " + ".join(extras)
                comment = (comment + " | " + extra_txt).strip(" |") if comment else extra_txt
        comm_list.append(comment)

    out["Teachers"] = t_main
    out["Teacher1"] = t1_list
    out["Comments"] = comm_list

    # Group Name
    out["Group Name"] = out["Students Sets"].apply(lambda s: build_group_name(s, intake_label=intake_label))

    # Enforce ERP output columns
    for c in ERP_OUTPUT_COLS:
        if c not in out.columns:
            out[c] = ""
    out_erp = out[ERP_OUTPUT_COLS].copy()

    st.success("✅ ERP Week output generated.")
    if unknown_all:
        st.warning("Lecturer codes not found in Lec Codes.xlsx (kept as-is): " + ", ".join(sorted(unknown_all)))
    if extra_total > 0 and not put_extras_in_comments:
        st.warning("Extra lecturers beyond Teacher1 were found but not included (enable checkbox).")

    st.subheader("Preview: ERP Week Output")
    st.dataframe(out_erp, use_container_width=True)

    # Downloads for ERP
    colA, colB = st.columns(2)
    with colA:
        st.download_button(
            "⬇️ Download ERP CSV",
            data=out_erp.to_csv(index=False).encode("utf-8"),
            file_name="ERP_ready_week.csv",
            mime="text/csv",
        )
    with colB:
        st.download_button(
            "⬇️ Download ERP Excel",
            data=df_to_excel_bytes_simple(out_erp, sheet_name="ERP"),
            file_name="ERP_ready_week.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -----------------------------
    # CBS Output
    # -----------------------------
    st.divider()
    st.subheader("CBS Template Output")

    # Build session-level rows (From/To)
    sessions = build_sessions_for_cbs(out_erp)

    if sessions.empty:
        st.error("Could not build CBS sessions (check Hour column format like 08:30, 09:30, etc.)")
        st.stop()

    # Build CBS dataframe with correct headers
    cbs_rows = []
    for _, r in sessions.iterrows():
        base = {
            "Cal Id": "",  # leave blank unless you have a value
            "Course": norm(r.get("Subject", "")),
            "Course Variant": infer_course_variant(r.get("Activity Tags", "")),
            "Section": norm(r.get("Group Name", "")),          # ✅ relationship applied here
            "Room": norm(r.get("Room", "")),
            "Faculty": norm(r.get("Teachers", "")),
            "Day": norm(r.get("Day", "")),
            "From Time Slot": r.get("From Time Slot", ""),
            "To Time Slot": r.get("To Time Slot", ""),
            "AcademyLocationID": "",  # optional (you can map from room prefix if needed)
            "isAllFaculties": "FALSE"
        }
        cbs_rows.append(base)

        # If Teacher1 exists, optionally create another row
        t1 = norm(r.get("Teacher1", ""))
        if duplicate_rows_for_teacher1_in_cbs and t1:
            base2 = base.copy()
            base2["Faculty"] = t1
            cbs_rows.append(base2)

    cbs_df = pd.DataFrame(cbs_rows)
    for c in CBS_COLS:
        if c not in cbs_df.columns:
            cbs_df[c] = ""
    cbs_df = cbs_df[CBS_COLS].copy()

    st.write("Preview: CBS rows")
    st.dataframe(cbs_df, use_container_width=True)

    if cbs_template_file:
        filled_bytes = write_cbs_template(cbs_template_file.getvalue(), cbs_df)
        st.download_button(
            "⬇️ Download CBS Filled Template (Excel)",
            data=filled_bytes,
            file_name="CBS_Timetable_Filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Upload the CBS template Excel to download a filled version (format preserved).")

else:
    st.info("Upload at least: Timetable CSV + Lec Codes.xlsx")
