# app.py
# Full Streamlit system (ALL features)
# Inputs:
#   1) Timetable CSV (software export)
#   2) Lec Codes.xlsx (Code -> User name)
#   3) (Optional) CBS Template Excel (CBSTimeTableTemplate*.xlsx)
#
# Outputs:
#   A) ERP Week output (CSV + Excel) in Week1 format + Group Name column
#   B) CBS filled template (Excel) using "EntrySheet"
#
# Extras:
#   - Intake label is editable (text box)
#   - Column-wise editing via st.data_editor (download the edited file)
#   - Optional: auto-regenerate Group Name when intake label changes

import re
from io import BytesIO
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

# openpyxl is required for Excel template writing
import openpyxl

st.set_page_config(page_title="Timetable → ERP + CBS Outputs", layout="wide")
st.title("📅 Timetable CSV → ERP Upload + CBS Template Output (Full System)")

# -----------------------------
# ERP output columns (Week1 style) + Group Name added
# -----------------------------
ERP_OUTPUT_COLS = [
    "Activity Id",
    "Day",
    "Hour",
    "Students Sets",
    "Group Name",
    "Subject",
    "Teachers",
    "Teacher1",
    "Activity Tags",
    "Room",
    "Comments",
]

# -----------------------------
# CBS template columns (from your CBSTimeTableTemplate EntrySheet)
# -----------------------------
CBS_COLS = [
    "Cal Id",
    "Course",
    "Course Variant",
    "Section",
    "Room",
    "Faculty",
    "Day",
    "From Time Slot",
    "To Time Slot",
    "AcademyLocationID",
    "isAllFaculties",
]

# -----------------------------
# Helpers
# -----------------------------
def norm(s) -> str:
    if pd.isna(s) or s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def normalize_students_sets(s: str) -> str:
    return norm(s)

def split_teacher_codes(raw_teachers: str):
    # "BALA+SHK" -> ["BALA","SHK"]
    if pd.isna(raw_teachers) or raw_teachers is None:
        return []
    return [t.strip() for t in str(raw_teachers).split("+") if t.strip()]

def build_code_to_username_map(lec_df: pd.DataFrame) -> dict:
    """
    Lec Codes.xlsx expected columns:
      - Code
      - User name (ERP username like IITxxxx/Name or 7000xx/Name)
    """
    lec_df = lec_df.copy()
    lec_df.columns = [c.strip() for c in lec_df.columns]

    required = {"Code", "User name"}
    missing = [c for c in required if c not in lec_df.columns]
    if missing:
        raise ValueError(f"Lec Codes.xlsx missing columns: {missing}. Required: {sorted(required)}")

    lec_df["Code"] = lec_df["Code"].astype(str).map(norm).str.upper()
    lec_df["User name"] = lec_df["User name"].astype(str).map(norm)

    lec_df = lec_df.dropna(subset=["Code", "User name"]).drop_duplicates(subset=["Code"], keep="first")
    return dict(zip(lec_df["Code"], lec_df["User name"]))

def map_teachers(raw_teachers: str, code_map: dict):
    """
    Returns: main, teacher1, unknown(list), extras(list beyond 2)
    Unknown codes kept as-is (e.g., TBA 2)
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

def build_group_name(students_sets: str, intake_label: str):
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

def ensure_columns(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = ""
    return out[cols]

def df_to_excel_bytes(df: pd.DataFrame, sheet_name="ERP") -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return bio.getvalue()

def infer_course_variant(activity_tags: str) -> str:
    s = norm(activity_tags).upper()
    if "LAB" in s:
        return "Lab"
    if "TUT" in s:
        return "Tutorial"
    if "LEC" in s:
        return "Lecture"
    return ""

def parse_time_hhmm(t: str) -> datetime:
    return datetime.strptime(t.strip(), "%H:%M")

def fmt_time_hhmmss(dt: datetime) -> str:
    return dt.strftime("%H:%M:%S")

def build_sessions_for_cbs(df_hourly: pd.DataFrame) -> pd.DataFrame:
    """
    CBS needs From/To times. Your timetable is hourly.
    We group by session keys, merge consecutive hours:
      start = first hour
      end   = last hour + 1 hour
    """
    df = df_hourly.copy()
    df["Hour"] = df["Hour"].map(norm)

    session_keys = ["Activity Id", "Day", "Students Sets", "Group Name", "Subject", "Room", "Teachers", "Teacher1", "Activity Tags"]
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

def write_cbs_template(template_bytes: bytes, cbs_df: pd.DataFrame) -> bytes:
    """
    Fill template sheet 'EntrySheet' from row 2 down (row1 headers).
    Keeps template formatting and validations.
    """
    wb = openpyxl.load_workbook(BytesIO(template_bytes))
    if "EntrySheet" not in wb.sheetnames:
        raise ValueError("CBS template does not contain a sheet named 'EntrySheet'.")

    ws = wb["EntrySheet"]

    # Clear existing data rows (row 2 onwards)
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    # Write rows
    records = cbs_df.to_dict("records")
    for i, rec in enumerate(records, start=2):
        for j, col in enumerate(CBS_COLS, start=1):
            ws.cell(row=i, column=j).value = rec.get(col, "")

    out_bio = BytesIO()
    wb.save(out_bio)
    return out_bio.getvalue()

# -----------------------------
# UI: uploads
# -----------------------------
c1, c2, c3 = st.columns(3)
with c1:
    timetable_file = st.file_uploader("1) Timetable CSV (software export)", type=["csv"])
with c2:
    lec_codes_file = st.file_uploader("2) Lec Codes.xlsx", type=["xlsx"])
with c3:
    cbs_template_file = st.file_uploader("3) CBS Template Excel (optional)", type=["xlsx"])

st.subheader("Settings")
intake_label = st.text_input("Intake label for Group Name (user can change)", value="L5 Jan 26")
auto_regen_group = st.checkbox("Auto-regenerate Group Name using Intake label", value=True)
put_extras_in_comments = st.checkbox("If >2 lecturers, append extras into Comments", value=True)

st.subheader("CBS Settings (optional)")
duplicate_rows_for_teacher1_in_cbs = st.checkbox("CBS: If Teacher1 exists, create another CBS row", value=True)
default_is_all_faculties = st.selectbox("CBS: isAllFaculties", options=["FALSE", "TRUE"], index=0)

st.divider()

# -----------------------------
# Main processing
# -----------------------------
if timetable_file and lec_codes_file:
    # Load inputs
    try:
        df_raw = pd.read_csv(timetable_file)
    except Exception as e:
        st.error(f"Could not read timetable CSV: {e}")
        st.stop()

    try:
        lec_df = pd.read_excel(lec_codes_file)
        code_map = build_code_to_username_map(lec_df)
    except Exception as e:
        st.error(f"Could not read Lec Codes.xlsx or build mapping: {e}")
        st.stop()

    required_cols = ["Activity Id", "Day", "Hour", "Students Sets", "Subject", "Teachers", "Activity Tags", "Room"]
    missing = [c for c in required_cols if c not in df_raw.columns]
    if missing:
        st.error(f"Timetable CSV missing columns: {missing}")
        st.stop()

    out = df_raw.copy()
    out["Students Sets"] = out["Students Sets"].map(normalize_students_sets)

    if "Comments" not in out.columns:
        out["Comments"] = ""

    # Lecturer mapping
    unknown_all = set()
    extra_total = 0
    new_teachers, new_teacher1, new_comments = [], [], []

    for _, row in out.iterrows():
        main, t1, unknown, extras = map_teachers(row["Teachers"], code_map)
        new_teachers.append(main)
        new_teacher1.append(t1)

        for u in unknown:
            unknown_all.add(u)

        comment = norm(row.get("Comments", ""))

        if extras:
            extra_total += len(extras)
            if put_extras_in_comments:
                extra_txt = "Extra lecturers: " + " + ".join(extras)
                comment = (comment + " | " + extra_txt).strip(" |") if comment else extra_txt

        new_comments.append(comment)

    out["Teachers"] = new_teachers
    out["Teacher1"] = new_teacher1
    out["Comments"] = new_comments

    # Group Name
    if auto_regen_group:
        out["Group Name"] = out["Students Sets"].apply(lambda s: build_group_name(s, intake_label=intake_label))
    else:
        if "Group Name" not in out.columns:
            out["Group Name"] = ""

    # Enforce ERP output format + column order
    out_erp = ensure_columns(out, ERP_OUTPUT_COLS)

    st.success("✅ ERP table generated. You can edit it below before downloading.")

    if unknown_all:
        st.warning("Lecturer codes not found in Lec Codes.xlsx (kept as-is): " + ", ".join(sorted(unknown_all)))
    if extra_total > 0 and not put_extras_in_comments:
        st.warning("Extra lecturers beyond Teacher1 exist but were not included (enable checkbox).")

    # -----------------------------
    # Editable preview (column-wise editing)
    # -----------------------------
    st.subheader("Edit before download (click any cell to edit)")
    edited_erp = st.data_editor(
        out_erp,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
    )

    # Downloads (edited)
    st.subheader("Download edited ERP output")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "⬇️ Download ERP CSV (edited)",
            data=edited_erp.to_csv(index=False).encode("utf-8"),
            file_name="ERP_ready_week_edited.csv",
            mime="text/csv",
        )
    with d2:
        st.download_button(
            "⬇️ Download ERP Excel (edited)",
            data=df_to_excel_bytes(edited_erp, sheet_name="ERP"),
            file_name="ERP_ready_week_edited.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -----------------------------
    # CBS output (optional)
    # -----------------------------
    st.divider()
    st.subheader("CBS Template Output (optional)")

    # Build sessions from EDITED ERP (so CBS respects your edits)
    sessions = build_sessions_for_cbs(edited_erp)

    if sessions.empty:
        st.error("Could not build CBS sessions. Check 'Hour' format (must be like 08:30, 09:30...).")
        st.stop()

    # Create CBS rows
    cbs_rows = []
    for _, r in sessions.iterrows():
        base = {
            "Cal Id": "",
            "Course": norm(r.get("Subject", "")),
            "Course Variant": infer_course_variant(r.get("Activity Tags", "")),
            "Section": norm(r.get("Group Name", "")),  # relationship applied here
            "Room": norm(r.get("Room", "")),
            "Faculty": norm(r.get("Teachers", "")),
            "Day": norm(r.get("Day", "")),
            "From Time Slot": r.get("From Time Slot", ""),
            "To Time Slot": r.get("To Time Slot", ""),
            "AcademyLocationID": "",
            "isAllFaculties": default_is_all_faculties,
        }
        cbs_rows.append(base)

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

    st.write("CBS Preview (you can download filled template if you uploaded it):")
    st.dataframe(cbs_df, use_container_width=True)

    if cbs_template_file:
        try:
            filled_bytes = write_cbs_template(cbs_template_file.getvalue(), cbs_df)
            st.download_button(
                "⬇️ Download CBS Filled Template (Excel)",
                data=filled_bytes,
                file_name="CBS_Timetable_Filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Could not fill CBS template: {e}")
    else:
        st.info("Upload the CBS template Excel to download a filled version (format preserved).")

else:
    st.info("Upload at least: Timetable CSV + Lec Codes.xlsx")
