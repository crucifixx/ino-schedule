"""
Staff Scheduler - Streamlit App
Save as scheduler_app.py and run with: streamlit run scheduler_app.py
"""

import streamlit as st
from ortools.sat.python import cp_model
from dataclasses import dataclass
from typing import List, Dict, Tuple
import pandas as pd
from io import BytesIO
from collections import defaultdict
import datetime

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Staff Scheduler",
    page_icon="📅",
    layout="wide"
)

st.title("📅 Weekly Staff Scheduler")
st.markdown("Upload your employees → Adjust settings → Generate schedule")

# ============================================================
# SIDEBAR - SETTINGS & UPLOAD
# ============================================================
with st.sidebar:
    st.header("⚙️ Settings")

    full_time_min = st.number_input("Full-time minimum hours", value=40.0, step=1.0)
    part_time_max = st.number_input("Part-time maximum hours", value=29.5, step=0.5)
    preferred_max_shift = st.number_input("Preferred max shift length (hours)", value=6.0, step=0.5)
    lunch_threshold = st.number_input("Lunch starts after (hours)", value=5.5, step=0.5)
    lunch_minutes = st.number_input("Lunch length (minutes)", value=30, step=5)
    closing_min_people = st.number_input("Minimum people for closing", value=7, step=1)

    st.divider()
    st.header("👥 Employees")

    uploaded_file = st.file_uploader("Upload Employees (Excel or CSV)", type=["xlsx", "csv"])

    # Sample template download
    sample_data = pd.DataFrame({
        "Name": ["Alex Rivera", "Jordan Lee", "Taylor Kim", "Casey Morgan", "Riley Quinn"],
        "Level": [8, 8, 5, 4, 3],
        "FullTime": [True, True, True, True, False],
        "Available_Monday": ["08:00-22:00", "10:00-22:00", "05:00-18:00", "07:00-20:00", ""],
        "Available_Tuesday": ["08:00-22:00", "10:00-22:00", "05:00-18:00", "07:00-20:00", "09:00-17:00"],
        "Available_Wednesday": ["08:00-22:00", "10:00-22:00", "05:00-18:00", "07:00-20:00", ""],
        "Available_Thursday": ["08:00-22:00", "10:00-22:00", "05:00-18:00", "07:00-20:00", "09:00-17:00"],
        "Available_Friday": ["08:00-23:00", "12:00-23:00", "05:00-18:00", "07:00-22:00", "10:00-20:00"],
        "Available_Saturday": ["09:00-23:00", "10:00-23:00", "", "08:00-22:00", "10:00-20:00"],
        "Available_Sunday": ["", "10:00-20:00", "", "", "10:00-18:00"],
    })

    # Create sample Excel in memory
    sample_buffer = BytesIO()
    sample_data.to_excel(sample_buffer, index=False)
    sample_buffer.seek(0)

    st.download_button(
        label="📥 Download Sample Employee Template",
        data=sample_buffer,
        file_name="employee_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.caption("Level 8 = Manager | Levels 1-7 = Worker")
    st.caption("Availability format: 08:00-16:00  (leave blank if not available)")

# ============================================================
# HELPER FUNCTIONS
# ============================================================

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

def time_to_minutes(t: str) -> int:
    h, m = map(int, t.split(":"))
    return h * 60 + m

def shift_duration_hours(start: str, end: str) -> float:
    mins = time_to_minutes(end) - time_to_minutes(start)
    if mins < 0:
        mins += 24 * 60
    return mins / 60.0

def paid_hours(start: str, end: str, lunch_threshold: float, lunch_minutes: int) -> float:
    hours = shift_duration_hours(start, end)
    if hours >= lunch_threshold:
        hours -= lunch_minutes / 60.0
    return max(hours, 0)

def parse_availability(row) -> List[Tuple[str, str, str]]:
    availability = []
    for day in DAYS:
        col = f"Available_{day}"
        if col in row and pd.notna(row[col]) and str(row[col]).strip():
            value = str(row[col]).strip()
            if "-" in value:
                start, end = value.split("-")
                availability.append((day, start.strip(), end.strip()))
    return availability

@dataclass
class Employee:
    id: str
    name: str
    level: int
    full_time: bool
    availability: List[Tuple[str, str, str]]

@dataclass
class Shift:
    id: str
    day: str
    name: str
    start: str
    end: str
    required: Dict
    is_closing: bool = False
    needs_hands_free_manager: bool = False

def create_default_shifts() -> List[Shift]:
    """Default daily pattern – edit this list to match your real shifts."""
    shifts = []
    counter = 1

    # ---------- EDIT THIS PATTERN TO MATCH YOUR REAL NEEDS ----------
    base_pattern = [
        # (name, start, end, required, is_closing, needs_manager)
        ("Cleanup",       "05:00", "11:00", {"any_worker": 3}, False, False),
        ("Prep",          "07:00", "13:00", {"any_worker": 1}, False, False),
        ("Morning Open",  "08:00", "14:00", {"any_worker": 4}, False, False),
        ("Mid Morning",   "09:50", "15:50", {"any_worker": 2}, False, False),
        ("Late Morning",  "10:30", "16:30", {"any_worker": 2}, False, False),
        ("Afternoon",     "15:00", "21:00", {"any_worker": 3}, False, True),
        ("Night Manager", "16:30", "23:00", {8: 1},           True,  True),
        ("Evening",       "17:00", "23:00", {"any_worker": 4}, True,  False),
        ("Relief",        "19:00", "23:00", {"any_worker": 1}, True,  False),
        ("Late Evening",  "20:00", "23:00", {"any_worker": 3}, True,  False),
    ]
    # ----------------------------------------------------------------

    for day in DAYS:
        for name, start, end, required, is_closing, needs_mgr in base_pattern:
            shifts.append(Shift(
                id=f"S{counter:03d}",
                day=day,
                name=name,
                start=start,
                end=end,
                required=required,
                is_closing=is_closing,
                needs_hands_free_manager=needs_mgr
            ))
            counter += 1
    return shifts

def is_available(emp: Employee, day: str, start: str, end: str) -> bool:
    for avail_day, a_start, a_end in emp.availability:
        if avail_day != day:
            continue
        if time_to_minutes(a_start) <= time_to_minutes(start) and time_to_minutes(a_end) >= time_to_minutes(end):
            return True
    return False

# ============================================================
# SOLVER
# ============================================================

def solve_schedule(employees: List[Employee], shifts: List[Shift],
                   full_time_min, part_time_max, preferred_max_shift,
                   lunch_threshold, lunch_minutes, closing_min_people):

    model = cp_model.CpModel()

    assign = {}
    for emp in employees:
        for shift in shifts:
            assign[(emp.id, shift.id)] = model.NewBoolVar(f"a_{emp.id}_{shift.id}")

    # Availability
    for emp in employees:
        for shift in shifts:
            if not is_available(emp, shift.day, shift.start, shift.end):
                model.Add(assign[(emp.id, shift.id)] == 0)

    # No overlapping shifts
    for emp in employees:
        for day in DAYS:
            day_shifts = [s for s in shifts if s.day == day]
            for i, s1 in enumerate(day_shifts):
                for s2 in day_shifts[i+1:]:
                    start1, end1 = time_to_minutes(s1.start), time_to_minutes(s1.end)
                    start2, end2 = time_to_minutes(s2.start), time_to_minutes(s2.end)
                    if max(start1, start2) < min(end1, end2):
                        model.Add(assign[(emp.id, s1.id)] + assign[(emp.id, s2.id)] <= 1)

    # Hours limits
    for emp in employees:
        total = []
        for shift in shifts:
            hours = paid_hours(shift.start, shift.end, lunch_threshold, lunch_minutes)
            total.append(assign[(emp.id, shift.id)] * int(round(hours * 10)))
        total_hours = sum(total)

        if emp.full_time:
            model.Add(total_hours >= int(full_time_min * 10))
        else:
            model.Add(total_hours <= int(part_time_max * 10))

    # Coverage
    for shift in shifts:
        if 8 in shift.required:
            mgrs = [assign[(e.id, shift.id)] for e in employees if e.level == 8]
            model.Add(sum(mgrs) >= shift.required[8])

        if "any_worker" in shift.required:
            workers = [assign[(e.id, shift.id)] for e in employees if 1 <= e.level <= 7]
            model.Add(sum(workers) >= shift.required["any_worker"])

        for lvl, count in shift.required.items():
            if isinstance(lvl, int) and lvl != 8:
                lvl_vars = [assign[(e.id, shift.id)] for e in employees if e.level == lvl]
                model.Add(sum(lvl_vars) >= count)

    # Hands-free manager
    for shift in shifts:
        if shift.needs_hands_free_manager:
            mgrs = [assign[(e.id, shift.id)] for e in employees if e.level == 8]
            model.Add(sum(mgrs) >= 1)

    # Closing coverage
    for day in DAYS:
        closing_shifts = [s for s in shifts if s.day == day and s.is_closing]
        if not closing_shifts:
            continue
        closing_people = []
        for emp in employees:
            on_closing = model.NewBoolVar(f"close_{emp.id}_{day}")
            model.AddMaxEquality(on_closing, [assign[(emp.id, s.id)] for s in closing_shifts])
            closing_people.append(on_closing)
        model.Add(sum(closing_people) >= closing_min_people)

    # Soft objective: penalize long shifts
    penalties = []
    for emp in employees:
        for shift in shifts:
            duration = shift_duration_hours(shift.start, shift.end)
            if duration > preferred_max_shift:
                penalty = int((duration - preferred_max_shift) * 10)
                penalties.append(assign[(emp.id, shift.id)] * penalty)

    model.Minimize(sum(penalties))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 45.0
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None, solver

    results = []
    for emp in employees:
        for shift in shifts:
            if solver.Value(assign[(emp.id, shift.id)]) == 1:
                results.append({
                    "Employee": emp.name,
                    "Level": emp.level,
                    "FullTime": "Yes" if emp.full_time else "No",
                    "Day": shift.day,
                    "Shift": shift.name,
                    "Start": shift.start,
                    "End": shift.end,
                    "Paid Hours": round(paid_hours(shift.start, shift.end, lunch_threshold, lunch_minutes), 2)
                })

    return results, solver

# ============================================================
# MAIN APP LOGIC
# ============================================================

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_emp = pd.read_csv(uploaded_file)
        else:
            df_emp = pd.read_excel(uploaded_file)

        # Build employee list
        employees = []
        for idx, row in df_emp.iterrows():
            emp = Employee(
                id=f"E{idx+1:03d}",
                name=str(row["Name"]),
                level=int(row["Level"]),
                full_time=bool(row["FullTime"]),
                availability=parse_availability(row)
            )
            employees.append(emp)

        st.success(f"Loaded **{len(employees)}** employees")

        # Show preview
        with st.expander("Preview loaded employees"):
            st.dataframe(df_emp, use_container_width=True)

        shifts = create_default_shifts()

        if st.button("🚀 Generate Schedule", type="primary", use_container_width=True):
            with st.spinner("Solving schedule... this can take 15–45 seconds"):
                results, solver = solve_schedule(
                    employees, shifts,
                    full_time_min, part_time_max, preferred_max_shift,
                    lunch_threshold, lunch_minutes, closing_min_people
                )

            if results is None:
                st.error("No feasible schedule found. Try relaxing constraints or adding more staff availability.")
            else:
                st.success(f"Schedule created! Solver status: **{solver.StatusName()}**")

                df_results = pd.DataFrame(results)

                # Hours summary
                hours_summary = (
                    df_results.groupby(["Employee", "Level", "FullTime"])["Paid Hours"]
                    .sum()
                    .reset_index()
                    .sort_values(["FullTime", "Paid Hours"], ascending=[False, False])
                )

                # Tabs for viewing
                tab1, tab2, tab3 = st.tabs(["📅 Full Schedule", "⏱️ Hours Summary", "📥 Download"])

                with tab1:
                    for day in DAYS:
                        day_df = df_results[df_results["Day"] == day].sort_values(["Start", "Level"])
                        if not day_df.empty:
                            st.subheader(day)
                            st.dataframe(day_df[["Employee", "Level", "Shift", "Start", "End", "Paid Hours"]],
                                         use_container_width=True, hide_index=True)

                with tab2:
                    st.dataframe(hours_summary, use_container_width=True, hide_index=True)

                with tab3:
                    # Excel download
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                        df_results.to_excel(writer, sheet_name="Schedule", index=False)
                        hours_summary.to_excel(writer, sheet_name="Hours Summary", index=False)
                    excel_buffer.seek(0)

                    st.download_button(
                        label="📥 Download Full Schedule (Excel)",
                        data=excel_buffer,
                        file_name=f"schedule_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                    # CSV download
                    csv_buffer = df_results.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        label="📥 Download Schedule (CSV)",
                        data=csv_buffer,
                        file_name=f"schedule_{datetime.date.today()}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.info("Make sure your file has columns: Name, Level, FullTime, Available_Monday, Available_Tuesday, ...")

else:
    st.info("👈 Upload an employee file in the sidebar to get started.")
    st.markdown("""
    ### How to use this app
    1. Download the **Sample Employee Template** from the sidebar
    2. Fill it with your real employees (Name, Level, FullTime, Availability)
    3. Upload the file
    4. Adjust the settings if needed
    5. Click **Generate Schedule**
    """)