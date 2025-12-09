
import streamlit as st
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ==========================================
# Streamlit Config & File Path
# ==========================================
st.set_page_config(page_title="Learning Objective Mapping Form", layout="centered")
# Excel file is in the same directory as app.py
FILE = Path("LOreferenceData_final_formfeedversion2.xlsx")

# ==========================================
# Custom CSS for Section Styling
# ==========================================
st.markdown("""
    <style>
        .title-style {
            font-size: 28px;
            font-weight: bold;
            color: #3366cc;
        }
        .section-box {
            background-color: #f5f8ff;
            border-left: 5px solid #3366cc;
            padding: 1rem;
            margin-bottom: 1rem;
            border-radius: 6px;
        }
        .stSelectbox > label, .stTextInput > label, .stTextArea > label {
            font-weight: bold;
            color: #333;
        }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# Reference Data Caching
# ==========================================
@st.cache_data
def load_reference_data():
    sheets = {
        "courses": "tb_courses",
        "bloomlevel": "tb_bloomlevel",
        "activity": "tb_activity",
        "methods": "tb_methods",
        "difficulty": "tb_difficulty",
        "assessed": "tb_assessed",
        "nbeo": "tb_nbeo",
        "asco": "tb_asco",
        "uhco": "tb_uhco"
    }
    return {key: pd.read_excel(FILE, sheet_name=sheet) for key, sheet in sheets.items()}

@st.cache_data
def build_hierarchy(df, max_levels=5, category=None):
    if category:
        df = df[df["category"] == category]

    df = df.dropna(subset=["code"])
    code_map = df.drop_duplicates(subset="code").set_index("code").to_dict("index")
    children = {code: [] for code in df["code"]}
    for _, row in df.iterrows():
        parent = row.get("parent_code")
        if pd.notna(parent):
            children.setdefault(parent, []).append(row["code"])

    leaves = set(code for code, row in code_map.items()
                 if row.get("is_leaf", 0) == 1 or not children.get(code))

    paths = []
    def dfs(code, path):
        path = path + [code]
        if code in leaves or len(path) >= max_levels:
            paths.append(path)
        else:
            for child in children.get(code, []):
                dfs(child, path)

    roots = df[df["parent_code"].isna()]["code"]
    for root in roots:
        dfs(root, [])

    records = []
    for path in paths:
        record = {}
        for i, code in enumerate(path):
            record[f"Level{i+1}_Code"] = code
            record[f"Level{i+1}_Title"] = code_map[code].get("title", "")
        record["Leaf_Code"] = path[-1]
        record["Leaf_Title"] = code_map[path[-1]].get("title", "")
        records.append(record)
    return pd.DataFrame(records)

def hierarchy_select(label, flat_df):
    levels = [col for col in flat_df.columns if col.startswith("Level") and col.endswith("_Title")]
    selections = []
    filtered_df = flat_df.copy()
    for level in levels:
        options = sorted(filtered_df[level].dropna().unique().tolist())
        if not options:
            break
        selection = st.selectbox(f"{label} - {level}", options)
        selections.append(selection)
        filtered_df = filtered_df[filtered_df[level] == selection]
    if not filtered_df.empty:
        row = filtered_df.iloc[0]
        return {
            "code": row["Leaf_Code"],
            "title": row["Leaf_Title"],
            "combined": f"{row['Leaf_Code']} – {row['Leaf_Title']}"
        }
    return {"code": "", "title": "", "combined": ""}

# ==========================================
# Load Reference Data
# ==========================================
refs = load_reference_data()
nbeo_conditions = build_hierarchy(refs["nbeo"], category="Condition")
nbeo_disciplines = build_hierarchy(refs["nbeo"], category="Discipline")
asco = build_hierarchy(refs["asco"])
uhco = build_hierarchy(refs["uhco"])

# ==========================================
# Form UI
# ==========================================
st.markdown("<div class='title-style'>Learning Objective Mapping Form</div>", unsafe_allow_html=True)

with st.container():
    st.markdown("<div class='section-box'>", unsafe_allow_html=True)
    st.subheader("Course Information")
    year = st.selectbox("Year", sorted(refs["courses"]["year"].dropna().unique()))
    semester = st.selectbox("Semester", sorted(refs["courses"].query("year == @year")["semester"].dropna().unique()))
    course_options = refs["courses"].query("year == @year and semester == @semester")
    course_name = st.selectbox("Course Name", sorted(course_options["description"].dropna().unique()))
    type_options = course_options.query("description == @course_name")["lecture_or_lab"].dropna().unique()
    course_type = st.selectbox("Type", type_options if len(type_options) else ["(not specified)"])
    lecture_name = st.text_input("Lecture Name")
    st.markdown("</div>", unsafe_allow_html=True)

with st.container():
    st.markdown("<div class='section-box'>", unsafe_allow_html=True)
    st.subheader("Learning Objective Details")
    learning_objective = st.text_area("Learning Objective")
    # Bloom Level dropdown (show description, save level)
    bloom_df = refs["bloomlevel"].dropna(subset=["description", "level"])
    desc_to_level_map = dict(zip(bloom_df["description"], bloom_df["level"]))
    desc_options = bloom_df["description"].tolist()
    selected_desc = st.selectbox("Bloom Level", desc_options)
    bloom_level = desc_to_level_map.get(selected_desc, "")
    activity_df = refs["activity"].dropna(subset=["description"]) 
    activity_options = activity_df["description"].tolist() 
    activity = st.selectbox("Activity", activity_options)
    method = st.selectbox("Assessment Method", refs["methods"].iloc[:, -1].dropna().unique().tolist())
    difficulty = st.selectbox("Difficulty", refs["difficulty"].iloc[:, -1].dropna().unique().tolist())
    is_assessed = st.selectbox("Is Assessed?", refs["assessed"].iloc[:, -1].dropna().unique().tolist())
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# Conditional Standards + Questions
# ==========================================
questions = []
acoe_standard = ""
nbeo_cond_result = {"code": "", "title": "", "combined": ""}
nbeo_disc_result = {"code": "", "title": "", "combined": ""}
asco_result = {"code": "", "title": "", "combined": ""}
uhco_result = {"code": "", "title": "", "combined": ""}
justification = ""

if is_assessed.lower() == "yes":
    with st.container():
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        st.subheader("Standards Mapping (Hierarchical)")
        nbeo_cond_result = hierarchy_select("NBEO Condition", nbeo_conditions)
        nbeo_disc_result = hierarchy_select("NBEO Discipline", nbeo_disciplines)
        asco_result = hierarchy_select("ASCO", asco)
        uhco_result = hierarchy_select("UHCO", uhco)
        acoe_standard = st.selectbox("ACOE Standard", [
            "2.12 By the time of graduation, students must be able to demonstrate basic life support skills for emergencies encountered in independent optometric practice.",
            "2.13 By the time of graduation, students must be able to identify, record, and analyze pertinent history and problems presented by the patient.",
            "2.14 By the time of graduation, students must be able to examine and evaluate the patient to arrive at an appropriate diagnosis.",
            "2.15 By the time of graduation, students must be able to formulate a rational treatment and management plan and understand the implications of various treatment and management options.",
            "2.16 By the time of graduation, students must be able to provide relevant patient education and counseling.",
            "2.17 By the time of graduation, students must be able to use the knowledge of optometry’s role and the roles of other health professions to appropriately assess and address the health care needs of individual patients and the public health aspects related to the populations being served.",
            "2.18 By the time of graduation, students must be able to apply knowledge of interprofessional collaborative care, ethics, and medico-legal aspects for the delivery of optometric care.",
            "2.19 By the time of graduation, students must be able to demonstrate an understanding of research principles and conduct to critically assess the literature.",
            "2.2 By the time of graduation, students must be able to demonstrate effective and culturally sensitive communications, both oral and written, with other professionals and patients.",
            "2.21 By the time of graduation, students must be able to demonstrate an understanding of the basic principles and philosophy of optometric practice management."
        ])
        st.markdown("</div>", unsafe_allow_html=True)

    questions_text = st.text_area("Enter Questions (one per line)")
    questions = [q.strip() for q in questions_text.split("\n") if q.strip()]

else:
    justification = st.text_area("Justification for not assessing this Learning Objective")

# ==========================================
# Save Section
# ==========================================
if st.button("Save this Learning Objective"):
    if not learning_objective:
        st.error("Learning Objective is required.")
    elif is_assessed.lower() == "yes" and not questions:
        st.error("At least one question is required when LO is assessed.")
    else:
        new_rows = []
        target_sheet = "tblLO_Mapping"

        if is_assessed.lower() == "yes":
            for question in questions:
                new_rows.append({
                    "Year": year,
                    "Semester": semester,
                    "Type": course_type,
                    "CourseName": course_name,
                    "Lecture_Name": lecture_name,
                    "LearningObjective": learning_objective,
                    "BloomLevel": bloom_level,
                    "Activity": activity,
                    "AssessmentMethod": method,
                    "Difficulty": difficulty,
                    "IsAssessed": is_assessed,
                    "NBEO_Condition_Code": nbeo_cond_result["code"],
                    "NBEO_Condition_Title": nbeo_cond_result["title"],
                    "NBEO_Condition": nbeo_cond_result["combined"],
                    "NBEO_Discipline_Code": nbeo_disc_result["code"],
                    "NBEO_Discipline_Title": nbeo_disc_result["title"],
                    "NBEO_Discipline": nbeo_disc_result["combined"],
                    "ASCO_Standard_Code": asco_result["code"],
                    "ASCO_Standard_Title": asco_result["title"],
                    "ASCO_Standard": asco_result["combined"],
                    "UHCO_Standard_Code": uhco_result["code"],
                    "UHCO_Standard_Title": uhco_result["title"],
                    "UHCO_Standard": uhco_result["combined"],
                    "ACOE_Standard": acoe_standard,
                    "Questions": question
                })
        else:
            new_rows.append({
                "Year": year,
                "Semester": semester,
                "Type": course_type,
                "CourseName": course_name,
                "Lecture_Name": lecture_name,
                "LearningObjective": learning_objective,
                "BloomLevel": bloom_level,
                "Activity": activity,
                "AssessmentMethod": method,
                "Difficulty": difficulty,
                "IsAssessed": is_assessed,
                "ACOE_Standard": "",
                "Questions": justification
            })

        try:
            book = load_workbook(FILE)
            if target_sheet in book.sheetnames:
                existing = pd.read_excel(FILE, sheet_name=target_sheet)
                combined = pd.concat([existing, pd.DataFrame(new_rows)], ignore_index=True)
            else:
                combined = pd.DataFrame(new_rows)
        except Exception:
            combined = pd.DataFrame(new_rows)

        with pd.ExcelWriter(FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            combined.to_excel(writer, sheet_name=target_sheet, index=False)

        st.success(f"Saved {len(new_rows)} row(s) to {target_sheet}.")

# ==========================================
# Display Existing Mappings
# ==========================================
try:
    book = load_workbook(FILE)
    if "tblLO_Mapping" in book.sheetnames:
        st.markdown("### Saved Learning Objectives")
        st.dataframe(pd.read_excel(FILE, sheet_name="tblLO_Mapping"))
except Exception:
    pass

# ==========================================
# Learning Objective Visual Dashboard - Summary
# ==========================================
import matplotlib.pyplot as plt

st.markdown("---")
with st.expander("Learning Objective Visual Dashboard - Summary", expanded=False):
    try:
        df = pd.read_excel(FILE, sheet_name="tblLO_Mapping")

        # --- Mappings ---
        def normalize_bloom(level):
            level = str(level).strip().lower()
            for keyword, label in {
                "remember": "Remember",
                "understand": "Understand",
                "apply": "Apply",
                "analyze": "Analyze",
                "evaluate": "Evaluate",
                "create": "Create"
            }.items():
                if keyword in level:
                    return label
            return "Other"

        def map_teaching(activity):
            act = str(activity).lower()
            if "lecture" in act:
               return "Lecture"
            elif "small group" in act or "discussion" in act:
               return "Discussion"
            elif "case" in act:
               return "Case"
            elif "lab" in act:
               return "Lab"
            elif "simulation" in act:
               return "Simulation"
            elif "tbl" in act:
               return "TBL"
            elif "flipped" in act:
               return "Flipped"
            return "Other"

        def map_assessment(method):
            m = str(method).lower()
            if "mcq" in m:
                return "MCQ"
            elif "osce" in m or "practic" in m:
                return "OSCE"
            elif "essay" in m or "long" in m:
                return "Essay"
            elif "oral" in m or "viva" in m:
                return "Oral"
            elif "project" in m:
                return "Project"
            elif "quiz" in m:
                return "Quiz"
            return "Other"

        def get_alignment(activity, method):
            a = str(activity).strip()
            m = str(method).strip()
            if a and m:
                return "Aligned"
            elif a:
                return "Taught_only"
            elif m:
                return "Tested_only"
            return "Ignored"

        df["BloomCategory"] = df["BloomLevel"].apply(normalize_bloom)
        df["TeachingMethodCategory"] = df["Activity"].apply(map_teaching)
        df["AssessmentCategory"] = df["AssessmentMethod"].apply(map_assessment)
        df["AlignmentStatus"] = df.apply(lambda row: get_alignment(row["Activity"], row["AssessmentMethod"]), axis=1)
        df = df[df["AlignmentStatus"] != "Ignored"]

        # --- Filters ---
        st.subheader("Filter by Year and Semester")
        year_options = sorted(df["Year"].dropna().unique())
        selected_year = st.selectbox("Select Year", year_options)
        semester_options = sorted(df[df["Year"] == selected_year]["Semester"].dropna().unique())
        selected_semester = st.selectbox("Select Semester", semester_options)
        filtered_df = df[(df["Year"] == selected_year) & (df["Semester"] == selected_semester)]

        # --- Charts ---
        st.markdown("### Bloom Level Distribution")
        st.bar_chart(filtered_df["BloomCategory"].value_counts())

        st.markdown("### Teaching Method Distribution")
        fig1, ax1 = plt.subplots()
        filtered_df["TeachingMethodCategory"].value_counts().plot.pie(autopct='%1.1f%%', ax=ax1)
        ax1.set_ylabel("")
        st.pyplot(fig1)

        st.markdown("### Assessment Method Distribution")
        fig2, ax2 = plt.subplots()
        filtered_df["AssessmentCategory"].value_counts().plot.pie(autopct='%1.1f%%', ax=ax2)
        ax2.set_ylabel("")
        st.pyplot(fig2)

        st.markdown("### Teach–Test Alignment")
        fig3, ax3 = plt.subplots()
        filtered_df["AlignmentStatus"].value_counts().plot.pie(autopct='%1.1f%%', ax=ax3)
        ax3.set_ylabel("")
        st.pyplot(fig3)

    except Exception as e:
        st.error(f"Error generating dashboard: {e}")
