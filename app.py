import streamlit as st
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
import warnings
import matplotlib.pyplot as plt

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
st.markdown(
    """
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
    """,
    unsafe_allow_html=True,
)

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
        "uhco": "tb_uhco",
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

    leaves = set(
        code
        for code, row in code_map.items()
        if row.get("is_leaf", 0) == 1 or not children.get(code)
    )

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
    levels = [
        col
        for col in flat_df.columns
        if col.startswith("Level") and col.endswith("_Title")
    ]
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
            "combined": f"{row['Leaf_Code']} – {row['Leaf_Title']}",
        }

    return {"code": "", "title": "", "combined": ""}


# ==========================================
# Load Reference Data
# ==========================================
if st.button("Reload reference tables from Excel"):
    load_reference_data.clear()  # clear cached copy

refs = load_reference_data()
nbeo_conditions = build_hierarchy(refs["nbeo"], category="Condition")
nbeo_disciplines = build_hierarchy(refs["nbeo"], category="Discipline")
asco = build_hierarchy(refs["asco"])
uhco = build_hierarchy(refs["uhco"])

# ==========================================
# Form UI
# ==========================================
st.markdown(
    "<div class='title-style'>Learning Objective Mapping Form</div>",
    unsafe_allow_html=True,
)

# ------------------------------------------
# Course Information
# ------------------------------------------
with st.container():
    st.markdown("<div class='section-box'>", unsafe_allow_html=True)
    st.subheader("Course Information")

    year = st.selectbox("Year", sorted(refs["courses"]["year"].dropna().unique()))

    semester = st.selectbox(
        "Semester",
        sorted(
            refs["courses"]
            .query("year == @year")["semester"]
            .dropna()
            .unique()
        ),
    )

    course_options = refs["courses"].query(
        "year == @year and semester == @semester"
    )
    course_name = st.selectbox(
        "Course Name",
        sorted(course_options["description"].dropna().unique()),
    )

    type_options = (
        course_options.query("description == @course_name")["lecture_or_lab"]
        .dropna()
        .unique()
    )
    course_type = st.selectbox(
        "Type", type_options if len(type_options) else ["(not specified)"]
    )

    lecture_name = st.text_input("Lecture Name")

    st.markdown("</div>", unsafe_allow_html=True)

# ------------------------------------------
# Learning Objective Details
# ------------------------------------------
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

    # Teaching Method / Micro-Activity / Summative Assessment
    teaching_methods = [
        "Direct instruction",
        "Mini-lecture",
        "Demonstration",
        "Video micro-lecture",
        "Concept mapping",
        "Example vs. Non-example",
        "Guided comparison",
        "Visual explanation",
        "Worked examples",
        "Skill demonstration",
        "Simulation (SimLab)",
        "Problem sets",
        "Case-Based Learning (CBL)",
        "Problem-Based Learning (PBL)",
        "Team-Based Learning (TBL)",
        "Diagnostic simulation",
        "Debate",
        "Ranking rounds",
        "Critique rounds",
        "Treatment reasoning",
        "Protocol design",
        "Project build",
        "Case plan development",
        "Clinical pathway creation",
    ]

    micro_activities = [
        "Label diagrams",
        "Match terms",
        "Flash-classify",
        "Explain-back",
        "Compare summaries",
        "Describe anatomy/phys differences",
        "Follow algorithm",
        "Execute skill step",
        "Perform skill (e.g., retinoscopy)",
        "Pattern recognition (OCT, VF, cornea)",
        "Error spotting",
        "Annotate findings",
        "Defend best treatment",
        "Rank options",
        "Critique management plan",
        "Build treatment plan",
        "Build protocol",
        "Propose diagnostic strategy",
    ]

    summative_assessments = [
        "Knowledge MCQs",
        "Definition SAQ",
        "Labeling quiz",
        "Explanation MCQ",
        "Short answer reasoning",
        "Map completion scoring",
        "Applied MCQ",
        "Numeric SAQ",
        "OSCE skill station",
        "Case-based MCQ",
        "Image interpretation MCQ",
        "OSCE diagnostic station",
        "Treatment choice SAQ",
        "OSCE case justification",
        "Chart audit Mini-CEX",
        "Workplace-Based Assessment (WBA)",
        "Project scoring",
        "Protocol proposal rubric",
        "Presentation assessment",
        "Scholarly project evaluation",
    ]

    teaching_method = st.selectbox("Teaching Method (macro activity)", teaching_methods)
    micro_activity = st.selectbox("Micro-Activity", micro_activities)
    summative_assessment = st.selectbox(
        "Summative Assessment (assessment method)", summative_assessments
    )

    difficulty = st.selectbox(
        "Difficulty",
        refs["difficulty"].iloc[:, -1].dropna().unique().tolist(),
    )

    is_assessed = st.selectbox(
        "Is Assessed?",
        refs["assessed"].iloc[:, -1].dropna().unique().tolist(),
    )

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

assessed_flag = str(is_assessed).strip().lower() == "yes"

if assessed_flag:
    # Standards mapping
    with st.container():
        st.markdown("<div class='section-box'>", unsafe_allow_html=True)
        st.subheader("Standards Mapping (Hierarchical)")

        nbeo_cond_result = hierarchy_select("NBEO Condition", nbeo_conditions)
        nbeo_disc_result = hierarchy_select("NBEO Discipline", nbeo_disciplines)
        asco_result = hierarchy_select("ASCO", asco)
        uhco_result = hierarchy_select("UHCO", uhco)

        acoe_standard = st.selectbox(
            "ACOE Standard",
            [
                "2.8  Students must attain the defined set of clinical competencies established by the program.",
                "2.12 By graduation, students must be able to recognize and appropriately respond to ocular and systemic emergencies in optometric practice.",
                "2.13 By graduation, students must be able to identify and analyze relevant history and presenting problems for each patient.",
                "2.14 By graduation, students must be able to perform appropriate examinations and evaluate findings to reach an accurate diagnosis.",
                "2.15 By graduation, students must be able to formulate and justify management plans, understanding risks and benefits of options.",
                "2.16 By graduation, students must be able to provide clear, relevant patient education and counseling.",
                "2.17 By graduation, students must be able to consider and address public and population health factors in patient care.",
                "2.18 By graduation, students must be able to apply ethical, legal, and medico-legal principles in delivering optometric care.",
                "2.19 By graduation, students must be able to use research principles to critically appraise scientific and clinical literature.",
                "2.20 By graduation, students must be able to communicate effectively, orally and in writing, with patients and other professionals.",
                "2.21 By graduation, students must be able to demonstrate understanding of optometric practice management principles.",
            ],
        )

        st.markdown("</div>", unsafe_allow_html=True)

    # Dynamic questions section
    st.markdown("<div class='section-box'>", unsafe_allow_html=True)
    st.subheader("Exam Question(s)")

    # Initialise question count in session state
    if "question_count" not in st.session_state:
        st.session_state.question_count = 1

    # Button to add another question input
    if st.button("Add another question"):
        st.session_state.question_count += 1

    # Render one text_input per question
    question_texts = []
    for i in range(st.session_state.question_count):
        q = st.text_input(
            f"Question {i + 1}",
            key=f"question_{i + 1}",
        )
        question_texts.append(q)

    st.markdown("</div>", unsafe_allow_html=True)

    # Build list of non-empty questions
    questions = [q.strip() for q in question_texts if q.strip()]
else:
    justification = st.text_area(
        "Justification for not assessing this Learning Objective"
    )

# ==========================================
# Save Section
# ==========================================
if st.button("Save this Learning Objective"):
    if not learning_objective:
        st.error("Learning Objective is required.")
    elif assessed_flag and not questions:
        st.error("At least one question is required when LO is assessed.")
    else:
        new_rows = []
        target_sheet = "tblLO_Mapping"

        if assessed_flag:
            # one row per question
            for question in questions:
                new_rows.append(
                    {
                        "Year": year,
                        "Semester": semester,
                        "Type": course_type,
                        "CourseName": course_name,
                        "Lecture_Name": lecture_name,
                        "LearningObjective": learning_objective,
                        "BloomLevel": bloom_level,
                        "Activity": teaching_method,
                        "MicroActivity": micro_activity,
                        "AssessmentMethod": summative_assessment,
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
                        "Questions": question,
                    }
                )
        else:
            new_rows.append(
                {
                    "Year": year,
                    "Semester": semester,
                    "Type": course_type,
                    "CourseName": course_name,
                    "Lecture_Name": lecture_name,
                    "LearningObjective": learning_objective,
                    "BloomLevel": bloom_level,
                    "Activity": teaching_method,
                    "MicroActivity": micro_activity,
                    "AssessmentMethod": summative_assessment,
                    "Difficulty": difficulty,
                    "IsAssessed": is_assessed,
                    "ACOE_Standard": "",
                    "Questions": justification,
                }
            )

        # Build combined dataframe and write safely
        try:
            if FILE.exists():
                try:
                    existing = pd.read_excel(FILE, sheet_name=target_sheet)
                    combined = pd.concat(
                        [existing, pd.DataFrame(new_rows)],
                        ignore_index=True,
                    )
                    write_mode = "a"
                except ValueError:
                    # File exists but sheet does not
                    combined = pd.DataFrame(new_rows)
                    write_mode = "a"
            else:
                combined = pd.DataFrame(new_rows)
                write_mode = "w"
        except Exception:
            combined = pd.DataFrame(new_rows)
            write_mode = "w"

        try:
            if write_mode == "a" and FILE.exists():
                with pd.ExcelWriter(
                    FILE,
                    engine="openpyxl",
                    mode="a",
                    if_sheet_exists="replace",
                ) as writer:
                    combined.to_excel(
                        writer, sheet_name=target_sheet, index=False
                    )
            else:
                with pd.ExcelWriter(
                    FILE,
                    engine="openpyxl",
                    mode="w",
                ) as writer:
                    combined.to_excel(
                        writer, sheet_name=target_sheet, index=False
                    )

            st.success(f"Saved {len(new_rows)} row(s) to {target_sheet}.")
        except Exception as e:
            st.error(f"Error while saving: {e}")

# ==========================================
# Display Existing Mappings
# ==========================================
if FILE.exists():
    try:
        existing_df = pd.read_excel(FILE, sheet_name="tblLO_Mapping")
        st.markdown("### Saved Learning Objectives")
        st.dataframe(existing_df)
    except Exception:
        pass

# ==========================================
# Learning Objective Visual Dashboard - Summary
# ==========================================
st.markdown("---")
with st.expander("Learning Objective Visual Dashboard - Summary", expanded=False):
    try:
        if not FILE.exists():
            st.info(
                "No saved data yet. Save at least one Learning Objective to view the dashboard."
            )
        else:
            df = pd.read_excel(FILE, sheet_name="tblLO_Mapping")

            # --- Mapping helpers ---
            def normalize_bloom(level):
                level = str(level).strip().lower()
                mapping = {
                    "remember": "Remember",
                    "understand": "Understand",
                    "apply": "Apply",
                    "analyze": "Analyze",
                    "evaluate": "Evaluate",
                    "create": "Create",
                }
                for keyword, label in mapping.items():
                    if keyword in level:
                        return label
                return "Other"

            def get_alignment(micro_value, summative_value):
                a = str(micro_value).strip()
                m = str(summative_value).strip()
                if a and m:
                    return "Aligned"
                elif a:
                    return "Taught_only"
                elif m:
                    return "Tested_only"
                return "Ignored"

            # Derived categories based on new fields
            df["BloomCategory"] = df["BloomLevel"].apply(normalize_bloom)
            df["TeachingMethodCategory"] = df["Activity"].fillna("Unspecified")
            df["MicroActivityCategory"] = df.get(
                "MicroActivity", pd.Series(index=df.index)
            ).fillna("Unspecified")
            df["SummativeCategory"] = df["AssessmentMethod"].fillna("Unspecified")
            df["AlignmentStatus"] = df.apply(
                lambda row: get_alignment(
                    row.get("MicroActivity", ""), row.get("AssessmentMethod", "")
                ),
                axis=1,
            )

            df = df[df["AlignmentStatus"] != "Ignored"]

            if df.empty:
                st.info(
                    "No mapped teaching and assessment data to display yet."
                )
            else:
                # --- Filters ---
                st.subheader("Filter by Year and Semester")
                year_options = sorted(df["Year"].dropna().unique())
                selected_year = st.selectbox("Select Year", year_options)

                semester_options = sorted(
                    df[df["Year"] == selected_year]["Semester"]
                    .dropna()
                    .unique()
                )
                selected_semester = st.selectbox(
                    "Select Semester", semester_options
                )

                filtered_df = df[
                    (df["Year"] == selected_year)
                    & (df["Semester"] == selected_semester)
                ]

                if filtered_df.empty:
                    st.info(
                        "No records for the selected Year and Semester."
                    )
                else:
                    # --- Charts ---
                    st.markdown("### Bloom Level Distribution")
                    st.bar_chart(filtered_df["BloomCategory"].value_counts())

                    st.markdown("### Teaching Method (macro) Distribution")
                    st.bar_chart(filtered_df["TeachingMethodCategory"].value_counts())

                    st.markdown("### Micro-Activity Distribution")
                    st.bar_chart(filtered_df["MicroActivityCategory"].value_counts())

                    st.markdown("### Summative Assessment Distribution")
                    st.bar_chart(filtered_df["SummativeCategory"].value_counts())

                    st.markdown("### Micro vs Summative Alignment")
                    fig3, ax3 = plt.subplots()
                    filtered_df["AlignmentStatus"].value_counts().plot(
                        kind="pie", autopct="%1.1f%%", ax=ax3
                    )
                    ax3.set_ylabel("")
                    st.pyplot(fig3)

    except Exception as e:
        st.error(f"Error generating dashboard: {e}")
