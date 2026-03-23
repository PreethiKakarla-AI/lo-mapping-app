# Learning Objective Mapping App

**Course Project — University of Houston College of Optometry (UHCO)**
**Developer:** Preethi Kakarla
**GitHub Repository:** https://github.com/PreethiKakarla-AI/lo-mapping-app

---

## Project Overview

This is a **Streamlit web application** that provides a structured digital form for mapping Learning Objectives (LOs) to curriculum standards used at UHCO. It allows faculty to record, tag, and visualize how each learning objective aligns with:

- **Bloom's Taxonomy** levels (Remember → Create)
- **NBEO** (National Board of Examiners in Optometry) standards — Conditions and Disciplines
- **ASCO** (Association of Schools and Colleges of Optometry) standards
- **UHCO** internal curriculum standards
- **ACOE** (Accreditation Council on Optometric Education) standards

All submitted data is saved directly to this GitHub repository as a CSV file (`tblLO_Mapping.csv`), making it version-controlled and accessible to the whole team.

---

## Repository File Structure

```
lo-mapping-app/
│
├── app.py                                        # Main Streamlit application (all UI + logic)
├── requirements.txt                              # Python dependencies
├── LOreferenceData_final_formfeedversion2.xlsx   # Reference data (courses, Bloom levels, standards hierarchy)
├── tblLO_Mapping.csv                             # Output file — all submitted LO mappings (auto-updated)
├── Procfile                                      # Server startup command (for cloud deployment)
└── README.md                                     # This file
```

---

## Application Features

### 1. Course Information
- Select Year, Semester, Course Name, and Course Type (Lecture/Lab) from dropdown menus
- All course data is pulled from the Excel reference file
- Free-text field for Lecture Name

### 2. Learning Objective Details
- Text area for entering the full Learning Objective
- Bloom Level selection (with full descriptions)
- Teaching Method (macro activity) dropdown — 24 options
- Micro-Activity dropdown — 18 options (what students actually do)
- Summative Assessment dropdown — 20 options (how it is graded)
- Difficulty level and "Is Assessed?" flag

### 3. Standards Mapping (shown only when LO is assessed)
- **NBEO Condition** — hierarchical drill-down selector
- **NBEO Discipline** — hierarchical drill-down selector
- **ASCO Standard** — hierarchical drill-down selector
- **UHCO Standard** — hierarchical drill-down selector
- **ACOE Standard** — dropdown with 11 graduate competency standards
- **Exam Questions** — add one or multiple exam questions linked to this LO

### 4. Data Persistence via GitHub API
- On clicking "Save", the app writes a new row to `tblLO_Mapping.csv` in this GitHub repo
- Uses a GitHub Personal Access Token stored in Streamlit Secrets
- The CSV file grows incrementally — each save appends to the existing data

### 5. Visual Dashboard
- Expandable dashboard showing charts of saved LO data
- **Chart 1:** Bloom Level coverage (bar chart)
- **Chart 2:** Bloom Levels by Teaching Method (stacked)
- **Chart 3:** Bloom Levels by Micro-Activity (top 8)
- **Chart 4:** Bloom Levels by Summative Assessment (top 8)
- Filter by Year and Semester

---

## Technology Stack

| Component | Technology |
|---|---|
| Frontend + Backend | Python / Streamlit |
| Data Storage | GitHub API (CSV file in this repo) |
| Reference Data | Excel (.xlsx) via openpyxl |
| Charts | Streamlit built-in + Matplotlib |
| Secrets Management | Streamlit Secrets (`st.secrets`) |
| Dependencies | streamlit, pandas, openpyxl, matplotlib, requests |

---

## How to Run Locally

### Prerequisites
- Python 3.9 or higher
- Git

### Steps

```bash
# 1. Clone the repository
git clone https://github.com/PreethiKakarla-AI/lo-mapping-app.git
cd lo-mapping-app

# 2. Install dependencies
pip install -r requirements.txt

# 3. Set up Streamlit secrets
mkdir -p .streamlit
cat > .streamlit/secrets.toml << EOF
GITHUB_TOKEN = "your_github_personal_access_token"
GITHUB_REPO = "PreethiKakarla-AI/lo-mapping-app"
GITHUB_BRANCH = "main"
GITHUB_FILE_PATH = "tblLO_Mapping.csv"
EOF

# 4. Run the app
streamlit run app.py
```

The app will open at `http://localhost:8501`

---

## How to Deploy on Streamlit Community Cloud (Free Hosting)

Streamlit Community Cloud is the recommended way to host this app for free.

### Step-by-Step Deployment

1. Go to **https://share.streamlit.io** and sign in with your GitHub account

2. Click **"New app"**

3. Fill in:
   - **Repository:** `PreethiKakarla-AI/lo-mapping-app`
   - **Branch:** `main`
   - **Main file path:** `app.py`

4. Click **"Advanced settings"** and add these secrets:
   ```toml
   GITHUB_TOKEN = "your_github_personal_access_token"
   GITHUB_REPO = "PreethiKakarla-AI/lo-mapping-app"
   GITHUB_BRANCH = "main"
   GITHUB_FILE_PATH = "tblLO_Mapping.csv"
   ```

5. Click **"Deploy"** — the app will be live in about 2 minutes

6. You will receive a public URL like: `https://yourname-lo-mapping-app.streamlit.app`

### What is a GitHub Personal Access Token?
The app needs a GitHub token to write saved LO data back to the CSV file in this repo.

To create one:
1. Go to GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. Click "Generate new token (classic)"
3. Give it a name, set expiration, and check the **`repo`** scope
4. Copy the token and paste it into Streamlit secrets as `GITHUB_TOKEN`

---

## Data Output: tblLO_Mapping.csv

Every time a faculty member submits a Learning Objective, a new row is appended to `tblLO_Mapping.csv` with these columns:

| Column | Description |
|---|---|
| Year | Academic year |
| Semester | Fall / Spring / Summer |
| Type | Lecture or Lab |
| CourseName | Full course name |
| Lecture_Name | Lecture title |
| LearningObjective | The full LO text |
| BloomLevel | Bloom's taxonomy level (numeric) |
| Activity | Teaching method |
| MicroActivity | Specific student activity |
| AssessmentMethod | Summative assessment type |
| Difficulty | Difficulty rating |
| IsAssessed | Yes / No |
| NBEO_Condition_Code | NBEO condition code |
| NBEO_Condition_Title | NBEO condition title |
| NBEO_Condition | Combined code + title |
| NBEO_Discipline_Code | NBEO discipline code |
| NBEO_Discipline_Title | NBEO discipline title |
| NBEO_Discipline | Combined code + title |
| ASCO_Standard_Code | ASCO standard code |
| ASCO_Standard_Title | ASCO standard title |
| ASCO_Standard | Combined code + title |
| UHCO_Standard_Code | UHCO standard code |
| UHCO_Standard_Title | UHCO standard title |
| UHCO_Standard | Combined code + title |
| ACOE_Standard | Full ACOE standard text |
| Questions | Exam question(s) linked to this LO |

---

## Reference Data: LOreferenceData_final_formfeedversion2.xlsx

This Excel file is the backbone of the app's dropdown menus. It contains these sheets:

| Sheet | Contents |
|---|---|
| tb_courses | All courses with year, semester, type |
| tb_bloomlevel | Bloom's taxonomy levels + descriptions |
| tb_activity | (reference) |
| tb_methods | (reference) |
| tb_difficulty | Difficulty options |
| tb_assessed | Is-assessed options |
| tb_nbeo | NBEO standards hierarchy (conditions + disciplines) |
| tb_asco | ASCO standards hierarchy |
| tb_uhco | UHCO standards hierarchy |

---

## Architecture Diagram

```
[Faculty Browser]
      |
      | HTTPS
      v
[Streamlit App - app.py]
      |              |
      |              | Read reference data
      |              v
      |     [LOreferenceData.xlsx]
      |
      | GitHub API (read/write)
      v
[GitHub Repo - tblLO_Mapping.csv]
      |
      | (version controlled, downloadable)
      v
[Professor / Admin downloads CSV for analysis]
```

---

## Known Limitations & Future Improvements

- **No authentication:** Any person with the app URL can submit entries. Future versions could add a login system.
- **No edit/delete:** Submitted LOs cannot be edited from the app — must edit the CSV directly in GitHub.
- **Single file storage:** All data is in one CSV. For large-scale use, a database (e.g., PostgreSQL, Supabase) would be better.
- **GitHub token expiry:** The GitHub Personal Access Token must be manually renewed when it expires.

---

## Contact

**Developer:** Preethi Kakarla
**Email:** pkakarla@CougarNet.UH.EDU
**University:** University of Houston College of Optometry (UHCO)
