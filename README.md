# Engineering Cost Intelligence

A Python + Streamlit application designed to assist engineers in **BOQ analysis, budget planning, historical benchmarking, and project risk evaluation**.

This tool helps engineers quickly understand project costs, identify cost drivers, and evaluate risks using historical project data.

---

## Key Features

### BOQ Analysis
- Upload BOQ files
- Automatic cost breakdown
- Category-level analysis

### Budget Allocation Engine
- Material, labor, transport, contingency allocation
- Profit margin evaluation
- Budget validation rules

### Historical Cost Intelligence
- Compare current project against historical project data
- Detect deviations from historical averages

### Engineering Risk Intelligence
- Project risk score
- Cost concentration detection
- Execution risk insights

### Project Comparison
- Compare multiple saved projects
- Allocation comparison
- BOQ variance analysis

### Reporting
- Excel report generation
- PDF report generation
- Historical comparison Excel export

### Multi-BOQ Upload (Phase 2)
- Upload multiple BOQs
- Merge them as a single project OR analyze separately
- BOQ source tracking

---

## Technology Stack

- Python
- Streamlit
- Pandas
- Matplotlib
- OpenPyXL
- SQLite

---

## How to Run

Install dependencies:

```
pip install -r requirements.txt
```

Run the application:

```
streamlit run streamlit_app.py
```

---

## Project Evolution

### Phase 1
Initial BOQ analysis and budget allocation system.

### Phase 2
Engineering intelligence features added:
- historical benchmarking
- risk scoring
- project comparison
- multi-BOQ workflow
- advanced reporting

---

## Phase 3 (Planned)

- Similar project detection
- Recommended budget ranges
- Estimator assistant intelligence

---

## What This Tool Solves

Engineers often analyze BOQs manually in spreadsheets.

This system provides:
- structured cost analysis
- historical comparison
- risk insights
- automated reporting

to support better **engineering decision making**.

