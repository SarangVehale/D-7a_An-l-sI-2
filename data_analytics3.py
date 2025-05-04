import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
import seaborn as sns
from jinja2 import Environment, FileSystemLoader
import os
import base64

# -----------------------------
# CONFIGURATION
# -----------------------------
EXCEL_FILE = "Awareness and Effectiveness of Workplace Harassment Policies in the Private Sector  (Responses).xlsx"
DB_FILE = "harassment_survey.db"
CHART_DIR = "charts"
REPORT_FILE = "enhanced_harassment_policy_report.html"
EMBED_FILE_NAME = "survey_data.xlsx"

# -----------------------------
# LOAD AND CLEAN DATA
# -----------------------------
df = pd.read_excel(EXCEL_FILE)
df.columns = [col.strip().replace(" ", "_").replace("?", "").replace(",", "").replace("(", "").replace(")", "").replace("/", "_") for col in df.columns]

# Save cleaned Excel to embed
df.to_excel(EMBED_FILE_NAME, index=False)

# Convert Excel file to base64 for embedding
with open(EMBED_FILE_NAME, "rb") as f:
    encoded_excel = base64.b64encode(f.read()).decode("utf-8")

# Create charts directory
os.makedirs(CHART_DIR, exist_ok=True)

# -----------------------------
# SAVE TO SQLITE
# -----------------------------
conn = sqlite3.connect(DB_FILE)
df.to_sql("survey", conn, if_exists="replace", index=False)

# -----------------------------
# QUERIES
# -----------------------------
awareness_df = pd.read_sql_query("""
    SELECT Are_you_aware_of_what_constitutes_workplace_harassment AS Awareness, COUNT(*) AS Count
    FROM survey GROUP BY Awareness;
""", conn)

awareness_reporting_df = pd.read_sql_query("""
    SELECT Are_you_aware_of_what_constitutes_workplace_harassment AS Awareness,
           Do_you_know_whom_to_report_workplace_harassment_incidents_to AS Know_Whom,
           COUNT(*) AS Count
    FROM survey GROUP BY Awareness, Know_Whom;
""", conn)

training_reporting_df = pd.read_sql_query("""
    SELECT Have_you_ever_received_any_formal_training_on_workplace_harassment_policies AS Training,
           Do_you_know_whom_to_report_workplace_harassment_incidents_to AS Know_Whom,
           COUNT(*) AS Count
    FROM survey GROUP BY Training, Know_Whom;
""", conn)

gender_awareness_df = pd.read_sql_query("""
    SELECT What_is_your_Gender AS Gender,
           Are_you_aware_of_what_constitutes_workplace_harassment AS Awareness,
           COUNT(*) AS Count
    FROM survey GROUP BY Gender, Awareness;
""", conn)

conn.close()

# -----------------------------
# VISUALIZATION FUNCTIONS
# -----------------------------
def save_chart(data, title, filename, kind='bar', pivot=False, index=None, columns=None, y='Count', palette='Set2', stacked=False):
    plt.figure(figsize=(8, 5))
    if pivot:
        data = data.pivot(index=index, columns=columns, values=y).fillna(0)
        data.plot(kind=kind, stacked=stacked)
    else:
        sns.barplot(data=data, x=data.columns[0], y=y, palette=palette)
    plt.title(title)
    plt.xticks(rotation=45)
    plt.tight_layout()
    path = os.path.join(CHART_DIR, filename)
    plt.savefig(path)
    plt.close()
    return path

# Save all charts
chart_paths = {
    "awareness": save_chart(awareness_df, "Overall Awareness", "awareness.png"),
    "awareness_vs_reporting": save_chart(awareness_reporting_df, "Awareness vs Reporting Knowledge", "awareness_vs_reporting.png", pivot=True, index='Awareness', columns='Know_Whom', stacked=True),
    "training_vs_reporting": save_chart(training_reporting_df, "Training vs Reporting Knowledge", "training_vs_reporting.png", pivot=True, index='Training', columns='Know_Whom', stacked=True),
    "gender_awareness": save_chart(gender_awareness_df, "Gender-based Awareness", "gender_awareness.png", pivot=True, index='Gender', columns='Awareness', stacked=True),
}

# -----------------------------
# GENERATE INTERPRETATIONS
# -----------------------------
total = awareness_df['Count'].sum()
yes_count = awareness_df[awareness_df['Awareness'] == 'Yes']['Count'].values[0]
yes_pct = yes_count / total * 100

summary = f"""
<strong>Executive Summary:</strong><br>
Out of {total} respondents, <strong>{yes_pct:.1f}%</strong> reported being fully aware of what constitutes workplace harassment. This reflects a broadly educated workforce, but additional analysis suggests critical gaps remain in knowledge of reporting procedures and disparities across gender and training status.
"""

# -----------------------------
# BUILD HTML REPORT
# -----------------------------
html_template = """
<!DOCTYPE html>
<html>
<head>
    <title>Enhanced Workplace Harassment Survey Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; background-color: #f9f9f9; }
        h1, h2 { color: #2c3e50; }
        h2 { margin-top: 40px; }
        img { max-width: 100%; border: 1px solid #ccc; padding: 4px; background: #fff; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; background: white; }
        th, td { padding: 8px; border: 1px solid #ccc; text-align: left; }
        .chart { margin: 30px 0; }
        .section { background: #fff; padding: 20px; border-radius: 6px; box-shadow: 0 0 10px #ccc; margin-bottom: 30px; }
    </style>
</head>
<body>
    <h1>Survey Report: Awareness and Effectiveness of Workplace Harassment Policies</h1>
    
    <div class="section">
        <h2>1. Executive Summary</h2>
        <p>{{ summary|safe }}</p>
    </div>

    <div class="section">
        <h2>2. Overall Awareness</h2>
        <img src="{{ charts.awareness }}" alt="Awareness Chart">
        {{ awareness_df|safe }}
    </div>

    <div class="section">
        <h2>3. Awareness vs Reporting Knowledge</h2>
        <img src="{{ charts.awareness_vs_reporting }}" alt="Awareness vs Reporting Chart">
        {{ awareness_reporting_df|safe }}
    </div>

    <div class="section">
        <h2>4. Formal Training vs Reporting Knowledge</h2>
        <img src="{{ charts.training_vs_reporting }}" alt="Training vs Reporting Chart">
        {{ training_reporting_df|safe }}
    </div>

    <div class="section">
        <h2>5. Gender-based Awareness</h2>
        <img src="{{ charts.gender_awareness }}" alt="Gender Awareness Chart">
        {{ gender_awareness_df|safe }}
    </div>

    <div class="section">
        <h2>6. Original Survey Dataset</h2>
        <p>You can download the original Excel file used in this report below:</p>
        <a download="SurveyData.xlsx" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{{ encoded_excel }}">Download Excel File</a>
    </div>
</body>
</html>
"""

env = Environment()
template = env.from_string(html_template)

html_output = template.render(
    summary=summary,
    awareness_df=awareness_df.to_html(index=False),
    awareness_reporting_df=awareness_reporting_df.to_html(index=False),
    training_reporting_df=training_reporting_df.to_html(index=False),
    gender_awareness_df=gender_awareness_df.to_html(index=False),
    charts=chart_paths,
    encoded_excel=encoded_excel
)

with open(REPORT_FILE, "w", encoding="utf-8") as f:
    f.write(html_output)

print(f"✅ Enhanced report generated: {REPORT_FILE}")

