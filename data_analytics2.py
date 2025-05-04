import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
import seaborn as sns
from jinja2 import Environment, FileSystemLoader
import os

# Set up paths
EXCEL_FILE = "data.xlsx"
DB_FILE = "harassment_survey.db"
CHART_DIR = "charts"
REPORT_FILE = "harassment_policy_report.html"

# Create charts folder
os.makedirs(CHART_DIR, exist_ok=True)

# Load Excel and clean column names
df = pd.read_excel(EXCEL_FILE)
df.columns = [col.strip().replace(" ", "_").replace("?", "").replace(",", "").replace("(", "").replace(")", "").replace("/", "_") for col in df.columns]

# Save to SQLite
conn = sqlite3.connect(DB_FILE)
df.to_sql("survey", conn, if_exists="replace", index=False)

# Run Queries
awareness_df = pd.read_sql_query("""
    SELECT Are_you_aware_of_what_constitutes_workplace_harassment AS Awareness, COUNT(*) AS Count
    FROM survey
    GROUP BY Awareness;
""", conn)

awareness_reporting_df = pd.read_sql_query("""
    SELECT 
        Are_you_aware_of_what_constitutes_workplace_harassment AS Awareness,
        Do_you_know_whom_to_report_workplace_harassment_incidents_to AS Know_Whom_To_Report,
        COUNT(*) AS Count
    FROM survey
    GROUP BY Awareness, Know_Whom_To_Report;
""", conn)

training_reporting_df = pd.read_sql_query("""
    SELECT 
        Have_you_ever_received_any_formal_training_on_workplace_harassment_policies AS Training,
        Do_you_know_whom_to_report_workplace_harassment_incidents_to AS Know_Whom_To_Report,
        COUNT(*) AS Count
    FROM survey
    GROUP BY Training, Know_Whom_To_Report;
""", conn)

gender_awareness_df = pd.read_sql_query("""
    SELECT 
        What_is_your_Gender AS Gender,
        Are_you_aware_of_what_constitutes_workplace_harassment AS Awareness,
        COUNT(*) AS Count
    FROM survey
    GROUP BY Gender, Awareness;
""", conn)

conn.close()

# Plot Charts
def save_chart(data, x, y, title, filename, kind='bar', stacked=False, pivot=False, index=None, columns=None):
    plt.figure(figsize=(8, 5))
    if pivot:
        data = data.pivot(index=index, columns=columns, values=y).fillna(0)
        data.plot(kind=kind, stacked=stacked)
    else:
        sns.barplot(data=data, x=x, y=y, palette='Set2')
    plt.title(title)
    plt.tight_layout()
    path = os.path.join(CHART_DIR, filename)
    plt.savefig(path)
    plt.close()
    return path

# Save chart images
chart_paths = {
    "awareness": save_chart(awareness_df, 'Awareness', 'Count', "Overall Awareness", "awareness.png"),
    "awareness_vs_reporting": save_chart(awareness_reporting_df, None, 'Count', "Awareness vs Reporting Knowledge", "awareness_vs_reporting.png", kind='bar', stacked=True, pivot=True, index='Awareness', columns='Know_Whom_To_Report'),
    "training_vs_reporting": save_chart(training_reporting_df, None, 'Count', "Training vs Reporting Knowledge", "training_vs_reporting.png", kind='bar', stacked=True, pivot=True, index='Training', columns='Know_Whom_To_Report'),
    "gender_awareness": save_chart(gender_awareness_df, None, 'Count', "Gender-based Awareness", "gender_awareness.png", kind='bar', stacked=True, pivot=True, index='Gender', columns='Awareness'),
}

# Prepare data for template
total = awareness_df['Count'].sum()
yes_pct = awareness_df[awareness_df['Awareness'] == 'Yes']['Count'].values[0] / total * 100

template_data = {
    "awareness_pct": f"{yes_pct:.1f}",
    "awareness_data": awareness_df.to_html(index=False),
    "awareness_vs_reporting": awareness_reporting_df.to_html(index=False),
    "training_vs_reporting": training_reporting_df.to_html(index=False),
    "gender_awareness": gender_awareness_df.to_html(index=False),
    "charts": chart_paths
}

# Jinja2 HTML Template Rendering
html_template = """
<!DOCTYPE html>
<html>
<head>
    <title>Workplace Harassment Policy Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1 { color: #333; }
        h2 { margin-top: 40px; }
        img { max-width: 100%; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th, td { padding: 8px; border: 1px solid #ccc; text-align: left; }
    </style>
</head>
<body>
    <h1>Survey Report: Awareness and Effectiveness of Workplace Harassment Policies</h1>

    <h2>1. Awareness of Workplace Harassment</h2>
    <p><strong>{{ awareness_pct }}%</strong> of respondents reported full awareness.</p>
    <img src="{{ charts.awareness }}" alt="Awareness Chart">
    {{ awareness_data|safe }}

    <h2>2. Awareness vs Knowledge of Reporting</h2>
    <img src="{{ charts.awareness_vs_reporting }}" alt="Awareness vs Reporting Chart">
    {{ awareness_vs_reporting|safe }}

    <h2>3. Formal Training vs Reporting Knowledge</h2>
    <img src="{{ charts.training_vs_reporting }}" alt="Training vs Reporting Chart">
    {{ training_vs_reporting|safe }}

    <h2>4. Gender-based Awareness</h2>
    <img src="{{ charts.gender_awareness }}" alt="Gender Awareness Chart">
    {{ gender_awareness|safe }}

    <p><em>All visualizations and conclusions are based on sample survey data.</em></p>
</body>
</html>
"""

# Write the HTML file
env = Environment()
template = env.from_string(html_template)
html_output = template.render(**template_data)

with open(REPORT_FILE, "w", encoding="utf-8") as f:
    f.write(html_output)

print(f"✅ Report generated: {REPORT_FILE}")

