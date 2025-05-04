import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Step 1: Load Excel data
file_path = "data.xlsx"
df = pd.read_excel(file_path)

# Step 2: Clean column names for SQL compatibility
df.columns = [col.strip().replace(" ", "_").replace("?", "").replace(",", "").replace("(", "").replace(")", "").replace("/", "_") for col in df.columns]

# Step 3: Create SQLite DB and insert data
conn = sqlite3.connect("harassment_survey.db")
df.to_sql("survey", conn, if_exists="replace", index=False)

# Step 4: Run SQL Queries for Insights

# 1. Overall Awareness
awareness_df = pd.read_sql_query("""
    SELECT Are_you_aware_of_what_constitutes_workplace_harassment AS Awareness, COUNT(*) AS Count
    FROM survey
    GROUP BY Awareness;
""", conn)

# 2. Awareness vs Knowledge of Reporting
awareness_reporting_df = pd.read_sql_query("""
    SELECT 
        Are_you_aware_of_what_constitutes_workplace_harassment AS Awareness,
        Do_you_know_whom_to_report_workplace_harassment_incidents_to AS Know_Whom_To_Report,
        COUNT(*) AS Count
    FROM survey
    GROUP BY Awareness, Know_Whom_To_Report;
""", conn)

# 3. Training vs Reporting Knowledge
training_reporting_df = pd.read_sql_query("""
    SELECT 
        Have_you_ever_received_any_formal_training_on_workplace_harassment_policies AS Training,
        Do_you_know_whom_to_report_workplace_harassment_incidents_to AS Know_Whom_To_Report,
        COUNT(*) AS Count
    FROM survey
    GROUP BY Training, Know_Whom_To_Report;
""", conn)

# 4. Gender-based Awareness
gender_awareness_df = pd.read_sql_query("""
    SELECT 
        What_is_your_Gender AS Gender,
        Are_you_aware_of_what_constitutes_workplace_harassment AS Awareness,
        COUNT(*) AS Count
    FROM survey
    GROUP BY Gender, Awareness;
""", conn)

conn.close()

# Step 5: Plot Charts
os.makedirs("charts", exist_ok=True)

# Plot 1: Overall Awareness
plt.figure(figsize=(6, 4))
sns.barplot(data=awareness_df, x='Awareness', y='Count', palette='pastel')
plt.title("Overall Awareness of Workplace Harassment")
plt.ylabel("Number of Respondents")
plt.tight_layout()
plt.savefig("charts/awareness_chart.png")
plt.close()

# Plot 2: Awareness vs Knowledge of Reporting
pivot1 = awareness_reporting_df.pivot(index='Awareness', columns='Know_Whom_To_Report', values='Count').fillna(0)
pivot1.plot(kind='bar', stacked=True, colormap='coolwarm', figsize=(8, 5))
plt.title("Awareness vs Knowledge of Reporting")
plt.ylabel("Number of Respondents")
plt.tight_layout()
plt.savefig("charts/awareness_vs_reporting.png")
plt.close()

# Plot 3: Training vs Knowledge of Reporting
pivot2 = training_reporting_df.pivot(index='Training', columns='Know_Whom_To_Report', values='Count').fillna(0)
pivot2.plot(kind='bar', stacked=True, colormap='viridis', figsize=(8, 5))
plt.title("Training vs Knowledge of Reporting")
plt.ylabel("Number of Respondents")
plt.tight_layout()
plt.savefig("charts/training_vs_reporting.png")
plt.close()

# Plot 4: Gender-based Awareness
pivot3 = gender_awareness_df.pivot(index='Gender', columns='Awareness', values='Count').fillna(0)
pivot3.plot(kind='bar', stacked=True, colormap='Set2', figsize=(8, 5))
plt.title("Gender-based Awareness of Workplace Harassment")
plt.ylabel("Number of Respondents")
plt.tight_layout()
plt.savefig("charts/gender_awareness.png")
plt.close()

# Step 6: Print Conclusive Inferences
print("=== Research-Grade Inferences ===\n")

# Inference 1: Awareness
total = awareness_df['Count'].sum()
yes_pct = awareness_df[awareness_df['Awareness'] == 'Yes']['Count'].values[0] / total * 100
print(f"1. Awareness of Harassment: {yes_pct:.1f}% of respondents are fully aware of what constitutes workplace harassment.\n"
      "   ➤ This high level of awareness provides a strong base for assessing policy effectiveness.\n")

# Inference 2: Awareness vs Reporting
print("2. Awareness vs Knowledge of Reporting:")
print(awareness_reporting_df)
print("   ➤ Helps identify if awareness translates into actionability (knowing where to report).\n")

# Inference 3: Training vs Knowledge of Reporting
print("3. Training vs Reporting Knowledge:")
print(training_reporting_df)
print("   ➤ If those who received training show greater reporting knowledge, it confirms training efficacy.\n")

# Inference 4: Gender Differences
print("4. Gender-based Awareness:")
print(gender_awareness_df)
print("   ➤ Useful for gender-targeted interventions if disparities in awareness exist.\n")

print("✅ All charts saved to the 'charts/' directory.")

