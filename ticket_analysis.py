import pandas as pd
import re
import os

# 1. Load the Dataset
file_path = r"C:\Users\dhana\OneDrive\Desktop\future intern\customer_support_tickets.csv"  # Adjust if needed
df = pd.read_csv(file_path)

# Show available columns
print("Available columns:\n", df.columns)

# ✅ Use the correct column for issue description
issue_column = 'Ticket Description'

# 2. Clean Text
def clean_text(text):
    text = str(text).lower()
    text = re.sub(r'[^a-z\s]', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

df['clean_issue'] = df[issue_column].apply(clean_text)

# 3. Categorize Issues
def categorize_issue(text):
    if "login" in text:
        return "Login Issues"
    elif "payment" in text:
        return "Payment Issues"
    elif "error" in text or "fail" in text or "bug" in text:
        return "Technical Error"
    elif "slow" in text or "delay" in text:
        return "Performance Issue"
    elif "cancel" in text or "refund" in text:
        return "Cancellation/Refund"
    else:
        return "Other"

df['Category'] = df['clean_issue'].apply(categorize_issue)

# 4. Frequency of Issues
issue_counts = df['Category'].value_counts()
print("\nTop Reported Issues:\n", issue_counts)

# 5. Response Time Analysis
if 'First Response Time' in df.columns and 'Time to Resolution' in df.columns:
    df['First Response Time'] = pd.to_timedelta(df['First Response Time'], errors='coerce')
    df['Time to Resolution'] = pd.to_timedelta(df['Time to Resolution'], errors='coerce')
    avg_response = df.groupby('Category')['First Response Time'].mean().sort_values()
    avg_resolution = df.groupby('Category')['Time to Resolution'].mean().sort_values()

    print("\nAverage First Response Time by Issue Type:\n", avg_response)
    print("\nAverage Time to Resolution by Issue Type:\n", avg_resolution)
else:
    print("\nResponse time columns not found.")

# 6. Export to Excel
output_file = "ticket_analysis_output.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Detailed_Tickets', index=False)
    issue_counts.to_frame(name='Count').to_excel(writer, sheet_name='Issue_Frequency')
    if 'First Response Time' in df.columns and 'Time to Resolution' in df.columns:
        avg_response.to_frame(name='Avg_First_Response_Time').to_excel(writer, sheet_name='Avg_Response')
        avg_resolution.to_frame(name='Avg_Resolution_Time').to_excel(writer, sheet_name='Avg_Resolution')

print(f"\n✅ Analysis complete. Results saved to: {os.path.abspath(output_file)}")
