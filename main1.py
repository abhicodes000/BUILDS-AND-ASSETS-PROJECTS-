import pandas as pd
from pathlib import Path

def process_reports(input_file: str, output_file: str):
   

                                                                                   # Load input Excel
    df = pd.read_excel(input_file, engine="openpyxl")
    df['Date'] = pd.to_datetime(df['Date'])

                                                                               # Filter only missed tasks
    missed = df[df['Status'] == 'Not Done'].copy()
    missed.sort_values(by=['Employee', 'Task', 'Date'], inplace=True)

    results = []

                                                                               # Group by Employee + Task
    for (employee, task), group in missed.groupby(['Employee', 'Task']):
        dates = group['Date'].dt.strftime("%Y-%m-%d").tolist()

                                                                        # Always log the first miss as Follow-up n Extend consecutive block
        results.append({
            "Employee": employee,
            "Task": task,
            "Date(s) Missed": dates[0],
            "Action": "Follow-up"
        })

        block = [dates[0]]
        for i in range(1, len(dates)):
            prev, curr = pd.to_datetime(dates[i-1]), pd.to_datetime(dates[i])
            if (curr - prev).days == 1:
                
                block.append(dates[i])
                
                results.append({
                    "Employee": employee,
                    "Task": task,
                    "Date(s) Missed": ", ".join(block),
                    "Action": "Escalate"
                })
            else:
             
                block = [dates[i]]
                results.append({
                    "Employee": employee,
                    "Task": task,
                    "Date(s) Missed": dates[i],
                    "Action": "Follow-up"
                })

                                                                                         # Save to Excel
    output_df = pd.DataFrame(results, columns=["Employee", "Task", "Date(s) Missed", "Action"])

                                                                                       
    output_df.to_excel(output_file, index=False, engine="openpyxl")
    print(f" Report generated successfully: {output_file}")

if __name__ == "__main__":
    input_path = Path("C://Users//abhis//Downloads//Candidate_Task_DailyReports.xlsx")
    output_path = Path("DailyReport_FollowupEscalate.xlsx")
    process_reports(input_path, output_path)
