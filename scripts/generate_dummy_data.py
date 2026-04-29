import pandas as pd
import os

data = {
    "Employee Name": ["Alice Johnson", "Bob Smith", "Charlie Davis", "Diana Evans"],
    "Task Name": ["Backend API Integration", "Frontend UI Design", "Database Migration", "Testing & QA"],
    "Progress %": [85, 40, 100, 60],
    "Status": ["In Progress", "In Progress", "Completed", "In Progress"]
}

df = pd.DataFrame(data)

# Ensure the directory exists
os.makedirs("data", exist_ok=True)

df.to_excel("data/tasks.xlsx", index=False)
print("Dummy data generated at data/tasks.xlsx")
