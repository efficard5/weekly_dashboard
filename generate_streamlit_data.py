import pandas as pd
from datetime import datetime, timedelta
import os
import random

# Generate dates for the current week
today = datetime.now()
dates = [(today - timedelta(days=i)).strftime("%Y-%m-%d") for i in range(5)]

persons = ["Alice", "Bob", "Charlie", "Diana"]
projects = ["Automation Web", "Data Pipeline", "Vision Beta"]
statuses = ["Live", "Holded", "Postponed"]

data = {
    "Date": [],
    "Person Name": [],
    "Project": [],
    "Completion %": [],
    "Errors Faced": [],
    "Trials Taken": [],
    "Status": []
}

for dateStr in dates:
    for person in persons:
        project = random.choice(projects)
        completion = random.randint(10, 100)
        errors = random.randint(0, 5)
        trials = random.randint(1, 10)
        
        status = "Live"
        if completion == 100:
            status = "Live"
        else:
            status = random.choice(statuses)

        data["Date"].append(dateStr)
        data["Person Name"].append(person)
        data["Project"].append(project)
        data["Completion %"].append(completion)
        data["Errors Faced"].append(errors)
        data["Trials Taken"].append(trials)
        data["Status"].append(status)

df = pd.DataFrame(data)

os.makedirs("data", exist_ok=True)
df.to_excel("data/team_tracker.xlsx", index=False)
print("Team tracker data generated at data/team_tracker.xlsx")
