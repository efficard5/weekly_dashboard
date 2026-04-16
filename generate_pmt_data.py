import pandas as pd
import os

columns = [
    "Task ID", "PM", "Project Number", "Assigned Date", "Station Name", "Task Name",
    "Assignee Email", "Responsible Person", "Department", "Due Date", "Planned Hrs",
    "Planned % Completion", "Actual Hrs", "Actual % Completion", "Status", "Re Shd",
    "Remarks", "Proof Link", "Deviation % in Completion", "Escalation to Hr Management",
    "Time Stamp", "Task Completion Approval Status from HOD", 
    "Task Completion Non Approval Status from HOD", "HOD REMARKS", "Uninformed Leave",
    "Notification Sent", "Task Completion Approval Status from PMT",
    "Task Completion Non Approval Status from PMT", "PMT Remarks", "Hrs Deviation", 
    "Task Approval Status"
]

data = [
    [1039, "SUDHARSAN", "R&D", "4/8/2026 8:00:00", "", "Dashboard creation for weekly update", "efficard5@gmail.com", "Gowtham-R&D", "R&D", "4/8/2026 16:30:00", 2, 100, "", "", "FALSE", "FALSE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
    [1040, "SUDHARSAN", "R&D", "4/8/2026 8:00:00", "", "Animation creation using antigravity check", "efficard5@gmail.com", "Gowtham-R&D", "R&D", "4/8/2026 16:30:00", 3, 100, "", "", "FALSE", "FALSE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
    [1041, "SUDHARSAN", "R&D", "4/8/2026 8:00:00", "", "Brush carry forward in R&d system", "efficard5@gmail.com", "Gowtham-R&D", "R&D", "4/8/2026 16:30:00", 2, 100, "", "", "FALSE", "FALSE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
    [1042, "SUDHARSAN", "R&D", "4/8/2026 8:00:00", "", "Dataset creation using roboflow - Flowchart for it", "efficard5@gmail.com", "Gowtham-R&D", "R&D", "4/8/2026 16:30:00", 1, 100, "", "", "FALSE", "FALSE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
]

df = pd.DataFrame(data, columns=columns)

os.makedirs("data", exist_ok=True)
df.to_excel("data/pmt_tracker.xlsx", index=False)
print("pmt_tracker.xlsx generated successfully with 31 columns.")
