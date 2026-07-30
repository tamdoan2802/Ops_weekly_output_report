import re
import json
import sys
import os

def parse_markdown_to_json(md_filepath, json_filepath):
    with open(md_filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    data = {
        "week_label": "Loading...",
        "overview": {
            "total_jobs": 0, "total_sqm": 0, "total_layouts": 0,
            "avg_leadtime": 0.0, "backlog_over_5d": 0
        },
        "teams": ["Frame & Truss", "Drafting", "Estimating", "Cleveland"],
        "trends": {
            "weeks": ["W-4", "W-3", "W-2", "W-1", "W0"],
            "jobs": {"Frame & Truss": [0,0,0,0,0], "Drafting": [0,0,0,0,0], "Estimating": [0,0,0,0,0], "Cleveland": [0,0,0,0,0]}
        },
        "kpi": [],
        "backlog": {
            "categories": ["0-1 ngày", "2 ngày", "3 ngày", "4 ngày", "5 ngày", ">5 ngày"],
            "data": {"Frame & Truss": [0,0,0,0,0,0], "Drafting": [0,0,0,0,0,0], "Estimating": [0,0,0,0,0,0], "Cleveland": [0,0,0,0,0,0]}
        },
        "complexity": {
            "categories": ["Frame & Truss", "Drafting", "Estimating", "Cleveland"],
            "design_mins": [103,493,528,79], "check_mins": [5,30,30,15]
        },
        "leaderboard": [
            {"name": "Han Pham", "team": "Frame & Truss", "layouts": 33, "sqm": 8157, "overtime": "+6.17h"},
            {"name": "Quy Vo", "team": "Drafting", "layouts": 284, "sqm": 1325, "overtime": "+0h"},
            {"name": "Thanh Nguyen", "team": "Estimating", "layouts": 79, "sqm": 20200, "overtime": "+2.12h"},
            {"name": "Truong Le", "team": "Cleveland", "layouts": 17, "sqm": 3955, "overtime": "+0.83h"}
        ]
    }

    # Extract Week Label
    for line in lines[:10]:
        if line.startswith("**Tuần báo cáo:**"):
            data["week_label"] = line.replace("**Tuần báo cáo:**", "").strip()
            break

    # Parse KPI Targets
    current_team = None
    current_customer = None
    kpi_target = 0
    kpi_actual = 0
    
    for i, line in enumerate(lines):
        line = line.strip()
        # Detect Team Section
        if line.startswith("# ") and "TEAM:" in line:
            if "FRAME & TRUSS" in line: current_team = "Frame & Truss"
            elif "DRAFTING" in line: current_team = "Drafting"
            elif "ESTIMATING" in line: current_team = "Estimating"
            elif "CLEVELAND" in line: current_team = "Cleveland"
            
        # Detect Customer
        if line.startswith("## ") and current_team and "Tổng hợp" not in line:
            m = re.search(r'## \d+\.\d+\.\s*(.+)', line)
            if m: current_customer = m.group(1).strip()
            
        # Detect Target KPI
        if line.startswith("**Loại khách hàng:**") and "Target" in line:
            m = re.search(r'Target\):\s*(\d+)', line)
            if m: kpi_target = int(m.group(1))
            else: kpi_target = 0
                
        # Detect Actual KPI
        if line.startswith("**KPI W0:**"):
            if kpi_target > 0:
                m = re.search(r'Thực tế =\s*\*\*(\d+)\*\*', line)
                if m:
                    kpi_actual = int(m.group(1))
                    data["kpi"].append({"customer": current_customer, "actual": kpi_actual, "target": kpi_target, "unit": "Jobs"})
                    
        # Parse Trends (Tổng Job hoàn thành của Team)
        if line.startswith("| **Tổng Job hoàn thành** |"):
            parts = [p.strip() for p in line.split("|") if p.strip()]
            if len(parts) >= 6 and current_team in data["trends"]["jobs"]:
                try:
                    vals = [int(re.sub(r'\D', '', p)) for p in parts[1:6]]
                    data["trends"]["jobs"][current_team] = vals
                except:
                    pass

        # Parse Backlog
        if "| **FRAME & TRUSS** |" in line: current_team = "Frame & Truss"
        if "| **DRAFTING** |" in line: current_team = "Drafting"
        if "| **ESTIMATING** |" in line: current_team = "Estimating"
        if "| **CLEVELAND** |" in line: current_team = "Cleveland"
        
        if line.startswith("| **") and current_team and current_team.upper() in line:
            parts = [p.strip().replace('**', '').replace('-', '0') for p in line.split("|") if p.strip()]
            if len(parts) >= 7:
                try:
                    data["backlog"]["data"][current_team] = [int(p) for p in parts[1:7]]
                except: pass

        # Parse Overview Stats (Total Sqm and Layouts for W0)
        if line.startswith("| Thiết kế: Diện tích (Sqm) |"):
            parts = [p.strip() for p in line.split("|") if p.strip()]
            if len(parts) >= 6:
                try:
                    val = float(parts[5].replace('m²', '').strip())
                    data["overview"]["total_sqm"] += val
                except: pass
                
        if line.startswith("| Thiết kế: Bản vẽ (Layouts) |"):
            parts = [p.strip() for p in line.split("|") if p.strip()]
            if len(parts) >= 6:
                try:
                    val = int(parts[5])
                    data["overview"]["total_layouts"] += val
                except: pass

    # Calculate Overview
    total_jobs = 0
    for t, vals in data["trends"]["jobs"].items():
        total_jobs += vals[4] # W0 is index 4
    data["overview"]["total_jobs"] = total_jobs
    
    # Backlog over 5d
    for t, vals in data["backlog"]["data"].items():
        data["overview"]["backlog_over_5d"] += vals[5]
        
    data["overview"]["avg_leadtime"] = 7.4 # Mocked avg leadtime

    # Save to JS file to avoid CORS on file:///
    os.makedirs(os.path.dirname(json_filepath), exist_ok=True)
    with open(json_filepath, 'w', encoding='utf-8') as f:
        f.write("const dashboardData = " + json.dumps(data, ensure_ascii=False, indent=2) + ";\n")
    print("Exported Dashboard JSON successfully.")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python parse_dashboard_json.py <md_file> <js_out>")
        sys.exit(1)
    parse_markdown_to_json(sys.argv[1], sys.argv[2])
