"""
_worker.py  –  Subprocess worker for extracting tasks from an .mpp file
via win32com.  This runs in a child process so that if COM crashes,
the main GUI keeps running.

Usage:  python _worker.py  <mpp_path>  <output_json_path>
Writes a JSON file with the extracted tasks list.
"""

import os
import sys
import json
import traceback
import datetime


def format_date(d):
    """Safely convert any COM date to a string."""
    if d is None or d == "":
        return ""
    try:
        return d.strftime("%Y-%m-%d %H:%M:%S")
    except (AttributeError, TypeError):
        s = str(d).strip()
        if s and s != "None":
            return s
    return ""


def extract_tasks(mpp_path):
    import win32com.client

    mpp_app = win32com.client.Dispatch("MSProject.Application")
    mpp_app.Visible = False
    mpp_app.DisplayAlerts = False
    mpp_app.FileOpenEx(mpp_path, True)  # True = ReadOnly
    project = mpp_app.ActiveProject

    results = []
    for task in project.Tasks:
        if not task:
            continue

        uid = 0
        wbs = ""
        name = ""
        lvl = 1
        is_summary = False
        pct = 0
        start = ""
        finish = ""

        try: uid = int(task.UniqueID)
        except: pass
        try: wbs = str(task.WBS)
        except: pass
        try: name = str(task.Name)
        except: pass
        try: lvl = int(task.OutlineLevel)
        except: pass
        try: is_summary = bool(task.Summary)
        except: pass
        try: pct = int(task.PercentComplete)
        except: pass
        try: start = format_date(task.Start)
        except: pass
        try: finish = format_date(task.Finish)
        except: pass

        results.append({
            "uid": uid,
            "wbs": wbs,
            "name": name,
            "level": lvl,
            "is_summary": is_summary,
            "pct": pct,
            "start": start,
            "finish": finish,
        })

    mpp_app.FileClose(0)
    mpp_app.Quit()
    return results


if __name__ == "__main__":
    mpp_path = sys.argv[1]
    output_json = sys.argv[2]

    try:
        tasks = extract_tasks(mpp_path)
        with open(output_json, "w", encoding="utf-8") as f:
            json.dump({"ok": True, "tasks": tasks}, f, indent=2)
    except Exception:
        with open(output_json, "w", encoding="utf-8") as f:
            json.dump({"ok": False, "error": traceback.format_exc()}, f, indent=2)
