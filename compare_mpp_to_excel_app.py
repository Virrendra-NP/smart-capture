"""
MS Project Schedule Comparison Dashboard
-----------------------------------------
Compare two .mpp schedules (Baseline vs Current) and produce
a presentable Excel dashboard on sheet 'Comp-1'.

Uses a subprocess worker to talk to MS Project via COM, so if
COM crashes, this GUI stays alive and shows a proper error.
"""

import os
import sys
import json
import subprocess
import datetime
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox


WORKER_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_worker.py")


def install_deps():
    missing = []
    try: import xlsxwriter
    except ImportError: missing.append("xlsxwriter")
    try: import win32com.client
    except ImportError: missing.append("pywin32")
    if missing:
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", *missing],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            return False, str(e)
    return True, ""


def read_mpp_via_subprocess(mpp_path, label_text):
    """Run _worker.py in a child process and return extracted tasks."""
    transfer_btn.config(text=label_text)
    root.update()

    tmp = tempfile.NamedTemporaryFile(suffix=".json", delete=False, dir=tempfile.gettempdir())
    tmp_path = tmp.name
    tmp.close()

    try:
        result = subprocess.run(
            [sys.executable, WORKER_SCRIPT, mpp_path, tmp_path],
            capture_output=True, text=True, timeout=120
        )

        if not os.path.exists(tmp_path) or os.path.getsize(tmp_path) == 0:
            stderr = result.stderr[:500] if result.stderr else "No details available"
            raise RuntimeError(
                f"MS Project COM worker crashed or timed out.\n"
                f"Return code: {result.returncode}\n"
                f"Error output:\n{stderr}")

        with open(tmp_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        if not data.get("ok"):
            raise RuntimeError(f"Worker reported error:\n{data.get('error', 'Unknown')}")

        return data["tasks"]

    finally:
        try: os.unlink(tmp_path)
        except: pass


def parse_date(s):
    """Parse a date string returned by the worker."""
    if not s:
        return None
    try:
        return datetime.datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        try:
            return datetime.datetime.strptime(s[:10], "%Y-%m-%d")
        except ValueError:
            pass
    return None


def fmt_date(dt):
    if dt:
        return dt.strftime("%d-%b-%Y")
    return ""


def process_comparison():
    mpp1 = mpp1_var.get()
    mpp2 = mpp2_var.get()
    xl   = excel_var.get()

    if not mpp1 or not mpp1.lower().endswith(".mpp"):
        messagebox.showerror("Error", "Please select a valid Baseline .mpp file.")
        return
    if not mpp2 or not mpp2.lower().endswith(".mpp"):
        messagebox.showerror("Error", "Please select a valid Current .mpp file.")
        return
    if not xl:
        messagebox.showerror("Error", "Please select an output Excel path.")
        return

    transfer_btn.config(state="disabled", text="Checking dependencies...")
    root.update()

    ok, err = install_deps()
    if not ok:
        messagebox.showerror("Error", f"Cannot install dependencies:\n{err}")
        transfer_btn.config(state="normal", text="Compare Schedules & Generate")
        return

    import xlsxwriter

    try:
        # ── extract baseline tasks ──
        baseline_list = read_mpp_via_subprocess(mpp1, "Extracting Baseline Schedule...")

        base_by_uid = {}
        base_by_name = {}
        for t in baseline_list:
            if t["uid"]:
                base_by_uid[t["uid"]] = t
            if t["name"]:
                base_by_name[t["name"]] = t

        # ── extract current tasks ──
        current_list = read_mpp_via_subprocess(mpp2, "Extracting Current Schedule...")

        # ── compare ──
        transfer_btn.config(text="Comparing Schedules...")
        root.update()

        rows = []
        for cur in current_list:
            base = base_by_uid.get(cur["uid"])
            if not base:
                base = base_by_name.get(cur["name"])

            b_start  = parse_date(base["start"])  if base else None
            b_finish = parse_date(base["finish"]) if base else None
            c_start  = parse_date(cur["start"])
            c_finish = parse_date(cur["finish"])

            # Variance = Baseline - Current
            #   negative → delayed (current is later than baseline)
            #   positive → ahead  (current is earlier)
            start_var  = (b_start  - c_start).days  if (b_start  and c_start)  else 0
            finish_var = (b_finish - c_finish).days if (b_finish and c_finish) else 0

            indicator = "On Time"
            if finish_var < 0:
                indicator = f"Delayed ({abs(finish_var)}d)"
            elif finish_var > 0:
                indicator = f"Ahead ({finish_var}d)"
            elif start_var < 0:
                indicator = f"Start Delayed ({abs(start_var)}d)"
            elif start_var > 0:
                indicator = f"Start Ahead ({start_var}d)"

            pct = cur.get("pct", 0) or 0
            status = "Completed" if pct == 100 else ("In Progress" if pct > 0 else "Not Started")

            rows.append({
                "wbs": cur["wbs"],
                "name": cur["name"],
                "level": cur["level"],
                "is_summary": cur["is_summary"],
                "status": status,
                "b_start": fmt_date(b_start),
                "b_finish": fmt_date(b_finish),
                "c_start": fmt_date(c_start),
                "c_finish": fmt_date(c_finish),
                "start_var": start_var,
                "finish_var": finish_var,
                "indicator": indicator,
            })

        # ── write Excel dashboard ──
        transfer_btn.config(text="Building Dashboard...")
        root.update()

        os.makedirs(os.path.dirname(xl), exist_ok=True)
        wb = xlsxwriter.Workbook(xl)

        title_fmt = wb.add_format({
            "bold": True, "font_size": 18, "font_color": "#FFF",
            "bg_color": "#833C0C", "align": "center", "valign": "vcenter"})
        hdr_fmt = wb.add_format({
            "bold": True, "font_color": "#FFF", "bg_color": "#C55A11",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})

        l1  = {"bold": True, "bg_color": "#FCE4D6", "border": 1, "valign": "vcenter"}
        l2  = {"bold": True, "bg_color": "#FDF2E9", "border": 1, "valign": "vcenter"}
        std = {"border": 1, "valign": "vcenter"}

        fc = {
            "l1":   wb.add_format(l1),
            "l2":   wb.add_format(l2),
            "std":  wb.add_format(std),
            "l1c":  wb.add_format({**l1, "align": "center"}),
            "l2c":  wb.add_format({**l2, "align": "center"}),
            "stdc": wb.add_format({**std, "align": "center"}),
            "l1n":  wb.add_format({**l1, "align": "center", "num_format": "#,##0"}),
            "l2n":  wb.add_format({**l2, "align": "center", "num_format": "#,##0"}),
            "stdn": wb.add_format({**std, "align": "center", "num_format": "#,##0"}),
        }
        delay_f = wb.add_format({
            "bg_color": "#FFC7CE", "font_color": "#9C0006",
            "align": "center", "valign": "vcenter", "border": 1})
        ahead_f = wb.add_format({
            "bg_color": "#C6EFCE", "font_color": "#006100",
            "align": "center", "valign": "vcenter", "border": 1})

        ws = wb.add_worksheet("Comp-1")
        ws.hide_gridlines(2)
        ws.merge_range("A1:J2", "SCHEDULE COMPARISON DASHBOARD", title_fmt)

        headers = [
            "WBS", "Task Name", "Status",
            "Baseline Start", "Baseline Finish",
            "Current Start", "Current Finish",
            "Start Variance (Days)", "Finish Variance (Days)",
            "Delay Indicator"]
        for c, h in enumerate(headers):
            ws.write(3, c, h, hdr_fmt)

        ws.set_row(3, 40)
        ws.set_column("A:A", 10)
        ws.set_column("B:B", 45)
        ws.set_column("C:C", 15)
        ws.set_column("D:G", 16)
        ws.set_column("H:I", 20)
        ws.set_column("J:J", 22)

        r = 4
        for t in rows:
            if t["is_summary"] and t["level"] <= 1:
                indent, k = "", "l1"
            elif t["is_summary"]:
                indent, k = "   " * (t["level"] - 1), "l2"
            else:
                indent, k = "   " * (t["level"] - 1), "std"

            f  = fc[k]
            cf = fc[k + "c"]
            nf = fc[k + "n"]

            ws.write_string(r, 0, t["wbs"], cf)
            ws.write_string(r, 1, indent + t["name"], f)
            ws.write_string(r, 2, t["status"], cf)
            ws.write_string(r, 3, t["b_start"], cf)
            ws.write_string(r, 4, t["b_finish"], cf)
            ws.write_string(r, 5, t["c_start"], cf)
            ws.write_string(r, 6, t["c_finish"], cf)
            ws.write_number(r, 7, t["start_var"], nf)
            ws.write_number(r, 8, t["finish_var"], nf)
            ws.write_string(r, 9, t["indicator"], cf)
            r += 1

        ws.conditional_format(f"J5:J{r}", {
            "type": "text", "criteria": "containing",
            "value": "Delayed", "format": delay_f})
        ws.conditional_format(f"J5:J{r}", {
            "type": "text", "criteria": "containing",
            "value": "Ahead", "format": ahead_f})

        wb.close()

        messagebox.showinfo("Success",
            f"Dashboard complete!\n"
            f"Compared {len(rows)} tasks.\n"
            f"Saved to 'Comp-1' sheet in:\n{xl}")

    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        err_file = os.path.join(os.path.dirname(xl), "mpp_comparison_error.txt")
        try:
            with open(err_file, "w") as fh:
                fh.write(err_detail)
        except: pass
        messagebox.showerror("Error",
            f"Something went wrong:\n\n{e}\n\nDetails saved to:\n{err_file}")
    finally:
        transfer_btn.config(state="normal", text="Compare Schedules & Generate")


# ── GUI ──────────────────────────────────────────────────────────────
def pick_mpp1():
    p = filedialog.askopenfilename(title="Select Baseline MS Project File",
                                   filetypes=[("MS Project", "*.mpp")])
    if p: mpp1_var.set(p)

def pick_mpp2():
    p = filedialog.askopenfilename(title="Select Current MS Project File",
                                   filetypes=[("MS Project", "*.mpp")])
    if p: mpp2_var.set(p)

def pick_excel():
    p = filedialog.asksaveasfilename(title="Save Dashboard As",
                                     defaultextension=".xlsx",
                                     filetypes=[("Excel", "*.xlsx")])
    if p: excel_var.set(p)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("MS Project Schedule Comparison Dashboard")
    root.geometry("700x380")
    root.eval("tk::PlaceWindow . center")

    tk.Label(root, text="Step 1: Select BASELINE MS Project File (.mpp)",
             font=("Arial", 10, "bold")).pack(pady=(12, 2))
    mpp1_var = tk.StringVar(value="")
    f1 = tk.Frame(root); f1.pack(fill="x", padx=20)
    tk.Entry(f1, textvariable=mpp1_var, width=78).pack(side="left", padx=(0, 8))
    tk.Button(f1, text="Browse", command=pick_mpp1).pack(side="left")

    tk.Label(root, text="Step 2: Select CURRENT / UN-BASELINED MS Project File (.mpp)",
             font=("Arial", 10, "bold")).pack(pady=(12, 2))
    mpp2_var = tk.StringVar(value="")
    f2 = tk.Frame(root); f2.pack(fill="x", padx=20)
    tk.Entry(f2, textvariable=mpp2_var, width=78).pack(side="left", padx=(0, 8))
    tk.Button(f2, text="Browse", command=pick_mpp2).pack(side="left")

    tk.Label(root, text="Output Dashboard Excel File:",
             font=("Arial", 10, "bold")).pack(pady=(15, 2))
    excel_var = tk.StringVar(
        value=r"D:\JKD_Folder\JKD-PROJECT SITE\DPR\Excel_DPR\schedule_dashboard.xlsx")
    f3 = tk.Frame(root); f3.pack(fill="x", padx=20)
    tk.Entry(f3, textvariable=excel_var, width=78).pack(side="left", padx=(0, 8))
    tk.Button(f3, text="Browse", command=pick_excel).pack(side="left")

    transfer_btn = tk.Button(root, text="Compare Schedules & Generate",
                              command=process_comparison,
                              bg="#C55A11", fg="white",
                              font=("Arial", 12, "bold"), padx=20, pady=5)
    transfer_btn.pack(pady=22)

    root.mainloop()
