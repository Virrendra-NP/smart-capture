import os
import sys
import traceback

with open("mpp_app_debug.txt", "w") as f:
    f.write("Starting script...\n")

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    import datetime
    with open("mpp_app_debug.txt", "a") as f:
        f.write("Imported tkinter...\n")

    import win32com.client
    import xlsxwriter
    with open("mpp_app_debug.txt", "a") as f:
        f.write("Imported win32com and xlsxwriter...\n")

    def get_task_dates(task):
        def fmt_date(d):
            if d:
                try:
                    return d.strftime("%d-%b-%Y")
                except:
                    return str(d)[:10]
            return ""
        return {
            'start': fmt_date(task.Start),
            'finish': fmt_date(task.Finish),
            'baseline_start': fmt_date(task.BaselineStart),
            'baseline_finish': fmt_date(task.BaselineFinish),
            'actual_start': fmt_date(task.ActualStart),
            'actual_finish': fmt_date(task.ActualFinish),
        }

    def process_mpp():
        mpp_path = mpp_var.get()
        excel_path = excel_var.get()
        
        if not mpp_path or not os.path.exists(mpp_path):
            messagebox.showerror("Error", "Please select a valid MS Project (.mpp) file.")
            return
            
        if not excel_path:
            messagebox.showerror("Error", "Please select a valid output Excel file path.")
            return

        transfer_btn.config(state="disabled", text="Reading MS Project...")
        root.update()

        mpp_app = None
        try:
            mpp_app = win32com.client.Dispatch("MSProject.Application")
            mpp_app.Visible = False
            mpp_app.DisplayAlerts = False
            
            # Use FileOpenEx for safety without prompting
            mpp_app.FileOpenEx(mpp_path, True)
            project_ext = mpp_app.ActiveProject
            
            tasks_data = []
            for task in project_ext.Tasks:
                if not task:
                    continue
                    
                dates = get_task_dates(task)
                row_data = {
                    'wbs': getattr(task, 'WBS', ''),
                    'name': getattr(task, 'Name', ''),
                    'level': getattr(task, 'OutlineLevel', 1),
                    'is_summary': getattr(task, 'Summary', False),
                    'start': dates['start'],
                    'finish': dates['finish'],
                    'base_start': dates['baseline_start'],
                    'base_finish': dates['baseline_finish'],
                    'act_start': dates['actual_start'],
                    'act_finish': dates['actual_finish'],
                    'percent_comp': getattr(task, 'PercentComplete', 0),
                }
                
                # Determine status
                pct = row_data['percent_comp']
                row_data['status'] = 'Completed' if pct == 100 else ('In Progress' if pct > 0 else 'Not Started')
                
                tasks_data.append(row_data)

            mpp_app.FileClose(0)
            
            # --- write to excel ---
            transfer_btn.config(text="Building Excel Dashboard...")
            root.update()
            
            workbook = xlsxwriter.Workbook(excel_path)
            title_format = workbook.add_format({'bold': True, 'font_size': 18, 'font_color': '#FFFFFF', 'bg_color': '#1F4E78', 'align': 'center', 'valign': 'vcenter'})
            header_format = workbook.add_format({
                'bold': True, 'font_color': 'white', 'bg_color': '#2F75B5',
                'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
            })
            
            level1_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'valign': 'vcenter'})
            level2_format = workbook.add_format({'bold': True, 'bg_color': '#EDF2F9', 'border': 1, 'valign': 'vcenter'})
            standard_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
            
            dash_sheet = workbook.add_worksheet("Project Dashboard")
            dash_sheet.hide_gridlines(2)
            dash_sheet.merge_range("A1:K2", f"PROJECT SCHEDULE DASHBOARD", title_format)
            
            headers = ["WBS", "Task Name", "Status", "% Complete", "Baseline Start", "Baseline Finish", "Actual Start", "Actual Finish", "Planned Start", "Planned Finish"]
            for col, h in enumerate(headers):
                dash_sheet.write(3, col, h, header_format)
                
            dash_sheet.set_row(3, 30)
            dash_sheet.set_column('A:A', 10)
            dash_sheet.set_column('B:B', 50)
            dash_sheet.set_column('C:C', 15)
            dash_sheet.set_column('D:D', 12)
            dash_sheet.set_column('E:J', 14)
            
            row_idx = 4
            for task in tasks_data:
                if task['is_summary'] and task['level'] == 1:
                    row_fmt = level1_format
                    name_indent = ""
                elif task['is_summary']:
                    row_fmt = level2_format
                    name_indent = "   " * (task['level'] - 1)
                else:
                    row_fmt = standard_format
                    name_indent = "   " * (task['level'] - 1)
                
                c_fmt = workbook.add_format(row_fmt.__dict__.copy())
                c_fmt.set_align('center')
                p_fmt = workbook.add_format(row_fmt.__dict__.copy())
                p_fmt.set_align('center')
                p_fmt.set_num_format('0"%"')
                
                dash_sheet.write_string(row_idx, 0, task['wbs'], c_fmt)
                dash_sheet.write_string(row_idx, 1, name_indent + task['name'], row_fmt)
                dash_sheet.write_string(row_idx, 2, task['status'], c_fmt)
                dash_sheet.write_number(row_idx, 3, task['percent_comp'], p_fmt)
                dash_sheet.write_string(row_idx, 4, task['base_start'], c_fmt)
                dash_sheet.write_string(row_idx, 5, task['base_finish'], c_fmt)
                dash_sheet.write_string(row_idx, 6, task['act_start'], c_fmt)
                dash_sheet.write_string(row_idx, 7, task['act_finish'], c_fmt)
                dash_sheet.write_string(row_idx, 8, task['start'], c_fmt)
                dash_sheet.write_string(row_idx, 9, task['finish'], c_fmt)
                
                row_idx += 1
                
            dash_sheet.conditional_format(f'D5:D{row_idx}', {'type': 'data_bar', 'bar_color': '#63C384', 'min_value': 0, 'max_value': 100})
            workbook.close()
            
            messagebox.showinfo("Success", f"Dashboard Generation Complete!\nExported {len(tasks_data)} tasks.")
            
        except Exception as e:
            with open("mpp_app_debug.txt", "a") as f:
                f.write(f"Error in process: {traceback.format_exc()}\n")
            messagebox.showerror("Error", f"Failed! See mpp_app_debug.txt.\n{e}")
        finally:
            if mpp_app:
                mpp_app.Quit()
            transfer_btn.config(state="normal", text="Generate Dashboard")

    def select_mpp():
        path = filedialog.askopenfilename(title="Select MS Project File", filetypes=[("MS Project", "*.mpp")])
        if path: mpp_var.set(path)

    def select_excel():
        path = filedialog.asksaveasfilename(title="Save Dashboard As", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path: excel_var.set(path)

    with open("mpp_app_debug.txt", "a") as f:
        f.write("Starting GUI...\n")

    root = tk.Tk()
    root.title("MS Project Smart Exporter Dashboard")
    root.geometry("620x280")
    root.eval('tk::PlaceWindow . center')

    tk.Label(root, text="Select Microsoft Project File (.mpp):", font=("Arial", 10, "bold")).pack(pady=(15, 2))
    mpp_var = tk.StringVar(value=r"D:\JKD_Folder\JKD-PROJECT SITE\DPR\Excel_DPR\project_schedule.mpp")
    frame_mpp = tk.Frame(root)
    frame_mpp.pack(fill="x", padx=20)
    tk.Entry(frame_mpp, textvariable=mpp_var, width=65).pack(side="left", padx=(0, 10))
    tk.Button(frame_mpp, text="Browse", command=select_mpp).pack(side="left")

    tk.Label(root, text="Destination Dashboard Excel File:", font=("Arial", 10, "bold")).pack(pady=(15, 2))
    excel_var = tk.StringVar(value=r"D:\JKD_Folder\JKD-PROJECT SITE\DPR\Excel_DPR\schedule_dashboard.xlsx")
    frame_excel = tk.Frame(root)
    frame_excel.pack(fill="x", padx=20)
    tk.Entry(frame_excel, textvariable=excel_var, width=65).pack(side="left", padx=(0, 10))
    tk.Button(frame_excel, text="Browse", command=select_excel).pack(side="left")

    transfer_btn = tk.Button(root, text="Generate Dashboard", command=process_mpp, bg="#1F4E78", fg="white", font=("Arial", 12, "bold"), padx=20, pady=5)
    transfer_btn.pack(pady=25)

    root.mainloop()

except Exception as e:
    with open("mpp_app_debug.txt", "a") as f:
        f.write(f"FATAL ERROR:\n{traceback.format_exc()}\n")
