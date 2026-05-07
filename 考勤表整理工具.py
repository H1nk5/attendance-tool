# -*- coding: utf-8 -*-
"""
考勤表整理工具
功能：从总考勤表中按姓名列表筛选人员，生成独立考勤表
依赖：pip install pywin32
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

try:
    import win32com.client
except ImportError:
    win32com = None


class AttendanceApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("考勤表整理工具 v1.2")
        self.root.geometry("720x680")
        self.root.resizable(True, True)

        self.source_file: str = ""
        self.employees: list[dict] = []

        self._build_ui()

    # ── UI ──────────────────────────────────────────────

    def _build_ui(self):
        f1 = ttk.LabelFrame(self.root, text="1. 选择总考勤表", padding=10)
        f1.pack(fill="x", padx=10, pady=(10, 5))

        self.var_file = tk.StringVar()
        ttk.Entry(f1, textvariable=self.var_file, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 10)
        )
        ttk.Button(f1, text="浏览...", command=self._select_file).pack(side="right")

        f2 = ttk.LabelFrame(self.root, text="2. 输入人员名单（每行一个姓名）", padding=10)
        f2.pack(fill="both", expand=True, padx=10, pady=5)

        toolbar = ttk.Frame(f2)
        toolbar.pack(fill="x", pady=(0, 5))
        ttk.Button(toolbar, text="从总表读取全部姓名", command=self._read_names).pack(side="left")
        ttk.Button(toolbar, text="清空", command=self._clear_names).pack(side="left", padx=(10, 0))
        ttk.Label(toolbar, text="  (也可直接粘贴姓名列表)").pack(side="left")

        self.txt_names = scrolledtext.ScrolledText(f2, height=12, font=("微软雅黑", 10))
        self.txt_names.pack(fill="both", expand=True)
        self.txt_names.insert("1.0", "张三\n李四\n王五")

        f3 = ttk.LabelFrame(self.root, text="3. 生成", padding=10)
        f3.pack(fill="x", padx=10, pady=5)

        dir_f = ttk.Frame(f3)
        dir_f.pack(fill="x", pady=(0, 5))
        ttk.Label(dir_f, text="输出目录:").pack(side="left")
        self.var_outdir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        ttk.Entry(dir_f, textvariable=self.var_outdir).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(dir_f, text="浏览", command=self._select_output_dir).pack(side="right")

        name_f = ttk.Frame(f3)
        name_f.pack(fill="x", pady=(0, 8))
        ttk.Label(name_f, text="输出文件名:").pack(side="left")
        self.var_filename = tk.StringVar(value="筛选考勤表.xlsx")
        ttk.Entry(name_f, textvariable=self.var_filename).pack(side="left", fill="x", expand=True, padx=5)

        self.btn_gen = ttk.Button(f3, text="生成考勤表", command=self._generate)
        self.btn_gen.pack(side="left")

        self.progress = ttk.Progressbar(f3, mode="indeterminate", length=200)
        self.progress.pack(side="left", padx=(15, 0), fill="x", expand=True)

        self.var_status = tk.StringVar(value="就绪 — 选择考勤表文件开始")
        self.status_entry = ttk.Entry(self.root, textvariable=self.var_status, state="readonly", font=("微软雅黑", 9))
        self.status_entry.pack(fill="x", padx=10, pady=(5, 10))

    # ── File dialogs ────────────────────────────────────

    def _select_file(self):
        path = filedialog.askopenfilename(
            title="选择总考勤表",
            filetypes=[("Excel文件", "*.xls *.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.source_file = path
            self.var_file.set(path)
            self.var_status.set(f"已选择: {os.path.basename(path)}")

    def _select_output_dir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.var_outdir.set(d)

    def _clear_names(self):
        self.txt_names.delete("1.0", "end")

    # ── Core logic ──────────────────────────────────────

    def _read_source(self) -> list[dict]:
        """Read employees with data from source file via WPS COM."""
        if not win32com:
            raise RuntimeError("缺少 pywin32 模块，请运行: pip install pywin32")

        wps = None
        try:
            wps = win32com.client.Dispatch("Ket.Application")
            wps.Visible = False
            wps.DisplayAlerts = False
            wb = wps.Workbooks.Open(os.path.abspath(self.source_file))
            ws = wb.Sheets(1)
            max_row = ws.UsedRange.Rows.Count
            max_col = ws.UsedRange.Columns.Count

            employees = []
            i = 5
            while i <= max_row:
                cell_a = ws.Cells(i, 1).Value
                if cell_a and "工" in str(cell_a):
                    name = str(ws.Cells(i, 11).Value or "").strip()
                    employees.append({
                        "name": name,
                        "info_row": i,
                        "data_row": i + 1,
                    })
                    i += 2
                else:
                    i += 1

            wb.Close()
            return employees

        finally:
            if wps:
                try:
                    wps.Quit()
                except Exception:
                    pass

    def _read_names(self):
        if not self.source_file:
            messagebox.showwarning("提示", "请先选择总考勤表文件")
            return
        self._set_busy("正在读取总考勤表...")

        def do():
            try:
                employees = self._read_source()
                self.employees = employees
                names = [e["name"] for e in employees if e["name"]]
                self.root.after(0, lambda: self._show_names(names))
            except Exception as ex:
                self.root.after(0, lambda: messagebox.showerror("读取失败", str(ex)))
                self.root.after(0, lambda: self._set_ready("读取失败"))

        threading.Thread(target=do, daemon=True).start()

    def _show_names(self, names: list[str]):
        self.txt_names.delete("1.0", "end")
        self.txt_names.insert("1.0", "\n".join(names))
        self._set_ready(f"读取完成，共 {len(names)} 人")

    def _generate(self):
        if not self.source_file:
            messagebox.showwarning("提示", "请先选择总考勤表文件")
            return

        names_text = self.txt_names.get("1.0", "end").strip()
        target_names = [line.strip() for line in names_text.split("\n") if line.strip()]
        if not target_names:
            messagebox.showwarning("提示", "请输入至少一个姓名（每行一个）")
            return

        output_name = self.var_filename.get().strip()
        if not output_name:
            output_name = "筛选考勤表.xlsx"
        if not output_name.endswith(".xlsx"):
            output_name += ".xlsx"
        output_path = os.path.join(self.var_outdir.get(), output_name)

        self._set_busy("正在处理...")
        self.btn_gen.config(state="disabled")

        def do():
            try:
                self.root.after(0, lambda: self._set_busy("[1/4] 读取总考勤表..."))
                employees = self._read_source()

                self.root.after(0, lambda: self._set_busy(f"[2/4] 匹配 {len(target_names)} 个姓名..."))
                emp_map = {}
                for emp in employees:
                    emp_map.setdefault(emp["name"], emp)

                matched = []
                not_found = []
                for name in target_names:
                    if name in emp_map:
                        matched.append(emp_map[name])
                    else:
                        not_found.append(name)

                if not matched:
                    self.root.after(0, lambda: messagebox.showwarning("未找到", "名单中无匹配人员"))
                    self.root.after(0, lambda: self._set_ready("未找到匹配人员"))
                    return

                self.root.after(0, lambda: self._set_busy(f"[3/4] 复制格式... {len(matched)} 人"))
                self._copy_with_format(matched, output_path)

                status = f"完成 — 匹配 {len(matched)} 人"
                if not_found:
                    status += f"，未找到 {len(not_found)} 人: {'、'.join(not_found)}"
                self.root.after(0, lambda: messagebox.showinfo("完成", f"匹配 {len(matched)} 人，已保存到:\n{output_path}" + (f"\n\n未找到 {len(not_found)} 人:\n{'、'.join(not_found)}" if not_found else "")))
                self.root.after(0, lambda: self._set_ready(status))

            except Exception as ex:
                self.root.after(0, lambda: messagebox.showerror("生成失败", str(ex)))
                self.root.after(0, lambda: self._set_ready("生成失败"))
            finally:
                self.root.after(0, lambda: self.btn_gen.config(state="normal"))

        threading.Thread(target=do, daemon=True).start()

    def _copy_with_format(self, matched: list[dict], output_path: str):
        """Copy rows from source via WPS COM copy-paste, preserving format + column widths."""
        wps = None
        try:
            wps = win32com.client.Dispatch("Ket.Application")
            wps.Visible = False
            wps.DisplayAlerts = False

            src_wb = wps.Workbooks.Open(os.path.abspath(self.source_file))
            src_ws = src_wb.Sheets(1)
            max_col = src_ws.UsedRange.Columns.Count

            new_wb = wps.Workbooks.Add()
            new_ws = new_wb.Sheets(1)
            new_ws.Name = "考勤记录"

            # ── Copy header rows 1-4 ──
            src_ws.Range(src_ws.Cells(1, 1), src_ws.Cells(4, max_col)).Copy()
            new_ws.Range(new_ws.Cells(1, 1), new_ws.Cells(4, max_col)).PasteSpecial(-4104)

            # ── Copy matched employee rows ──
            dest_row = 5
            for emp in matched:
                for src_row in [emp["info_row"], emp["data_row"]]:
                    src_ws.Range(src_ws.Cells(src_row, 1), src_ws.Cells(src_row, max_col)).Copy()
                    new_ws.Range(new_ws.Cells(dest_row, 1), new_ws.Cells(dest_row, max_col)).PasteSpecial(-4104)
                    dest_row += 1

            # ── Keep source column widths (like pressing W) ──
            src_ws.Range(src_ws.Cells(1, 1), src_ws.Cells(src_ws.UsedRange.Rows.Count, max_col)).Copy()
            new_ws.Range(new_ws.Cells(1, 1), new_ws.Cells(dest_row - 1, max_col)).PasteSpecial(8)

            # ── Row heights from source ──
            for r in range(1, 5):
                try:
                    new_ws.Rows(r).RowHeight = src_ws.Rows(r).RowHeight
                except Exception:
                    pass
            d = 5
            for emp in matched:
                for src_row in [emp["info_row"], emp["data_row"]]:
                    try:
                        new_ws.Rows(d).RowHeight = src_ws.Rows(src_row).RowHeight
                    except Exception:
                        pass
                    d += 1

            # Save
            abs_path = os.path.abspath(output_path)
            if os.path.exists(abs_path):
                os.remove(abs_path)
            new_wb.SaveAs(abs_path, 51)

            new_wb.Close()
            src_wb.Close()

        finally:
            if wps:
                try:
                    wps.Quit()
                except Exception:
                    pass

    # ── Helpers ──────────────────────────────────────────

    def _set_busy(self, msg: str):
        self.var_status.set(msg)
        self.progress.start(10)

    def _set_ready(self, msg: str):
        self.progress.stop()
        self.var_status.set(msg)


def main():
    root = tk.Tk()
    AttendanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
