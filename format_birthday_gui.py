"""
Excel 出生年月格式统一工具 - GUI 版本
将各种日期格式统一转换为指定格式
"""

import re
import copy
import threading
from datetime import datetime, date
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


# ─────────────────────────── 默认输入规则 ───────────────────────────

DEFAULT_INPUT_RULES = [
    {
        "name": "YYYY.MM（已是目标格式）",
        "pattern": r"(\d{4})\.(\d{1,2})",
        "fullmatch": True,
        "year_group": 1,
        "month_group": 2,
        "enabled": True,
    },
    {
        "name": "YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD",
        "pattern": r"(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})",
        "fullmatch": True,
        "year_group": 1,
        "month_group": 2,
        "enabled": True,
    },
    {
        "name": "YYYYMMDD（8位）",
        "pattern": r"(\d{4})(\d{2})(\d{2})",
        "fullmatch": True,
        "year_group": 1,
        "month_group": 2,
        "enabled": True,
    },
    {
        "name": "YYYYMM（6位）",
        "pattern": r"(\d{4})(\d{2})",
        "fullmatch": True,
        "year_group": 1,
        "month_group": 2,
        "enabled": True,
    },
]


# ─────────────────────────── 核心转换逻辑 ───────────────────────────

def _apply_out_fmt(year: str, month: str, fmt: str) -> str:
    """根据格式键或自定义模板生成输出字符串"""
    m = f"{int(month):02d}"
    presets = {
        "YYYY.MM": f"{year}.{m}",
        "YYYY-MM": f"{year}-{m}",
        "YYYY/MM": f"{year}/{m}",
        "YYYYMM":  f"{year}{m}",
    }
    if fmt in presets:
        return presets[fmt]
    return fmt.replace("{year}", year).replace("{month}", m)


def normalize_birthday(value, rules=None, out_fmt="YYYY.MM") -> tuple[str, bool]:
    """返回 (转换后字符串, 是否发生了变化)"""
    if rules is None:
        rules = DEFAULT_INPUT_RULES
    if value is None:
        return "", False

    original_text = str(value).strip()

    if isinstance(value, (datetime, date)):
        result = _apply_out_fmt(str(value.year), str(value.month), out_fmt)
        return result, True

    text_clean = re.sub(r"\s+", "", original_text)

    for rule in rules:
        if not rule.get("enabled", True):
            continue
        pattern = rule["pattern"]
        yg = rule["year_group"]
        mg = rule["month_group"]
        try:
            if rule.get("fullmatch", True):
                m = re.fullmatch(pattern, text_clean)
            else:
                m = re.search(pattern, text_clean)
        except re.error:
            continue
        if m:
            result = _apply_out_fmt(m.group(yg), m.group(mg), out_fmt)
            return result, result != original_text

    return original_text, False


def process_excel(input_path: str, col_name: str, output_path: str, log_fn,
                  rules=None, out_fmt="YYYY.MM", header_row: int = 1):
    import openpyxl

    if rules is None:
        rules = DEFAULT_INPUT_RULES

    wb = openpyxl.load_workbook(input_path)
    total_converted = 0

    for sheet in wb.worksheets:
        target_col = None
        for cell in sheet[header_row]:
            if str(cell.value).strip() == col_name:
                target_col = cell.column
                break

        if target_col is None:
            log_fn(f"⚠ Sheet「{sheet.title}」中未找到列「{col_name}」，跳过。\n")
            continue

        converted = 0
        for row in sheet.iter_rows(min_row=header_row + 1):
            cell = row[target_col - 1]
            new_val, changed = normalize_birthday(cell.value, rules, out_fmt)
            if changed:
                log_fn(f"  行 {cell.row}: 「{cell.value}」→「{new_val}」\n")
                converted += 1
            cell.value = new_val
            cell.number_format = "@"

        log_fn(f"✔ Sheet「{sheet.title}」：转换了 {converted} 个单元格\n")
        total_converted += converted

    wb.save(output_path)
    log_fn(f"\n✅ 完成！共转换 {total_converted} 个单元格\n")
    log_fn(f"📄 已保存到：{output_path}\n")


# ─────────────────────────── GUI ───────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 出生年月格式统一工具")
        self.resizable(True, True)
        self.minsize(560, 500)
        self._rules = copy.deepcopy(DEFAULT_INPUT_RULES)
        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = 660, 720
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        pad = dict(padx=12, pady=6)

        # ── 输入文件 ──
        frm_in = ttk.LabelFrame(self, text=" 输入文件 ")
        frm_in.pack(fill="x", **pad)
        self.var_input = tk.StringVar()
        ttk.Entry(frm_in, textvariable=self.var_input, width=55).pack(
            side="left", padx=8, pady=6, fill="x", expand=True)
        ttk.Button(frm_in, text="浏览…", command=self._browse_input).pack(
            side="left", padx=(0, 8), pady=6)

        # ── 输出文件 ──
        frm_out = ttk.LabelFrame(self, text=" 输出文件（留空则自动生成 _fixed.xlsx）")
        frm_out.pack(fill="x", **pad)
        self.var_output = tk.StringVar()
        ttk.Entry(frm_out, textvariable=self.var_output, width=55).pack(
            side="left", padx=8, pady=6, fill="x", expand=True)
        ttk.Button(frm_out, text="浏览…", command=self._browse_output).pack(
            side="left", padx=(0, 8), pady=6)

        # ── 列名 ──
        frm_col = ttk.LabelFrame(self, text=" 要处理的列名 ")
        frm_col.pack(fill="x", **pad)
        self.var_col = tk.StringVar(value="出生年月")
        ttk.Entry(frm_col, textvariable=self.var_col, width=20).pack(
            side="left", padx=8, pady=6)
        ttk.Label(frm_col, text="（与表头完全一致）",
                  foreground="gray").pack(side="left")
        ttk.Label(frm_col, text="  表头行号：").pack(side="left")
        self.var_header_row = tk.StringVar(value="1")
        ttk.Spinbox(frm_col, textvariable=self.var_header_row,
                    from_=1, to=100, width=4).pack(side="left", padx=(0, 8))

        # ── 输出格式 ──
        frm_fmt = ttk.LabelFrame(self, text=" 输出格式 ")
        frm_fmt.pack(fill="x", **pad)
        self.var_fmt = tk.StringVar(value="YYYY.MM")
        self._fmt_combo = ttk.Combobox(
            frm_fmt, textvariable=self.var_fmt,
            values=["YYYY.MM", "YYYY-MM", "YYYY/MM", "YYYYMM", "自定义…"],
            state="readonly", width=14)
        self._fmt_combo.pack(side="left", padx=8, pady=6)
        self._fmt_combo.bind("<<ComboboxSelected>>", self._on_fmt_change)
        ttk.Label(frm_fmt, text="自定义模板（{year} {month}）",
                  foreground="gray").pack(side="left")
        self._custom_fmt_var = tk.StringVar(value="{year}.{month}")
        self._custom_fmt_entry = ttk.Entry(
            frm_fmt, textvariable=self._custom_fmt_var, width=18)
        self._custom_fmt_entry.pack(side="left", padx=(4, 8), pady=6)
        self._custom_fmt_entry.configure(state="disabled")

        # ── 输入规则 ──
        frm_rules = ttk.LabelFrame(self, text=" 输入格式规则（双击行切换启用/禁用）")
        frm_rules.pack(fill="both", expand=False, **pad)
        cols = ("enabled", "name", "pattern")
        self._tree = ttk.Treeview(
            frm_rules, columns=cols, show="headings", height=5)
        self._tree.heading("enabled", text="启用")
        self._tree.heading("name", text="名称")
        self._tree.heading("pattern", text="正则模式")
        self._tree.column("enabled", width=45, anchor="center", stretch=False)
        self._tree.column("name", width=200)
        self._tree.column("pattern", width=300)
        self._tree.pack(fill="both", expand=True, padx=4, pady=(4, 0))
        self._tree.bind("<Double-1>", self._toggle_rule)

        btn_bar = ttk.Frame(frm_rules)
        btn_bar.pack(fill="x", padx=4, pady=4)
        ttk.Button(btn_bar, text="添加", command=self._add_rule).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="编辑", command=self._edit_rule).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="删除", command=self._delete_rule).pack(side="left", padx=2)
        ttk.Button(btn_bar, text="恢复默认", command=self._reset_rules).pack(side="right", padx=2)
        self._refresh_rules_tree()

        # ── 运行按钮 ──
        self.btn_run = ttk.Button(self, text="▶  开始转换", command=self._run)
        self.btn_run.pack(pady=(4, 2))

        # ── 进度条 ──
        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=12, pady=(0, 4))

        # ── 日志 ──
        frm_log = ttk.LabelFrame(self, text=" 运行日志 ")
        frm_log.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        self.log = scrolledtext.ScrolledText(
            frm_log, state="disabled", height=8,
            font=("Consolas", 9), wrap="word")
        self.log.pack(fill="both", expand=True, padx=4, pady=4)

    # ── 输出格式 ──
    def _on_fmt_change(self, _event=None):
        state = "normal" if self.var_fmt.get() == "自定义…" else "disabled"
        self._custom_fmt_entry.configure(state=state)

    def _get_out_fmt(self) -> str:
        v = self.var_fmt.get()
        if v == "自定义…":
            return self._custom_fmt_var.get() or "{year}.{month}"
        return v

    # ── 规则树 ──
    def _refresh_rules_tree(self):
        self._tree.delete(*self._tree.get_children())
        for rule in self._rules:
            mark = "✔" if rule.get("enabled", True) else "✘"
            self._tree.insert("", "end", values=(mark, rule["name"], rule["pattern"]))

    def _toggle_rule(self, event):
        row_id = self._tree.identify_row(event.y)
        if not row_id:
            return
        idx = self._tree.index(row_id)
        self._rules[idx]["enabled"] = not self._rules[idx].get("enabled", True)
        self._refresh_rules_tree()

    def _add_rule(self):
        self._open_rule_dialog(None)

    def _edit_rule(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选择一条规则")
            return
        self._open_rule_dialog(self._tree.index(sel[0]))

    def _delete_rule(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选择一条规则")
            return
        idx = self._tree.index(sel[0])
        if messagebox.askyesno("确认", f"删除规则「{self._rules[idx]['name']}」？"):
            self._rules.pop(idx)
            self._refresh_rules_tree()

    def _reset_rules(self):
        if messagebox.askyesno("确认", "恢复为默认规则？当前修改将丢失。"):
            self._rules = copy.deepcopy(DEFAULT_INPUT_RULES)
            self._refresh_rules_tree()

    # ── 添加/编辑规则对话框 ──
    def _open_rule_dialog(self, idx):
        editing = idx is not None
        rule = copy.deepcopy(self._rules[idx]) if editing else {
            "name": "", "pattern": "", "fullmatch": True,
            "year_group": 1, "month_group": 2, "enabled": True}

        dlg = tk.Toplevel(self)
        dlg.title("编辑规则" if editing else "添加规则")
        dlg.resizable(False, False)
        dlg.grab_set()

        def labeled_entry(label, var, **kw):
            f = ttk.Frame(dlg)
            f.pack(fill="x", padx=12, pady=3)
            ttk.Label(f, text=label, width=12, anchor="e").pack(side="left")
            e = ttk.Entry(f, textvariable=var, **kw)
            e.pack(side="left", fill="x", expand=True, padx=(6, 0))

        v_name = tk.StringVar(value=rule["name"])
        v_pattern = tk.StringVar(value=rule["pattern"])
        v_yg = tk.StringVar(value=str(rule["year_group"]))
        v_mg = tk.StringVar(value=str(rule["month_group"]))
        v_fullmatch = tk.BooleanVar(value=rule.get("fullmatch", True))

        labeled_entry("名称", v_name)
        labeled_entry("正则模式", v_pattern)
        labeled_entry("年份组号", v_yg, width=6)
        labeled_entry("月份组号", v_mg, width=6)

        f_fm = ttk.Frame(dlg)
        f_fm.pack(fill="x", padx=12, pady=3)
        ttk.Label(f_fm, text="全匹配", width=12, anchor="e").pack(side="left")
        ttk.Checkbutton(f_fm, variable=v_fullmatch).pack(side="left", padx=(6, 0))

        ttk.Separator(dlg, orient="horizontal").pack(fill="x", padx=12, pady=6)
        ttk.Label(dlg, text="测试输入：").pack(anchor="w", padx=12)
        v_test = tk.StringVar()
        v_result = tk.StringVar(value="—")
        ttk.Entry(dlg, textvariable=v_test).pack(fill="x", padx=12, pady=2)
        lbl_result = ttk.Label(dlg, textvariable=v_result, foreground="blue")
        lbl_result.pack(anchor="w", padx=12, pady=(0, 6))

        def _live_test(*_):
            sample = re.sub(r"\s+", "", v_test.get())
            pat = v_pattern.get()
            try:
                yg, mg = int(v_yg.get() or 1), int(v_mg.get() or 2)
                fn = re.fullmatch if v_fullmatch.get() else re.search
                m = fn(pat, sample)
                if m:
                    v_result.set(f"✔ 年={m.group(yg)}  月={m.group(mg)}")
                    lbl_result.configure(foreground="green")
                else:
                    v_result.set("✘ 不匹配")
                    lbl_result.configure(foreground="red")
            except Exception as e:
                v_result.set(f"错误: {e}")
                lbl_result.configure(foreground="orange")

        for var in (v_test, v_pattern, v_yg, v_mg, v_fullmatch):
            var.trace_add("write", _live_test)

        def _save():
            name = v_name.get().strip()
            pattern = v_pattern.get().strip()
            if not name or not pattern:
                messagebox.showwarning("提示", "名称和正则模式不能为空", parent=dlg)
                return
            try:
                re.compile(pattern)
            except re.error as e:
                messagebox.showerror("正则错误", str(e), parent=dlg)
                return
            new_rule = {
                "name": name, "pattern": pattern,
                "fullmatch": v_fullmatch.get(),
                "year_group": int(v_yg.get() or 1),
                "month_group": int(v_mg.get() or 2),
                "enabled": rule.get("enabled", True),
            }
            if editing:
                self._rules[idx] = new_rule
            else:
                self._rules.append(new_rule)
            self._refresh_rules_tree()
            dlg.destroy()

        btn_row = ttk.Frame(dlg)
        btn_row.pack(pady=8)
        ttk.Button(btn_row, text="保存", command=_save).pack(side="left", padx=6)
        ttk.Button(btn_row, text="取消", command=dlg.destroy).pack(side="left", padx=6)

        dlg.update_idletasks()
        dw, dh = 420, dlg.winfo_reqheight()
        px, py = self.winfo_x(), self.winfo_y()
        pw, ph = self.winfo_width(), self.winfo_height()
        dlg.geometry(f"{dw}x{dh}+{px+(pw-dw)//2}+{py+(ph-dh)//2}")

    # ── 文件对话框 ──
    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls"), ("所有文件", "*.*")])
        if path:
            self.var_input.set(path)
            if not self.var_output.get():
                p = Path(path)
                self.var_output.set(str(p.with_name(p.stem + "_fixed" + p.suffix)))

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="保存为", defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")])
        if path:
            self.var_output.set(path)

    # ── 日志输出 ──
    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg)
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── 运行 ──
    def _run(self):
        inp = self.var_input.get().strip()
        col = self.var_col.get().strip()
        out = self.var_output.get().strip()

        if not inp:
            messagebox.showwarning("提示", "请先选择输入文件！")
            return
        if not Path(inp).exists():
            messagebox.showerror("错误", f"文件不存在：\n{inp}")
            return
        if not col:
            messagebox.showwarning("提示", "请输入列名！")
            return
        if not out:
            p = Path(inp)
            out = str(p.with_name(p.stem + "_fixed" + p.suffix))
            self.var_output.set(out)

        rules = copy.deepcopy(self._rules)
        out_fmt = self._get_out_fmt()
        try:
            header_row = max(1, int(self.var_header_row.get()))
        except ValueError:
            header_row = 1

        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.btn_run.configure(state="disabled")
        self.progress.start(10)

        def worker():
            try:
                process_excel(inp, col, out, self._log, rules, out_fmt, header_row)
                self.after(0, lambda: messagebox.showinfo(
                    "完成", f"转换成功！\n已保存到：\n{out}"))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("错误", str(e)))
                self._log(f"\n❌ 错误：{e}\n")
            finally:
                self.after(0, self._done)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self):
        self.progress.stop()
        self.btn_run.configure(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()
