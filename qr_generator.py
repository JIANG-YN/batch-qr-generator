#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Windows 批量二维码生成工具：读取 Excel/CSV 名单，为每行生成一张 PNG 二维码。"""
import csv
import hashlib
import os
import re
import shutil
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from openpyxl import load_workbook
    import qrcode
    from PIL import Image
except ImportError:
    print("缺少依赖库，请运行：pip install -r requirements.txt")
    sys.exit(1)

APP_NAME = "批量二维码生成工具"
REQUIRED_COLUMNS = ["CodeID", "Content"]
ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]+$")
BATCH_PREFIX = "qr_codes_"
TIME_FMT = "%Y%m%d_%H%M%S"
MAX_ROWS = 1000
PREVIEW_ROWS = 10
QR_SIZES = [300, 500, 800]
QR_ERROR_LEVEL = qrcode.constants.ERROR_CORRECT_M
MANIFEST_NAME = "qr_manifest.csv"

def file_fingerprint(path):
    """返回 (size, mtime, sha256) 用于检测文件是否被修改。"""
    p = Path(path)
    if not p.exists():
        return None
    stat = p.stat()
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return (stat.st_size, int(stat.st_mtime), h.hexdigest())

def load_table(path):
    """读取名单，返回 (rows, error)。rows 每项是 {列名: 原始值}。"""
    p = Path(path)
    if not p.exists():
        return None, "文件不存在或已被移动/删除"
    suffix = p.suffix.lower()
    try:
        if suffix == ".xlsx":
            wb = load_workbook(p, data_only=True, read_only=True)
            try:
                ws = wb.worksheets[0]
            except IndexError:
                wb.close()
                return None, "文件没有工作表"
            rows_iter = ws.iter_rows(values_only=True)
            try:
                header = next(rows_iter)
            except StopIteration:
                wb.close()
                return None, "文件为空，没有表头"
            if header is None:
                wb.close()
                return None, "文件为空，没有表头"
            header = [str(c).strip() if c is not None else "" for c in header]
            missing = [c for c in REQUIRED_COLUMNS if c not in header]
            if missing:
                wb.close()
                return None, f"名单缺少必须列：{', '.join(missing)}"
            col_idx = {c: header.index(c) for c in REQUIRED_COLUMNS}
            rows = []
            for row in rows_iter:
                if row is None:
                    continue
                rows.append({col: (row[idx] if idx < len(row) else None)
                             for col, idx in col_idx.items()})
            wb.close()
            return rows, None
        elif suffix == ".csv":
            try:
                with open(p, "r", encoding="utf-8-sig", newline="") as f:
                    reader = csv.reader(f)
                    try:
                        header = next(reader)
                    except StopIteration:
                        return None, "文件为空，没有表头"
                    if header is None:
                        return None, "文件为空，没有表头"
                    header = [c.strip() for c in header]
                    missing = [c for c in REQUIRED_COLUMNS if c not in header]
                    if missing:
                        return None, f"名单缺少必须列：{', '.join(missing)}"
                    col_idx = {c: header.index(c) for c in REQUIRED_COLUMNS}
                    rows = []
                    for raw in reader:
                        rows.append({col: (raw[idx] if idx < len(raw) else None)
                                     for col, idx in col_idx.items()})
                    return rows, None
            except UnicodeDecodeError:
                return None, "CSV 编码不是 UTF-8/UTF-8 BOM，请用 Excel 另存为 UTF-8 CSV"
        else:
            return None, "仅支持 .xlsx 和 .csv 文件"
    except Exception as e:
        return None, f"文件读取失败：{e}"

def validate_table(rows, file_label):
    """校验名单，返回 (clean_rows, errors)。clean_rows 每项是 {CodeID, Content}。"""
    errors = []
    clean = []
    seen_ids = {}
    for excel_row, row in enumerate(rows, start=2):
        raw_id = row.get("CodeID")
        raw_content = row.get("Content")
        sid = "" if raw_id is None else str(raw_id).strip()
        content = "" if raw_content is None else str(raw_content)
        # 整行空：CodeID 和 Content 都为空 → 跳过
        if sid == "" and (raw_content is None or str(raw_content).strip() == ""):
            continue
        # 任一字段空 → 报错
        if sid == "":
            errors.append(f"{file_label}第{excel_row}行CodeID为空")
            continue
        if raw_content is None or str(raw_content).strip() == "":
            errors.append(f"{file_label}第{excel_row}行Content为空")
            continue
        # 编号格式
        if not ID_PATTERN.match(sid):
            errors.append(
                f"{file_label}第{excel_row}行CodeID包含非法字符：{sid}"
                f"（只允许字母、数字、下划线、短横线）"
            )
            continue
        # 判重（忽略大小写）
        key = sid.lower()
        if key in seen_ids:
            errors.append(
                f"{file_label}第{excel_row}行CodeID重复：{sid}"
                f"（与第{seen_ids[key]}行重复）"
            )
            continue
        seen_ids[key] = excel_row
        # Content 保持原值，仅判空时 strip 过
        clean.append({"CodeID": sid, "Content": str(raw_content)})
    if not clean and not errors:
        errors.append("名单没有有效数据（只有表头或全部为空白行）")
        return None, errors
    if len(clean) > MAX_ROWS:
        errors.append(f"名单有效行数为{len(clean)}，超过最大支持行数{MAX_ROWS}")
        return None, errors
    return clean, errors

def render_qr(content, target_size):
    """生成边长精确等于 target_size 的 PNG 二维码，白色背景，保留≥4模块静区。
    使用整数 box_size 保证模块边缘清晰；如需缩放只用 NEAREST。
    """
    qr = qrcode.QRCode(
        version=None,
        error_correction=QR_ERROR_LEVEL,
        box_size=10,
        border=4,
    )
    qr.add_data(content)
    qr.make(fit=True)
    modules = qr.modules_count + 2 * qr.border
    # 计算 box_size：使二维码边长尽量接近 target_size 但不超过
    box = max(1, target_size // modules)
    qr.box_size = box
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    actual = qr_img.size[0]  # modules * box
    if actual == target_size:
        return qr_img
    # 居中放到 target_size × target_size 白色画布
    canvas = Image.new("RGB", (target_size, target_size), (255, 255, 255))
    offset = ((target_size - actual) // 2, (target_size - actual) // 2)
    canvas.paste(qr_img, offset)
    return canvas

def resolve_batch_dir(out_dir):
    """生成新的批次目录，已存在追加序号。返回 Path。"""
    out_dir = Path(out_dir)
    timestamp = datetime.now().strftime(TIME_FMT)
    base = f"{BATCH_PREFIX}{timestamp}"
    candidate = out_dir / base
    if not candidate.exists():
        candidate.mkdir(parents=True)
        return candidate
    k = 1
    while True:
        candidate = out_dir / f"{base}_{k}"
        if not candidate.exists():
            candidate.mkdir(parents=True)
            return candidate
        k += 1


def _cleanup_dir(path):
    """递归删除目录内全部内容，再删除目录本身。用于整批失败回滚。"""
    shutil.rmtree(path, ignore_errors=False)


class QRGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("960x780")
        self.root.minsize(880, 700)

        self.file_path = None
        self.out_dir = None
        self.clean_rows = None
        self.roster_rows = None
        self.file_valid = False
        self.file_fp = None
        self.generating = False

        self.file_var = tk.StringVar(value="（未选择文件）")
        self.count_var = tk.StringVar(value="有效行数：—")
        self.out_var = tk.StringVar(value="（未选择输出文件夹）")
        self.size_var = tk.IntVar(value=500)
        self.status_var = tk.StringVar(value="请选择 Excel/CSV 文件")

        self._build_ui()
        self._refresh_button_state()

    def _build_ui(self):
        step1 = ttk.LabelFrame(self.root, text="第 1 步：选择 Excel/CSV 文件")
        step1.pack(fill="x", padx=10, pady=6)
        f1 = ttk.Frame(step1)
        f1.pack(fill="x", padx=10, pady=(8, 2))
        ttk.Button(f1, text="选择文件…", command=self.on_select_file).pack(side="left")
        ttk.Label(f1, textvariable=self.file_var).pack(side="left", padx=12)
        ttk.Label(f1, textvariable=self.count_var, foreground="#0a7").pack(side="right")

        preview_frame = ttk.LabelFrame(step1, text=f"名单预览（前 {PREVIEW_ROWS} 行）")
        preview_frame.pack(fill="both", expand=True, padx=10, pady=(2, 8))
        self.preview_text = tk.Text(preview_frame, height=8, wrap="none",
                                   font=("Consolas", 10))
        self.preview_text.pack(fill="both", expand=True, padx=6, pady=6)
        self.preview_text.config(state="disabled")

        step2 = ttk.LabelFrame(self.root, text="第 2 步：选择输出文件夹")
        step2.pack(fill="x", padx=10, pady=6)
        f2 = ttk.Frame(step2)
        f2.pack(fill="x", padx=10, pady=8)
        ttk.Button(f2, text="选择输出文件夹…", command=self.on_select_out).pack(side="left")
        ttk.Label(f2, textvariable=self.out_var).pack(side="left", padx=12)

        step3 = ttk.LabelFrame(self.root, text="第 3 步：设置二维码尺寸并生成")
        step3.pack(fill="x", padx=10, pady=6)
        f3 = ttk.Frame(step3)
        f3.pack(fill="x", padx=10, pady=8)
        for size in QR_SIZES:
            ttk.Radiobutton(f3, text=f"{size}px", value=size,
                            variable=self.size_var).pack(side="left", padx=8)
        self.btn_generate = ttk.Button(f3, text="开始生成", command=self.on_generate)
        self.btn_generate.pack(side="left", padx=12)
        ttk.Label(f3, textvariable=self.status_var, foreground="#06c").pack(side="left", padx=12)
        self.progress = ttk.Progressbar(step3, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(0, 8))

        step4 = ttk.LabelFrame(self.root, text="第 4 步：生成结果")
        step4.pack(fill="both", expand=True, padx=10, pady=6)
        rf = ttk.Frame(step4)
        rf.pack(fill="both", expand=True, padx=10, pady=10)
        self.result_text = tk.Text(rf, height=12, wrap="word", font=("Consolas", 10))
        self.result_text.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(rf, orient="vertical", command=self.result_text.yview)
        sb.pack(side="right", fill="y")
        self.result_text.config(yscrollcommand=sb.set, state="disabled")

    def _refresh_button_state(self):
        ready = (self.file_valid and self.out_dir is not None
                 and not self.generating)
        self.btn_generate.config(state="normal" if ready else "disabled")

    def _update_preview(self):
        self.preview_text.config(state="normal")
        self.preview_text.delete("1.0", "end")
        if not self.roster_rows:
            self.preview_text.config(state="disabled")
            return
        header = REQUIRED_COLUMNS
        col_widths = [len(h) for h in header]
        display_rows = []
        for row in self.roster_rows[:PREVIEW_ROWS]:
            vals = []
            for col in header:
                v = row.get(col)
                s = "" if v is None else str(v).strip()
                vals.append(s)
                i = header.index(col)
                col_widths[i] = max(col_widths[i], len(s))
            display_rows.append(vals)
        if len(self.roster_rows) > PREVIEW_ROWS:
            display_rows.append(["…"] * len(header))

        def fmt(vals):
            return " | ".join(v.ljust(w) for v, w in zip(vals, col_widths))
        self.preview_text.insert("end", fmt(header) + "\n")
        self.preview_text.insert("end", "-+-".join("-" * w for w in col_widths) + "\n")
        for vals in display_rows:
            self.preview_text.insert("end", fmt(vals) + "\n")
        self.preview_text.config(state="disabled")

    def _append_result(self, text):
        self.result_text.config(state="normal")
        self.result_text.insert("end", text + "\n")
        self.result_text.config(state="disabled")
        self.result_text.see("end")

    def on_select_file(self):
        path = filedialog.askopenfilename(
            title="选择名单文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv"),
                       ("所有文件", "*.*")]
        )
        if not path:
            return
        rows, err = load_table(path)
        if err:
            self.file_valid = False
            self.roster_rows = None
            self.clean_rows = None
            self.file_var.set(f"（无效：{Path(path).name}）")
            self.count_var.set("有效行数：—")
            self._update_preview()
            self._refresh_button_state()
            messagebox.showerror(APP_NAME,
                                 f"文件读取失败：\n\n文件：{Path(path).name}\n原因：{err}")
            return
        clean, errs = validate_table(rows, Path(path).name)
        if errs:
            self.file_valid = False
            self.roster_rows = rows
            self.clean_rows = None
            self.file_var.set(f"（无效：{Path(path).name}）")
            self.count_var.set("有效行数：—")
            self._update_preview()
            self._refresh_button_state()
            msg = "名单校验失败：\n\n" + "\n".join(f"- {e}" for e in errs[:20])
            if len(errs) > 20:
                msg += f"\n…（共 {len(errs)} 条错误）"
            messagebox.showerror(APP_NAME, msg)
            return
        self.file_path = path
        self.file_valid = True
        self.roster_rows = rows
        self.clean_rows = clean
        self.file_fp = file_fingerprint(path)
        self.file_var.set(Path(path).name)
        self.count_var.set(f"有效行数：{len(clean)}")
        self._update_preview()
        self._refresh_button_state()
        self.status_var.set("文件已就绪，请选择输出文件夹")

    def on_select_out(self):
        path = filedialog.askdirectory(title="选择输出文件夹")
        if not path:
            return
        self.out_dir = path
        self.out_var.set(path)
        self._refresh_button_state()
        if self.file_valid:
            self.status_var.set("已就绪，可以开始生成")

    def on_generate(self):
        if not (self.file_valid and self.out_dir):
            messagebox.showwarning(APP_NAME, "请先选择文件和输出文件夹")
            return
        # 生成前重新校验：文件大小/修改时间/哈希
        new_fp = file_fingerprint(self.file_path)
        if new_fp is None:
            self.file_valid = False
            self._refresh_button_state()
            messagebox.showerror(APP_NAME,
                                 f"文件已被移动或删除：{Path(self.file_path).name}")
            self.status_var.set("文件不存在，请重新选择")
            return
        if new_fp != self.file_fp:
            # 指纹不同：立即停止，不论能否通过校验都不自动接受新内容
            self.file_valid = False
            self.clean_rows = None
            self.file_fp = None
            self.file_var.set(f"（已变化：{Path(self.file_path).name}）")
            self.count_var.set("有效行数：—")
            self._update_preview()
            self._refresh_button_state()
            messagebox.showerror(
                APP_NAME,
                f"输入文件已修改：{Path(self.file_path).name}\n\n"
                f"为防止使用未经预览确认的数据，已停止生成。\n"
                f"请重新选择并确认名单。"
            )
            self.status_var.set("文件已修改，请重新选择")
            return

        count = len(self.clean_rows)
        size = self.size_var.get()
        if not messagebox.askyesno(APP_NAME,
                                   f"即将为 {count} 行数据生成 {size}×{size} 二维码，是否继续？"):
            self.status_var.set("已取消")
            return

        self.generating = True
        self._refresh_button_state()
        self.progress["maximum"] = count
        self.progress["value"] = 0
        self.status_var.set("正在生成…")
        self.root.update_idletasks()
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.config(state="disabled")

        # 临时目录方案：全部成功后再改名为正式目录，避免部分结果冒充成功
        out_dir = Path(self.out_dir)
        timestamp = datetime.now().strftime(TIME_FMT)
        base_name = f"{BATCH_PREFIX}{timestamp}"
        final_dir = out_dir / base_name
        if final_dir.exists():
            k = 1
            while True:
                candidate = out_dir / f"{base_name}_{k}"
                if not candidate.exists():
                    final_dir = candidate
                    break
                k += 1
        tmp_dir = out_dir / f"{base_name}.tmp"
        if tmp_dir.exists():
            _cleanup_dir(tmp_dir)
        tmp_dir.mkdir(parents=True)

        try:
            # 逐张生成：任一张失败立即终止整批
            for i, data in enumerate(self.clean_rows, start=1):
                cid = data["CodeID"]
                content = data["Content"]
                fname = f"{cid}.png"
                out_path = tmp_dir / fname
                img = render_qr(content, size)
                img.save(str(out_path), "PNG")
                self._append_result(f"  ✓ {fname}")
                self.progress["value"] = i
                self.status_var.set(f"正在生成…（{i}/{count}）")
                self.root.update_idletasks()

            # 全部二维码生成成功，写 manifest：失败也视为整批失败
            manifest_path = tmp_dir / MANIFEST_NAME
            with open(manifest_path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=["CodeID", "Content", "FileName", "Status"]
                )
                writer.writeheader()
                for data in self.clean_rows:
                    writer.writerow({
                        "CodeID": data["CodeID"],
                        "Content": data["Content"],
                        "FileName": f"{data['CodeID']}.png",
                        "Status": "Success",
                    })

            # 全部成功，改名为正式目录
            tmp_dir.rename(final_dir)

            self._append_result("")
            self._append_result(f"成功：{count} 张")
            self._append_result(f"输出目录：{final_dir}")
            self._append_result(f"清单文件：{final_dir / MANIFEST_NAME}")
            self.status_var.set("生成完成")
            messagebox.showinfo(APP_NAME,
                                f"生成完成：成功 {count} 张。\n\n输出目录：{final_dir}")
        except Exception as e:
            # 整批失败：清理临时目录的全部结果，不保留任何文件
            self._append_result(f"  ✗ 生成失败：{e}")
            if tmp_dir.exists():
                try:
                    _cleanup_dir(tmp_dir)
                    self._append_result("已清理本批次全部结果。")
                except Exception as ce:
                    self._append_result(f"  ✗ 清理失败：{ce}")
            self._append_result("")
            self._append_result("生成失败，未保留任何文件。")
            self.status_var.set("生成失败")
            messagebox.showerror(APP_NAME,
                                f"生成失败，未保留任何文件。\n\n原因：{e}")
        finally:
            self.generating = False
            self._refresh_button_state()

def main():
    root = tk.Tk()
    QRGenerator(root)
    root.mainloop()

if __name__ == "__main__":
    main()