from __future__ import annotations

import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from wpsfix import ConversionReport, convert_docx


APP_TITLE = "WPS 文档兼容修复工具"
VERSION = "V1.0.1"
AUTHOR = "鼎泰高科全球信息部"
CREATOR = "James"


def resource_path(relative_path: str) -> Path:
    """Resolve bundled resources both from source and a PyInstaller EXE."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative_path


class WpsFixApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_TITLE} · {CREATOR}.{VERSION}")
        self.geometry("860x590")
        self.minsize(760, 540)
        self.configure(bg="#EDF2F7")
        try:
            self.iconbitmap(str(resource_path("assets/wpsfix.ico")))
        except tk.TclError:
            pass

        self.selected_file: Path | None = None
        self.events: queue.Queue[tuple] = queue.Queue()
        self.path_text = tk.StringVar(value="尚未选择 DOCX 文件")
        self.status_text = tk.StringVar(value="请选择需要检查和转换的 DOCX 文件。")
        self.progress_value = tk.DoubleVar(value=0)
        self.progress_label = tk.StringVar(value="0%")

        self._configure_style()
        self._build_ui()
        self.after(100, self._drain_events)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Brand.Horizontal.TProgressbar",
            troughcolor="#E3EAF2", background="#FF6B3D", bordercolor="#E3EAF2",
            lightcolor="#FF6B3D", darkcolor="#FF6B3D", thickness=9,
        )

    def _build_ui(self) -> None:
        shell = tk.Frame(self, bg="#EDF2F7")
        shell.pack(fill="both", expand=True)

        hero = tk.Frame(shell, bg="#102A43", height=142)
        hero.pack(fill="x")
        hero.pack_propagate(False)
        hero_inner = tk.Frame(hero, bg="#102A43")
        hero_inner.pack(fill="both", expand=True, padx=38, pady=24)

        brand = tk.Frame(hero_inner, bg="#102A43")
        brand.pack(side="left", fill="y")
        mark = tk.Label(
            brand, text="WPS\nFIX", bg="#FF6B3D", fg="#FFFFFF",
            font=("Microsoft YaHei UI", 10, "bold"), width=7, height=3,
            justify="center",
        )
        mark.pack(side="left", padx=(0, 18))
        heading = tk.Frame(brand, bg="#102A43")
        heading.pack(side="left", anchor="center")
        tk.Label(
            heading, text=APP_TITLE, bg="#102A43", fg="#FFFFFF",
            font=("Microsoft YaHei UI", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            heading, text="自动修复页码域与外链图片 · 保留原文件",
            bg="#102A43", fg="#B8C7D9", font=("Microsoft YaHei UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        meta = tk.Frame(hero_inner, bg="#102A43")
        meta.pack(side="right", anchor="ne")
        tk.Label(
            meta, text=VERSION, bg="#1D3B57", fg="#FFB399",
            font=("Microsoft YaHei UI", 9, "bold"), padx=11, pady=5,
        ).pack(anchor="e")
        tk.Label(
            meta, text=AUTHOR, bg="#102A43", fg="#D7E1EC",
            font=("Microsoft YaHei UI", 9),
        ).pack(anchor="e", pady=(9, 0))

        body = tk.Frame(shell, bg="#EDF2F7")
        body.pack(fill="both", expand=True, padx=38, pady=(22, 14))
        card = tk.Frame(body, bg="#FFFFFF", highlightthickness=1, highlightbackground="#DCE5EE")
        card.pack(fill="both", expand=True)
        content = tk.Frame(card, bg="#FFFFFF")
        content.pack(fill="both", expand=True, padx=26, pady=22)

        tk.Label(
            content, text="01  选择文档", bg="#FFFFFF", fg="#17324D",
            font=("Microsoft YaHei UI", 11, "bold"),
        ).pack(anchor="w")
        path_box = tk.Frame(content, bg="#F4F7FA", highlightthickness=1, highlightbackground="#E1E8F0")
        path_box.pack(fill="x", pady=(9, 14))
        tk.Label(
            path_box, text="DOCX", bg="#DCE8F5", fg="#1C4F7C",
            font=("Microsoft YaHei UI", 8, "bold"), padx=9, pady=5,
        ).pack(side="left", padx=10, pady=9)
        path_label = tk.Label(
            path_box, textvariable=self.path_text, bg="#F4F7FA", fg="#40566F",
            font=("Microsoft YaHei UI", 9), anchor="w", justify="left",
        )
        path_label.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=9)

        button_row = tk.Frame(content, bg="#FFFFFF")
        button_row.pack(anchor="w", pady=(0, 20))
        self.select_button = tk.Button(
            button_row, text="选择 DOCX 文件", command=self._select_file,
            bg="#E8F0F8", fg="#194A73", activebackground="#D8E6F3",
            activeforeground="#123D61", relief="flat", bd=0, cursor="hand2",
            font=("Microsoft YaHei UI", 10, "bold"), padx=23, pady=10,
        )
        self.select_button.pack(side="left", padx=(0, 12))
        self.convert_button = tk.Button(
            button_row, text="检查并转换", command=self._start_conversion, state="disabled",
            bg="#B8C3CF", fg="#FFFFFF", activebackground="#B8C3CF",
            activeforeground="#FFFFFF", disabledforeground="#F4F7FA",
            relief="flat", bd=0, cursor="arrow",
            font=("Microsoft YaHei UI", 10, "bold"), padx=28, pady=10,
        )
        self.convert_button.pack(side="left")

        tk.Frame(content, bg="#E7EDF3", height=1).pack(fill="x", pady=(0, 16))
        tk.Label(
            content, text="02  处理进度", bg="#FFFFFF", fg="#17324D",
            font=("Microsoft YaHei UI", 11, "bold"),
        ).pack(anchor="w")
        progress_row = tk.Frame(content, bg="#FFFFFF")
        progress_row.pack(fill="x", pady=(10, 0))
        self.progress = ttk.Progressbar(
            progress_row, variable=self.progress_value, maximum=100,
            style="Brand.Horizontal.TProgressbar",
        )
        self.progress.pack(side="left", fill="x", expand=True)
        tk.Label(
            progress_row, textvariable=self.progress_label, bg="#FFFFFF", fg="#FF6B3D",
            font=("Microsoft YaHei UI", 9, "bold"), width=5, anchor="e",
        ).pack(side="left", padx=(12, 0))
        status_row = tk.Frame(content, bg="#FFFFFF")
        status_row.pack(fill="x", pady=(8, 15))
        tk.Label(status_row, text="●", bg="#FFFFFF", fg="#5A9A72", font=("Microsoft YaHei UI", 8)).pack(side="left")
        tk.Label(
            status_row, textvariable=self.status_text, bg="#FFFFFF", fg="#51667D",
            font=("Microsoft YaHei UI", 9),
        ).pack(side="left", padx=(6, 0))

        tk.Label(
            content, text="03  转换结果", bg="#FFFFFF", fg="#17324D",
            font=("Microsoft YaHei UI", 11, "bold"),
        ).pack(anchor="w")
        self.result_section = tk.Frame(
            content, bg="#FFFFFF", highlightthickness=2, highlightbackground="#DFE7EF"
        )
        self.result_section.pack(fill="both", expand=True, pady=(9, 0))
        self.result = tk.Text(
            self.result_section, height=5, wrap="word", borderwidth=0, highlightthickness=0,
            highlightbackground="#DFE7EF", background="#F7F9FC", foreground="#40566F",
            selectbackground="#CFE0F2", font=("Microsoft YaHei UI", 9), padx=12, pady=9,
        )
        self.result.pack(fill="both", expand=True)
        self._set_result("等待转换。选择文档后，程序将在原文件旁生成新的 WPS 兼容版本。")

        tk.Label(
            shell, text=f"{VERSION}  ·  {AUTHOR}", bg="#EDF2F7", fg="#71839A",
            font=("Microsoft YaHei UI", 8),
        ).pack(pady=(0, 10))

    def _set_result(self, text: str) -> None:
        self.result.configure(state="normal")
        self.result.delete("1.0", "end")
        self.result.insert("1.0", text)
        self.result.configure(state="disabled")

    def _set_convert_enabled(self, enabled: bool) -> None:
        if enabled:
            self.convert_button.configure(
                state="normal", bg="#FF6B3D", activebackground="#E95A2E",
                fg="#FFFFFF", cursor="hand2",
            )
        else:
            self.convert_button.configure(
                state="disabled", bg="#B8C3CF", activebackground="#B8C3CF",
                fg="#FFFFFF", cursor="arrow",
            )

    def _fit_window_to_result(self) -> None:
        """Resize for all rendered result lines, including wrapped long paths."""
        self.update_idletasks()
        display_line_count = self.result.count("1.0", "end", "displaylines")
        display_lines = display_line_count[0] if display_line_count else 1
        self.result.configure(height=max(5, display_lines + 1))
        self.update_idletasks()

        current_width = max(self.winfo_width(), 860)
        target_height = max(self.winfo_height(), self.winfo_reqheight() + 8, 590)

        # Keep the complete result inside the usable screen area. If necessary,
        # move the window upward before limiting its height.
        screen_height = self.winfo_screenheight()
        max_height = max(590, screen_height - 80)
        target_height = min(target_height, max_height)
        x = max(0, self.winfo_x())
        y = max(0, min(self.winfo_y(), screen_height - target_height - 60))
        self.geometry(f"{current_width}x{target_height}+{x}+{y}")
        self.update_idletasks()

    def _reveal_result_area(self, accent: str = "#FF6B3D", fit_content: bool = False) -> None:
        """Keep the live conversion result visible after the primary action."""
        self.deiconify()
        width = max(self.winfo_width(), 860)
        height = max(self.winfo_height(), 590)
        self.geometry(f"{width}x{height}")
        self.result_section.configure(highlightbackground=accent)
        if fit_content:
            self._fit_window_to_result()
        self.result.see("1.0")
        self.update_idletasks()

    def _show_notice(self, kind: str, heading: str, description: str, detail: str) -> None:
        """Show a brand-styled modal instead of a platform message box."""
        success = kind == "success"
        accent = "#5A9A72" if success else "#C84A43"
        dialog = tk.Toplevel(self)
        dialog.title(f"{heading} · {APP_TITLE}")
        dialog.configure(bg="#FFFFFF")
        dialog.resizable(False, False)
        dialog.transient(self)
        try:
            dialog.iconbitmap(str(resource_path("assets/wpsfix.ico")))
        except tk.TclError:
            pass

        header = tk.Frame(dialog, bg="#102A43", height=68)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header, text="WPS\nFIX", bg="#FF6B3D", fg="#FFFFFF",
            font=("Microsoft YaHei UI", 8, "bold"), width=6, height=3,
        ).pack(side="left", padx=(24, 14), pady=10)
        tk.Label(
            header, text=APP_TITLE, bg="#102A43", fg="#FFFFFF",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(side="left")
        tk.Label(
            header, text=VERSION, bg="#102A43", fg="#B8C7D9",
            font=("Microsoft YaHei UI", 9, "bold"),
        ).pack(side="right", padx=24)

        body = tk.Frame(dialog, bg="#FFFFFF")
        body.pack(fill="both", expand=True, padx=30, pady=24)
        state_icon = tk.Canvas(body, width=58, height=58, bg="#FFFFFF", highlightthickness=0)
        state_icon.pack(side="left", anchor="n", padx=(0, 20))
        state_icon.create_oval(3, 3, 55, 55, fill=accent, outline=accent)
        if success:
            state_icon.create_line(16, 30, 26, 40, fill="#FFFFFF", width=5, capstyle="round")
            state_icon.create_line(25, 40, 43, 19, fill="#FFFFFF", width=5, capstyle="round")
        else:
            state_icon.create_line(29, 16, 29, 34, fill="#FFFFFF", width=5, capstyle="round")
            state_icon.create_oval(26, 41, 32, 47, fill="#FFFFFF", outline="#FFFFFF")

        copy = tk.Frame(body, bg="#FFFFFF")
        copy.pack(side="left", fill="both", expand=True)
        tk.Label(
            copy, text=heading, bg="#FFFFFF", fg="#17324D",
            font=("Microsoft YaHei UI", 14, "bold"),
        ).pack(anchor="w")
        tk.Label(
            copy, text=description, bg="#FFFFFF", fg="#51667D",
            font=("Microsoft YaHei UI", 9),
        ).pack(anchor="w", pady=(5, 14))
        detail_box = tk.Label(
            copy, text=detail, bg="#F4F7FA", fg="#40566F",
            font=("Microsoft YaHei UI", 9), justify="left", anchor="w",
            wraplength=500, padx=14, pady=11,
            highlightthickness=1, highlightbackground="#DFE7EF",
        )
        detail_box.pack(fill="x")

        footer = tk.Frame(dialog, bg="#EDF2F7", height=68)
        footer.pack(fill="x")
        footer.pack_propagate(False)
        tk.Label(
            footer, text=AUTHOR, bg="#EDF2F7", fg="#71839A",
            font=("Microsoft YaHei UI", 8),
        ).pack(side="left", padx=26)
        confirm_button = tk.Button(
            footer, text="确定", command=dialog.destroy,
            bg="#FF6B3D", fg="#FFFFFF", activebackground="#E95A2E",
            activeforeground="#FFFFFF", relief="flat", bd=0, cursor="hand2",
            font=("Microsoft YaHei UI", 10, "bold"), padx=30, pady=8,
        )
        confirm_button.pack(side="right", padx=26, pady=14)

        dialog.update_idletasks()
        width = max(660, dialog.winfo_reqwidth())
        height = max(300, dialog.winfo_reqheight())
        x = self.winfo_rootx() + max(0, (self.winfo_width() - width) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - height) // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.grab_set()
        confirm_button.focus_set()
        dialog.wait_window()

    def _select_file(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self,
            title="选择需要转换的 DOCX 文件",
            filetypes=[("Word 文档", "*.docx")],
        )
        if not selected:
            return
        self.selected_file = Path(selected).resolve()
        self.path_text.set(str(self.selected_file))
        self.status_text.set("文件已选择，可以开始检查和转换。")
        self.progress_value.set(0)
        self.progress_label.set("0%")
        self._set_result("尚未开始转换。")
        self.result_section.configure(highlightbackground="#DFE7EF")
        self._set_convert_enabled(True)

    def _start_conversion(self) -> None:
        if self.selected_file is None:
            return
        self.select_button.configure(state="disabled")
        self._set_convert_enabled(False)
        self.progress_value.set(0)
        self.progress_label.set("0%")
        self.status_text.set("正在启动转换……")
        self._set_result("正在检查并转换，请稍候……")
        self._reveal_result_area()
        worker = threading.Thread(target=self._worker, args=(self.selected_file,), daemon=True)
        worker.start()

    def _worker(self, source: Path) -> None:
        def on_progress(value: int, text: str) -> None:
            self.events.put(("progress", value, text))
        try:
            report = convert_docx(source, progress=on_progress, update_with_word=True)
            self.events.put(("done", report))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _drain_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                kind = event[0]
                if kind == "progress":
                    value, text = event[1], event[2]
                    self.progress_value.set(value)
                    self.progress_label.set(f"{value}%")
                    self.status_text.set(text)
                    self._set_result("正在检查并转换，请稍候……")
                elif kind == "done":
                    self._finish(event[1])
                elif kind == "error":
                    self._fail(event[1])
        except queue.Empty:
            pass
        self.after(100, self._drain_events)

    def _finish(self, report: ConversionReport) -> None:
        self.select_button.configure(state="normal")
        self._set_convert_enabled(True)
        self.progress_value.set(100)
        self.progress_label.set("100%")
        self.status_text.set("转换完成。")
        notes = [
            f"页码域修复：{report.page_fields_rebuilt} 个",
            f"远程缓存图片内嵌：{report.include_picture_fields_unwrapped} 张",
            f"TIFF/BMP 转 PNG：{report.tiff_images_converted + report.bmp_images_converted} 张",
            f"输出文件：{report.output}",
            f"检查报告：{report.report_file}",
        ]
        if report.warnings:
            notes.append("提示：" + "；".join(report.warnings))
        self._set_result("\n".join(notes))
        self._reveal_result_area("#5A9A72", fit_content=True)
        self._show_notice(
            "success", "转换完成", "新的 WPS 兼容文档已经生成。",
            f"输出文件\n{report.output}",
        )

    def _fail(self, error: str) -> None:
        self.select_button.configure(state="normal")
        self._set_convert_enabled(self.selected_file is not None)
        self.status_text.set("转换失败。")
        self._set_result(f"转换失败：{error}")
        self._reveal_result_area("#C84A43", fit_content=True)
        self._show_notice(
            "error", "转换失败", "未能完成文档转换，请根据以下信息检查文件。",
            f"错误信息\n{error}",
        )


def main() -> None:
    app = WpsFixApp()
    app.mainloop()


if __name__ == "__main__":
    main()
