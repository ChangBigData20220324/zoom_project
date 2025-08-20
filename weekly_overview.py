import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
from openpyxl import load_workbook
from utils import  render_weekly_table, render_boss_table

class PageWeeklyOverview(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="white")
        self.controller = controller

        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(0, weight=1)


        # ====== 分頁與返回首頁 ======
        self.tab_frame = tk.Frame(self, bg="white")
        self.tab_frame.grid(row=0, column=0, sticky="ew", padx=40, pady=(20, 0))
        self.tab_frame.grid_columnconfigure((0, 1), weight=1)
        self.tab_frame.grid_columnconfigure(2, weight=0)

        self.tab_buttons = {}
        self.current_tab = "weekly"

        self.tab_buttons["weekly"] = tk.Button(
            self.tab_frame,
            text="本週預約總覽",
            command=lambda: self.switch_tab("weekly"),
            font=("Arial", 11, "bold"),
            bg="#3B82F6",
            fg="white",
            relief="flat",
            padx=10,
            pady=6
        )
        self.tab_buttons["weekly"].grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.tab_buttons["boss"] = tk.Button(
            self.tab_frame,
            text="Boss 專區",
            command=lambda: self.switch_tab("boss"),
            font=("Arial", 11, "bold"),
            bg="white",
            fg="#3B82F6",
            relief="solid",
            borderwidth=1,
            highlightbackground="#3B82F6",
            padx=10,
            pady=6
        )
        self.tab_buttons["boss"].grid(row=0, column=1, sticky="ew", padx=(4, 20))

        self.btn_home = tk.Button(
            self.tab_frame,
            text="返回首頁",
            command=lambda: controller.show_frame("PageDateInput"),
            font=("Arial", 10),
            bg="#3B82F6",
            fg="white",
            relief="flat",
            padx=10,
            pady=4,
            borderwidth=0
        )
        self.btn_home.grid(row=0, column=2, sticky="e")

        # ====== 圓點說明區（含重新整理按鈕） ======
        self.legend_frame = tk.Frame(self, bg="#f3f4f6", highlightbackground="#d1d5db", highlightthickness=1)
        self.legend_frame.grid(row=1, column=0, sticky="ew", padx=40, pady=(10, 6))

        # ====== 表格區塊 ======
        self.content_frame = tk.Frame(self, bg="white")
        self.content_frame.grid(row=3, column=0, sticky="nsew", padx=40, pady=(0, 30))

        self.table_frame = tk.Frame(self.content_frame, bg="white")
        self.boss_frame = tk.Frame(self.content_frame, bg="white")

        self.table_frame.pack(fill="both", expand=True)

        # ✅ 初始化時顯示 legend
        self.build_legend("weekly")

    def build_legend(self, tab_name):
        for widget in self.legend_frame.winfo_children():
            widget.destroy()

        # ========== 外層水平分區 ==========
        left = tk.Frame(self.legend_frame, bg="#f3f4f6")
        right = tk.Frame(self.legend_frame, bg="#f3f4f6")
        left.pack(side="left", padx=10, pady=8, fill="x", expand=True)
        right.pack(side="right", padx=10, pady=8)

        legend_map = {
            "weekly": [
                ("#4caf50", "全部可借"),
                ("#f59e0b", "部分可借"),
                ("#ef4444", "無可借"),
            ],
            "boss": [
                ("#ef4444", "有固定預約"),
                ("#3B82F6", "MIS 需支援"),
            ]
        }

        items = legend_map.get(tab_name, [])

        for i, (color, text) in enumerate(items):
            canvas = tk.Canvas(left, width=16, height=16, bg="#f3f4f6", highlightthickness=0)
            canvas.pack(side="left", padx=(0, 4))
            canvas.create_oval(2, 2, 14, 14, fill=color, outline=color)

            label = tk.Label(
                left,
                text=text,
                font=("Arial", 10),
                bg="#f3f4f6",
                fg="#111827"
            )
            label.pack(side="left", padx=(0, 20))

        # ========== 右側重新整理按鈕 ==========
        btn_refresh = tk.Button(
            right,
            text="🔄 重新整理",
            command=self.on_refresh,
            font=("Arial", 9),
            bg="#e5e7eb",        # 淡灰色底
            fg="#374151",        # 深灰字
            activebackground="#d1d5db",  # 點擊時稍微變灰
            relief="flat",
            bd=0,
            padx=8,
            pady=2,
            cursor="hand2"
        )

        btn_refresh.pack()

    def on_refresh(self):
        self.refresh()
        messagebox.showinfo("已更新", "預約資料已重新整理完成")
    def switch_tab(self, tab_name):
        if self.current_tab == tab_name:
            return

        for name, btn in self.tab_buttons.items():
            if name == tab_name:
                btn.config(bg="#3B82F6", fg="white", relief="flat")
            else:
                btn.config(bg="white", fg="#3B82F6", relief="solid", borderwidth=1)

        for child in self.content_frame.winfo_children():
            child.pack_forget()

        self.build_legend(tab_name)

        if tab_name == "weekly":
            self.table_frame.pack(fill="both", expand=True)
        elif tab_name == "boss":
            self.boss_frame.pack(fill="both", expand=True)

        self.current_tab = tab_name
        self.refresh()

    def refresh(self):
        if self.current_tab == "weekly":
            render_weekly_table(self.table_frame)
        elif self.current_tab == "boss":
            render_boss_table(self.boss_frame)

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1000x600")
    page = PageWeeklyOverview(root, None)
    page.pack(fill="both", expand=True)
    page.refresh()
    root.mainloop()
