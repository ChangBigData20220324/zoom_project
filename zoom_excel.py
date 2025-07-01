from openpyxl import Workbook, load_workbook
import os
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

FILENAME = "meeting_schedule.xlsx"

# 全域變數儲存跨畫面資料
app_state = {
    "selected_date": None,
    "selected_slots": [],
    "selected_room": None,
    "user_id": "",
    "purpose": ""
}

# 初始化 Excel 檔案
def init_excel_file(filename=FILENAME):
    if not os.path.exists(filename):
        wb = Workbook()

        ws_schedule = wb.active
        ws_schedule.title = "Schedule"
        ws_schedule.append(["流水號", "預約日期", "時段ID", "會議ID", "預約人ID", "使用目的", "取消狀態"])

        ws_rooms = wb.create_sheet("MeetingRooms")
        ws_rooms.append(["會議ID", "名稱", "帳號", "密碼", "連結", "用途", "停用狀態"])
        ws_rooms.append(["A201", "第一會議室", "acc1", "pwd1", "https://meet.link/A201", "一般用途", "FALSE"])
        ws_rooms.append(["B102", "第二會議室", "acc2", "pwd2", "https://meet.link/B102", "公司內部使用", "FALSE"])
        ws_rooms.append(["C303", "第三會議室", "acc3", "pwd3", "https://meet.link/C303", "研發專用", "TRUE"])

        ws_users = wb.create_sheet("MeetingUsers")
        ws_users.append(["預約人ID", "工號", "聯絡信箱", "分機"])

        ws_slots = wb.create_sheet("TimeSlots")
        ws_slots.append(["時間ID", "時間區段", "停用狀態"])
        for i in range(1, 9):
            start = 9 + i - 1
            end = start + 1
            time_range = f"{start:02d}:00–{end:02d}:00"
            ws_slots.append([i, time_range, "FALSE"])

        wb.save(filename)

# 載入可用時段
def load_time_slots(filename=FILENAME):
    wb = load_workbook(filename)
    ws = wb["TimeSlots"]
    time_slots = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        slot_id, time_range, disabled = row
        if str(disabled).strip().upper() != "TRUE":
            time_slots[int(slot_id)] = time_range
    return time_slots

# 查詢當天某些時段有哪些會議室可用
def get_available_rooms(date, slot_ids):
    wb = load_workbook(FILENAME)
    ws_schedule = wb["Schedule"]
    ws_rooms = wb["MeetingRooms"]

    booked_rooms = set()
    for row in ws_schedule.iter_rows(min_row=2, values_only=True):
        record_date, existing_slots, meeting_id, canceled = row[1], row[2], row[3], row[6]
        if record_date == date and not canceled:
            booked_slots = set(map(int, existing_slots.split(',')))
            if any(slot in booked_slots for slot in slot_ids):
                booked_rooms.add(meeting_id)

    available_rooms = []
    for row in ws_rooms.iter_rows(min_row=2, values_only=True):
        room_id, name, acc, pwd, link, usage, closed = row
        if room_id not in booked_rooms and str(closed).strip().upper() != "TRUE":
            available_rooms.append((room_id, name, usage))
    return available_rooms

# 寫入預約
def add_booking():
    wb = load_workbook(FILENAME)
    ws = wb["Schedule"]
    new_id = 1
    for row in ws.iter_rows(min_row=2, values_only=True):
        if isinstance(row[0], int):
            new_id = max(new_id, row[0] + 1)
    slot_str = ",".join(map(str, app_state["selected_slots"]))
    ws.append([new_id, app_state["selected_date"], slot_str, app_state["selected_room"],
               app_state["user_id"], app_state["purpose"], False])
    wb.save(FILENAME)

# 主應用程式
class MeetingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("會議室預約系統")
        self.geometry("550x450")
        self.frames = {}

        for F in (PageDateInput, PageTimeSelect, PageRoomSelect, PageConfirm, PageFinish):
            page_name = F.__name__
            frame = F(parent=self, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("PageDateInput")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()
        if hasattr(frame, "refresh"):
            frame.refresh()

# 頁面一：輸入日期
class PageDateInput(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        tk.Label(self, text="輸入預約日期（YYYY-MM-DD）", font=("Arial", 14)).pack(pady=20)
        self.date_entry = tk.Entry(self)
        self.date_entry.pack()
        tk.Button(self, text="查詢可預約時段", command=self.next_page).pack(pady=20)

    def next_page(self):
        date = self.date_entry.get().strip()
        try:
            datetime.strptime(date, "%Y-%m-%d")
            app_state["selected_date"] = date
            self.controller.show_frame("PageTimeSelect")
        except ValueError:
            messagebox.showerror("錯誤", "請輸入正確的日期格式 YYYY-MM-DD")

# 頁面二：選擇可預約時段
class PageTimeSelect(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.labels = []
        self.vars = {}
        self.time_slots = load_time_slots()
        tk.Label(self, text="選擇可預約時段", font=("Arial", 14)).pack(pady=10)
        self.slot_frame = tk.Frame(self)
        self.slot_frame.pack()
        tk.Button(self, text="下一步", command=self.next_page).pack(pady=10)

    def refresh(self):
        for widget in self.slot_frame.winfo_children():
            widget.destroy()
        self.vars = {}

        date = app_state["selected_date"]
        wb = load_workbook(FILENAME)
        ws = wb["Schedule"]
        used = {i: False for i in self.time_slots}

        # 判斷當天哪些時段已被預約
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[1] == date and not row[6]:
                for sid in map(int, row[2].split(',')):
                    used[sid] = True

        # 篩選尚未被預約的時段
        available_slots = [(sid, time_str) for sid, time_str in self.time_slots.items() if not used[sid]]

        # 如果沒有可預約時段，跳出訊息並返回首頁
        if not available_slots:
            messagebox.showinfo("預約已滿", f"{date} 所有時段都已被預約，請選擇其他日期")
            self.controller.show_frame("PageDateInput")
            return

        # 顯示可預約時段的 Checkbutton
        for sid, time_str in available_slots:
            var = tk.IntVar()
            cb = tk.Checkbutton(self.slot_frame, text=f"{sid}: {time_str}", variable=var)
            cb.pack(anchor="w")
            self.vars[sid] = var


    def next_page(self):
        app_state["selected_slots"] = [sid for sid, var in self.vars.items() if var.get()]
        if not app_state["selected_slots"]:
            messagebox.showwarning("提醒", "請至少選擇一個時段")
            return
        self.controller.show_frame("PageRoomSelect")

# 頁面三：選擇會議室
class PageRoomSelect(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        tk.Label(self, text="選擇會議室", font=("Arial", 14)).pack(pady=10)
        self.room_var = tk.StringVar()
        self.room_frame = tk.Frame(self)
        self.room_frame.pack()
        tk.Button(self, text="下一步", command=self.next_page).pack(pady=10)

    def refresh(self):
        for widget in self.room_frame.winfo_children():
            widget.destroy()
        date = app_state["selected_date"]
        slots = app_state["selected_slots"]
        rooms = get_available_rooms(date, slots)
        if not rooms:
            tk.Label(self.room_frame, text="無可用會議室，請返回上一步").pack()
            return
        for room_id, name, usage in rooms:
            rb = tk.Radiobutton(self.room_frame, text=f"{room_id} - {name}（{usage}）", variable=self.room_var, value=room_id)
            rb.pack(anchor="w")

    def next_page(self):
        room = self.room_var.get()
        if not room:
            messagebox.showwarning("提醒", "請選擇一間會議室")
            return
        app_state["selected_room"] = room
        self.controller.show_frame("PageConfirm")

# 頁面四：輸入預約人資料
class PageConfirm(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        tk.Label(self, text="輸入預約資訊", font=("Arial", 14)).pack(pady=10)
        tk.Label(self, text="預約人ID：").pack()
        self.entry_user = tk.Entry(self)
        self.entry_user.pack()
        tk.Label(self, text="使用目的：").pack()
        self.entry_purpose = tk.Entry(self)
        self.entry_purpose.pack()
        tk.Button(self, text="完成預約", command=self.finish).pack(pady=10)

    def finish(self):
        uid = self.entry_user.get().strip()
        purpose = self.entry_purpose.get().strip()
        if not uid or not purpose:
            messagebox.showerror("錯誤", "請填寫所有欄位")
            return
        app_state["user_id"] = uid
        app_state["purpose"] = purpose
        add_booking()
        self.controller.show_frame("PageFinish")

# 頁面五：完成畫面
class PageFinish(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        tk.Label(self, text="預約完成", font=("Arial", 16)).pack(pady=30)
        tk.Button(self, text="返回首頁", command=lambda: controller.show_frame("PageDateInput")).pack(pady=10)

# 執行主程式
if __name__ == "__main__":
    init_excel_file()
    app = MeetingApp()
    app.mainloop()
