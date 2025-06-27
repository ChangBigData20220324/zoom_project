from openpyxl import Workbook, load_workbook
import os

FILENAME = "meeting_schedule.xlsx"

# 從 TimeSlots 工作表載入時間對照表
def load_time_slots(filename=FILENAME):
    wb = load_workbook(filename)
    ws = wb["TimeSlots"]
    time_slots = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        slot_id, time_range, disabled = row
        if str(disabled).strip().upper() != "TRUE":
            time_slots[int(slot_id)] = time_range
    return time_slots

# 初始化 Excel 檔案並建立必要的工作表
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
    else:
        wb = load_workbook(filename)
        created = False

        if "Schedule" not in wb.sheetnames:
            ws = wb.create_sheet("Schedule")
            ws.append(["流水號", "預約日期", "時段ID", "會議ID", "預約人ID", "使用目的", "取消狀態"])
            created = True

        if "MeetingRooms" not in wb.sheetnames:
            ws = wb.create_sheet("MeetingRooms")
            ws.append(["會議ID", "名稱", "帳號", "密碼", "連結", "用途", "停用狀態"])
            ws.append(["A201", "第一會議室", "acc1", "pwd1", "https://meet.link/A201", "一般用途", "FALSE"])
            ws.append(["B102", "第二會議室", "acc2", "pwd2", "https://meet.link/B102", "公司內部使用", "FALSE"])
            ws.append(["C303", "第三會議室", "acc3", "pwd3", "https://meet.link/C303", "研發專用", "TRUE"])
            created = True

        if "MeetingUsers" not in wb.sheetnames:
            ws = wb.create_sheet("MeetingUsers")
            ws.append(["預約人ID", "工號", "聯絡信箱", "分機"])
            created = True

        if "TimeSlots" not in wb.sheetnames:
            ws = wb.create_sheet("TimeSlots")
            ws.append(["時間ID", "時間區段", "停用狀態"])
            for i in range(1, 9):
                start = 9 + i - 1
                end = start + 1
                time_range = f"{start:02d}:00–{end:02d}:00"
                ws.append([i, time_range, "FALSE"])
            created = True

        if created:
            wb.save(filename)

# 顯示會議室預約狀況
def check_availability(target_date, filename=FILENAME):
    time_slots = load_time_slots(filename)
    wb = load_workbook(filename)
    ws = wb["Schedule"]
    schedule = {i: [] for i in time_slots}

    for row in ws.iter_rows(min_row=2, values_only=True):
        record_date, slot_ids_str, meeting_id, canceled = row[1], row[2], row[3], row[6]
        if record_date == target_date and not canceled:
            for sid in map(int, slot_ids_str.split(',')):
                if sid in schedule:
                    schedule[sid].append(meeting_id)

    print(f"\n【{target_date} 會議室預約狀況】")
    for slot_id, time_str in time_slots.items():
        rooms = schedule.get(slot_id, [])
        if rooms:
            print(f"時段 {slot_id}（{time_str}）：已預約 - {', '.join(map(str, rooms))}")
        else:
            print(f"時段 {slot_id}（{time_str}）：可預約")

# 使用者輸入合法時段
def input_slot_ids(filename=FILENAME):
    time_slots = load_time_slots(filename)
    valid_ids = set(time_slots.keys())
    while True:
        user_input = input("請輸入欲預約的時段 ID（可輸入多個，以逗號分隔）：").strip()
        try:
            slot_ids = list(map(int, user_input.split(',')))
            if not slot_ids:
                print("不可為空白，請重新輸入。")
                continue
            if not all(slot in valid_ids for slot in slot_ids):
                print(f"輸入的時段中有無效或已停用的 ID，目前可用時段為：{sorted(valid_ids)}")
                continue
            return slot_ids
        except ValueError:
            print("格式錯誤，請輸入整數，例如：2 或 2,3,4")

# 查詢可用會議室
def get_available_meeting_rooms(date, slot_ids, filename=FILENAME):
    wb = load_workbook(filename)
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
        room_id, name, acc, pwd, link, usage, closed = row[:7]
        if room_id not in booked_rooms and not str(closed).strip().upper() == "TRUE":
            available_rooms.append({
                "ID": room_id,
                "名稱": name,
                "帳號": acc,
                "密碼": pwd,
                "連結": link,
                "用途": usage
            })

    return available_rooms

# 自動產生流水號
def get_next_id(ws):
    max_id = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and isinstance(row[0], int):
            max_id = max(max_id, row[0])
    return max_id + 1

# 新增預約紀錄
def add_booking(date, slot_ids, meeting_id, user_id, purpose, filename=FILENAME):
    wb = load_workbook(filename)
    ws = wb["Schedule"]
    new_id = get_next_id(ws)
    slot_str = ",".join(map(str, slot_ids))
    new_record = [new_id, date, slot_str, meeting_id, user_id, purpose, False]
    ws.append(new_record)
    wb.save(filename)
    print(f"已新增預約紀錄（流水號 {new_id}）")

# 主流程
def run_cli():
    init_excel_file()
    date = input("請輸入預約日期（格式 YYYY-MM-DD）：").strip()
    check_availability(date)

    while True:
        slot_ids = input_slot_ids()
        available_rooms = get_available_meeting_rooms(date, slot_ids)
        if not available_rooms:
            print("此時段無可用會議室，請更換時段。")
            continue
        break

    print("\n可用會議室列表：")
    for idx, room in enumerate(available_rooms, start=1):
        print(f"{idx}) {room['名稱']}（用途：{room['用途']}）")

    while True:
        try:
            choice = int(input("請選擇會議室（輸入數字）：").strip())
            if 1 <= choice <= len(available_rooms):
                meeting_id = available_rooms[choice - 1]["ID"]
                break
            else:
                print("選項超出範圍，請重新輸入。")
        except ValueError:
            print("請輸入正確的數字編號。")

    user_id = input("請輸入預約人 ID（如 U001）：").strip().upper()
    purpose = input("請輸入使用目的（如：專案會議）：").strip()
    if not all([meeting_id, user_id, purpose]):
        print("所有欄位皆為必填，請重新執行。")
        return

    add_booking(date, slot_ids, meeting_id, user_id, purpose)
    check_availability(date)

# 執行主流程
if __name__ == "__main__":
    run_cli()
