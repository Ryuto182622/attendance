from datetime import datetime, timedelta
from tkinter import Tk, Label, Button, filedialog, messagebox, StringVar, IntVar, OptionMenu, Frame, Entry
import pandas as pd

# 日時のフォーマットを自動的に判別して解析する関数
def parse_datetime(datetime_str):
    possible_formats = [
        '%Y/%m/%d %I:%M:%S %p',
        '%Y/%m/%d %H:%M:%S',
        '%Y/%m/%d %H:%M',
        '%m/%d/%y %H:%M:%S',
        '%m/%d/%y %H:%M',
        '%m/%d/%Y %H:%M:%S',
        '%m/%d/%Y %H:%M',
    ]
    for fmt in possible_formats:
        try:
            return datetime.strptime(datetime_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"日時データ '{datetime_str}' は既知のフォーマットに一致しません")

# カスタムコースの終了時間を自動設定する関数
def update_end_time(start_time_var, end_time_var, *args):
    start_time = start_time_var.get().strip()
    try:
        start_datetime = datetime.strptime(start_time, '%H:%M')
        end_datetime = start_datetime + timedelta(minutes=90)
        end_time_var.set(end_datetime.strftime('%H:%M'))
    except ValueError:
        print(f"Invalid start time format: '{start_time}'. Expected format: HH:MM")

# 最初のクラス名が変更されたときに他のクラス名を自動で連番入力する関数
def update_class_names(*args):
    base_name = subject_vars[0].get().strip()
    if base_name:
        if base_name.isdigit():
            base_number = int(base_name)
            for i in range(1, len(subject_vars)):
                subject_vars[i].set(f"{base_number + i}")
        else:
            for i in range(1, len(subject_vars)):
                subject_vars[i].set(f"{base_name}{i + 1}")

# クラスエントリを作成する共通関数
def create_class_entry(frame, index):
    Label(frame, text=f"{index+1}:").grid(row=index, column=0)
    subject_var = StringVar()
    subject_vars.append(subject_var)
    Entry(frame, textvariable=subject_var).grid(row=index, column=1)

    if index == 0:
        subject_var.trace_add("write", update_class_names)

    Label(frame, text="開始:").grid(row=index, column=2)
    start_time_var = StringVar()
    start_time_vars.append(start_time_var)
    Entry(frame, textvariable=start_time_var).grid(row=index, column=3)

    Label(frame, text="終了時間:").grid(row=index, column=4)
    end_time_var = StringVar()
    end_time_vars.append(end_time_var)
    Entry(frame, textvariable=end_time_var).grid(row=index, column=5)

    start_time_var.trace_add("write", lambda *args, s=start_time_var, e=end_time_var: update_end_time(s, e))

# カスタムコース用のクラス入力フィールドを生成する関数
def set_custom_fields(class_count):
    for widget in class_frame.winfo_children():
        widget.destroy()

    subject_vars.clear()
    start_time_vars.clear()
    end_time_vars.clear()

    for i in range(class_count):
        create_class_entry(class_frame, i)

# コース選択時にフィールドを更新する関数
COURSE_INFO = {
    "昼間コース": ['09:00 - 10:30', '10:40 - 12:10', '13:30 - 15:00', '15:10 - 16:40', '16:50 - 18:20'],
    "006": ['09:00 - 10:30', '10:40 - 12:10', '13:00 - 14:30', '14:40 - 16:10'],
    "夜間祝日コース": ['18:30 - 20:00', '20:10 - 21:40']
}

def update_class_fields(course):
    for widget in class_frame.winfo_children():
        widget.destroy()

    subject_vars.clear()
    start_time_vars.clear()
    end_time_vars.clear()

    if course == "カスタムコース":
        Label(class_frame, text="クラス数:").grid(row=0, column=0)
        custom_class_count = IntVar(value=1)
        Entry(class_frame, textvariable=custom_class_count).grid(row=0, column=1)
        Button(class_frame, text="設定", command=lambda: set_custom_fields(custom_class_count.get())).grid(row=0, column=2)
        return

    time_options = COURSE_INFO.get(course, [])
    for i, time_option in enumerate(time_options):
        create_class_entry(class_frame, i)
        start_time_vars[i].set(time_option.split(' - ')[0])
        end_time_vars[i].set(time_option.split(' - ')[1])

# 視聴時間を合計する関数
def calculate_total_watch_time(df, class_start, class_end, year, month, day):
    class_start = datetime.strptime(class_start, '%H:%M').replace(year=year, month=month, day=day)
    class_end = datetime.strptime(class_end, '%H:%M').replace(year=year, month=month, day=day)

    total_watch_time = 0
    watched_intervals = []

    for _, row in df.iterrows():
        join_time = parse_datetime(row['参加時間'])
        leave_time = parse_datetime(row['退出日時'])

        if join_time < class_start:
            join_time = class_start
        if leave_time > class_end:
            leave_time = class_end

        if join_time < leave_time:
            watched_intervals.append((join_time, leave_time))

    if watched_intervals:
        watched_intervals.sort()
        merged_intervals = [watched_intervals[0]]

        for start, end in watched_intervals[1:]:
            last_start, last_end = merged_intervals[-1]
            if start <= last_end:
                merged_intervals[-1] = (last_start, max(last_end, end))
            else:
                merged_intervals.append((start, end))

        for start, end in merged_intervals:
            total_watch_time += (end - start).total_seconds() / 60

    class_duration = (class_end - class_start).total_seconds() / 60
    watch_rate = (total_watch_time / class_duration * 100) if class_duration > 0 else 0
    return round(min(watch_rate, 100), 1), round(total_watch_time, 1)

# CSVファイルを選択して処理する関数
def process_csv():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            messagebox.showinfo("情報", "ファイルが選択されませんでした。")
            return

        file_name = file_path.split("/")[-1]

        # 日付が入力されていない場合のみ、CSVファイル名から取得
        if not month_var.get() or not day_var.get():
            try:
                default_month_day = file_name[:4]
                default_month = int(default_month_day[:2])
                default_day = int(default_month_day[2:])
                month_var.set(default_month)
                day_var.set(default_day)
            except ValueError:
                messagebox.showerror("エラー", "CSVファイル名に不正な日付形式が含まれています。ファイル名の先頭4桁をMMDD形式に修正してください。")
                return

        df = pd.read_csv(file_path)
        year = int(year_var.get())
        month = int(month_var.get())
        day = int(day_var.get())

        if not subject_vars:
            messagebox.showerror("エラー", "クラス情報が設定されていません。クラス名と時間を設定してください。")
            return

        schedule = []
        for i in range(len(subject_vars)):
            subject = subject_vars[i].get()
            start_time = start_time_vars[i].get()
            end_time = end_time_vars[i].get()
            if not subject or not start_time or not end_time:
                messagebox.showerror("エラー", f"クラス情報が不完全です。クラス{i+1}の情報を確認してください。")
                return
            schedule.append({'subject': subject, 'start': start_time, 'end': end_time})

        watch_rates = []
        for name, group in df.groupby('名前（本来の名前）'):
            for class_info in schedule:
                subject = class_info['subject']
                if not subject:
                    continue

                class_start = class_info['start']
                class_end = class_info['end']

                try:
                    watch_rate, total_watch_time = calculate_total_watch_time(group, class_start, class_end, year, month, day)
                    watch_rates.append({
                        'Name': name,
                        'Class': subject,
                        'Watch Rate (%)': watch_rate,
                        'Total Watch Time (minutes)': total_watch_time
                    })
                except Exception as e:
                    messagebox.showerror("エラー", f"視聴率計算中にエラーが発生しました: {e}")
                    return

        watch_rates_df = pd.DataFrame(watch_rates)
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if save_path:
            watch_rates_df.to_csv(save_path, index=False)
            messagebox.showinfo("成功", "視聴率と合計視聴時間が計算され保存されました！")

    except FileNotFoundError:
        messagebox.showerror("エラー", "指定されたファイルが見つかりません。")
    except pd.errors.EmptyDataError:
        messagebox.showerror("エラー", "CSVファイルが空です。内容を確認してください。")
    except Exception as e:
        messagebox.showerror("エラー", f"予期しないエラーが発生しました: {e}")

# GUIセットアップ
root = Tk()
root.title("視聴率計算アプリ")

# 年、月、日の入力
Label(root, text="年:").grid(row=0, column=0)
year_var = StringVar(value="2025")
Entry(root, textvariable=year_var).grid(row=0, column=1)

Label(root, text="月:").grid(row=0, column=2)
month_var = StringVar(value="")
Entry(root, textvariable=month_var).grid(row=0, column=3)

Label(root, text="日:").grid(row=0, column=4)
day_var = StringVar(value="")
Entry(root, textvariable=day_var).grid(row=0, column=5)

# コース選択
Label(root, text="コース:").grid(row=1, column=0)
course_var = StringVar(value="昼間コース")
course_options = ["昼間コース", "夜間祝日コース", "カスタムコース"]
course_menu = OptionMenu(root, course_var, *course_options, command=update_class_fields)
course_menu.grid(row=1, column=1)

# 注意書き（コースとクラス情報の間）
Label(
    root,
    text="※開始時間、終了時間は(HH:MM)の形式で半角で入力してください",
    anchor="w"
).grid(row=2, column=0, columnspan=6)

# クラス情報フレーム
class_frame = Frame(root)
class_frame.grid(row=3, column=0, columnspan=6)

# 動的にクラス情報を更新
subject_vars = []
start_time_vars = []
end_time_vars = []

update_class_fields("昼間コース")

# CSV処理ボタン
Button(root, text="CSVファイルを選択して処理", command=process_csv).grid(row=4, column=0, columnspan=6, pady=10)

# メインループ
root.mainloop()
