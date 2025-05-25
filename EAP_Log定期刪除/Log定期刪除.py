import os
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

def get_recent_months(n=3):
    today = datetime.today()
    months = []
    for i in range(n):
        y = (today.year * 12 + today.month - 1 - i) // 12
        m = (today.year * 12 + today.month - 1 - i) % 12 + 1
        months.append(f"{y}{m:02d}")
    return set(months)

def find_all_log_paths(root_path):
    log_paths = []
    for dirpath, dirnames, _ in os.walk(root_path):
        for dirname in dirnames:
            if dirname.lower() == "log":
                full_path = os.path.join(dirpath, dirname)
                log_paths.append(full_path)
    return log_paths

def find_old_folders(root_path, keep_months):
    folders_to_delete = []
    log_paths = find_all_log_paths(root_path)
    if not log_paths:
        print(f"找不到名為 Log 的資料夾於：{root_path}")
        return []

    print(f" 找到 {len(log_paths)} 個 Log 資料夾")
    for log_path in log_paths:
        for dirpath, dirnames, _ in os.walk(log_path, topdown=False):
            for dirname in dirnames:
                if len(dirname) == 6 and dirname.isdigit():
                    if dirname not in keep_months:
                        full_path = os.path.join(dirpath, dirname)
                        folders_to_delete.append(full_path)
    return folders_to_delete

def delete_folders(folder_paths):
    deleted = 0
    for folder_path in folder_paths:
        try:
            shutil.rmtree(folder_path)
            print(f" 已刪除：{folder_path}")
            deleted += 1
        except Exception as e:
            print(f"無法刪除 {folder_path}：{e}")
    print(f" 共刪除 {deleted} 筆資料")

def run_cleanup_with_preview():
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="請選擇路徑")
    if not folder_selected:
        print(" 已取消。")
        return
    print(" 選取路徑：", folder_selected)

    keep_months = get_recent_months(3)
    folders_to_delete = find_old_folders(folder_selected, keep_months)
    if not folders_to_delete:
        messagebox.showinfo("結果", "未找到需刪除的資料。")
        return

    preview = "\n".join(folders_to_delete)
    print(preview)
    confirm = messagebox.askyesno("刪除確認", f"找到『{len(folders_to_delete)}』筆資料將被刪除，是否繼續？")
    if confirm:
        delete_folders(folders_to_delete)
        messagebox.showinfo("完成", "刪除完成。")
    else:
        print(" 取消刪除。")

run_cleanup_with_preview()
