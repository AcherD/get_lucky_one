import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import random
import pandas as pd

# 抽奖系统类
class LotterySystem:
    def __init__(self):
        self.candidates = []  # 候选人列表

    def add_candidate(self, candidate):
        self.candidates.append(candidate)

    def draw_winners(self):
        if len(self.candidates) < 3:
            return None
        winners = random.sample(self.candidates, 3)  # 随机选取3个中奖者
        return winners

# 创建主窗口
root = tk.Tk()
root.title("抽奖系统")

# 创建抽奖系统对象
lottery_system = LotterySystem()

# 创建滚动区域
canvas = tk.Canvas(root)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(root, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=frame, anchor="nw")

# 处理添加候选人按钮点击事件
def add_candidate():
    candidate = entry.get()
    if candidate:
        lottery_system.add_candidate(candidate)
        listbox.insert(tk.END, candidate)
        entry.delete(0, tk.END)

# 处理从文件导入按钮点击事件
def import_from_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            candidates = df.iloc[:, 0].tolist()  # 读取第一列的数据
            for candidate in candidates:
                if not pd.isna(candidate):  # 排除空值
                    lottery_system.add_candidate(str(candidate))
                    listbox.insert(tk.END, str(candidate))
            messagebox.showinfo("成功", "成功导入候选人！")
        except Exception as e:
            messagebox.showerror("错误", f"导入候选人时发生错误：{e}")

# 处理抽奖按钮点击事件
def draw_winners():
    winners = lottery_system.draw_winners()
    if winners is None:
        messagebox.showerror("错误", "候选人数量不足3人，请添加更多候选人！")
    else:
        result = "\n".join(winners)
        result_label.config(text=result)

# 创建界面元素
label = tk.Label(frame, text="请输入候选人名称:")
label.pack(pady=10)

entry = tk.Entry(frame, width=30)
entry.pack()

add_button = tk.Button(frame, text="添加候选人", command=add_candidate)
add_button.pack(pady=10)

import_button = tk.Button(frame, text="从文件导入", command=import_from_file)
import_button.pack(pady=10)

listbox = tk.Listbox(frame)
listbox.pack()

draw_button = tk.Button(frame, text="开始抽奖", command=draw_winners)
draw_button.pack(pady=10)

result_label = tk.Label(frame, text="")
result_label.pack()
canvas.create_window((0, 0), window=frame, anchor="nw")

root.mainloop()