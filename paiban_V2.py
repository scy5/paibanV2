import tkinter as tk
from tkinter import ttk, messagebox
from calendar import monthrange
import pandas as pd
import json
import os

# 配置文件路径
CONFIG_FILE = "config.json"

config = []
days = []
weekdays = []

# 创建主窗口
root = tk.Tk()
root.title("测试工具")

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as file:
            global config
            config = json.load(file)
            print(config)
    

def generate_person_row(name):
    global config
    global days
    global weekdays
    
    i = 0
    person_index = 0
    first_index = 0
    total_num = 0
    for order_name in config["分工"]["顺序"]:
        i += 1
        if order_name == name:
            person_index = i - 1
        if order_name == config["分工"]["第一个工作者"]:
            first_index = i - 1
    total_num = i
    print(name, "顺序", person_index)
    print("第一个工作者", first_index)
    print("总分工人数", total_num)

    year = config["时间"]["年"]
    month = config["时间"]["月"]
    _, num_days = monthrange(year, month)
    person_config = config["人员"][person_index]
    work = [''] * num_days
    print("num_days:", num_days)
    print("days:", days)
    if person_config["工作类型"] == "日班":
        for day in days:
            week_index = pd.Timestamp(year, month, day).weekday()
            if week_index == 5 or week_index == 6:
                work[day - 1] = "休"
            else:
                work[day - 1] = "日"
            
            if "专" in person_config["个人"]:
                if week_index + 1 in person_config["个人"]["专"]:
                    if work[day - 1] == "日":
                        work[day - 1] += "/专"
            if "蔡" in person_config["个人"]:
                if (week_index + 1) in person_config["个人"]["蔡"]:
                    if work[day - 1] == "日":
                        work[day - 1] += "/蔡"
            if "门" in person_config["个人"]:
                if (week_index + 1) in person_config["个人"]["门"]:
                    if work[day - 1] == "日":
                        work[day - 1] += "/门"
            
            print("day:", day, "week_index:", week_index, "work:", work[day - 1])
    elif  person_config["工作类型"] == "值班":
        for day in days:
            week_index = pd.Timestamp(year, month, day).weekday()
            if (day - 1 + first_index) % total_num == person_index:
                print("day", day, " on duty:", person_config["姓名"])
                work[day - 1] = "值"
                work[day - 1 + 1] = "/"
            pass
        pass

    data = [name] + work

    return data

# 定义生成表格的函数
def generate_schedule():
    global days
    global weekdays

    year = config["时间"]["年"]
    month = config["时间"]["月"]
    
    if not year or not month:
        messagebox.showerror("错误", "请配置年份和月份")
        return
    
    days_in_month = monthrange(year, month)[1]
    days = list(range(1, days_in_month + 1))
    weekdays_cn = ["一", "二", "三", "四", "五", "六", "日"]
    weekdays = [weekdays_cn[pd.Timestamp(year, month, day).weekday()] for day in days]

    # 创建表格
    headers = [""] + days
    second_row = [""] + weekdays

    data = [second_row]
    
    for person in config["人员"]:    
        person_row = generate_person_row(person["姓名"])
        data.append(person_row)


    print("结果：")
    print(data)


    # # 假设有一组固定的医生数据
    # doctor_data = ["医生A", "医生B", "医生C"]
    # for doctor in doctor_data:
    #     row = [doctor] + [""] * (len(headers) - 1)
    #     data.append(row)









    df = pd.DataFrame(data, columns=headers)
    
    # 输出到Excel文件
    output_file = f"测试结果_{year}_{month}.xlsx"
    df.to_excel(output_file, index=False, header=True)

    messagebox.showinfo("成功", f"测试结果已生成: {output_file}")

load_config()

# # 创建生成按钮
# generate_button = tk.Button(root, text="生成测试结果", command=generate_schedule)
# generate_button.grid(row=2, column=0, columnspan=6)

# # 运行主循环
# root.mainloop()

generate_schedule()