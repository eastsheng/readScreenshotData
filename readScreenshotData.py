import pyautogui
import pytesseract
from PIL import Image
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import os
import numpy as np
import readlammpsdata as rld
# 配置 tesseract 路径
pytesseract.pytesseract.tesseract_cmd = r'D:\Program Files\Tesseract-OCR\tesseract.exe'

# 初始化坐标变量
start_x = start_y = end_x = end_y = 0

# 定义鼠标事件处理函数
def on_mouse_down(event):
    global start_x, start_y
    start_x, start_y = event.x_root, event.y_root

def on_mouse_drag(event):
    global end_x, end_y
    end_x, end_y = event.x_root, event.y_root
    canvas.coords(rect, start_x, start_y, end_x, end_y)

def on_mouse_up(event):
    global start_x, start_y, end_x, end_y
    root.quit()  # 退出 Tkinter 主循环

# 创建 Tkinter 窗口，用于捕捉鼠标的选取区域
root = tk.Tk()
root.attributes("-fullscreen", True)  # 全屏窗口
root.attributes("-alpha", 0.5)  # 半透明效果
canvas = tk.Canvas(root, cursor="cross")
canvas.pack(fill="both", expand=True)

# 绘制矩形用于显示选取的区域
rect = canvas.create_rectangle(0, 0, 0, 0, outline='red')

# 绑定鼠标事件
canvas.bind("<ButtonPress-1>", on_mouse_down)  # 鼠标按下
canvas.bind("<B1-Motion>", on_mouse_drag)  # 鼠标拖动
canvas.bind("<ButtonRelease-1>", on_mouse_up)  # 鼠标释放

# 运行 Tkinter 主循环以捕获区域
root.mainloop()

# 确定截图区域
x1, y1 = min(start_x, end_x), min(start_y, end_y)
x2, y2 = max(start_x, end_x), max(start_y, end_y)
width, height = x2 - x1, y2 - y1

# 如果截图区域不为空，进行截图并识别
if width > 0 and height > 0:
    # 截图指定区域
    screenshot = pyautogui.screenshot(region=(x1, y1, width, height))
    screenshot_path = "screenshot.png"
    screenshot.save(screenshot_path)

    # 打开截图并使用 pytesseract 进行文字识别
    image = Image.open(screenshot_path)
    text = pytesseract.image_to_string(image)
    # 将识别的内容拆分为每行
    lines = text.strip().split('\n')
    lines = [x for x in lines if x]
    # 创建一个 Pandas DataFrame 来存储识别结果
    data = {'Text': lines}
    df = pd.DataFrame(data)
    # print(df.values)
    strs = rld.array2str(df.values).strip()
    # 打印识别的内容
    print("识别的内容：")
    print(strs)
    # 将 DataFrame 写入 Excel 文件
    output_excel = "output.xlsx"
    df.to_excel(output_excel, index=False)

    print(f"识别结果已保存为 {output_excel}")

    # 清理临时文件
    if os.path.exists(screenshot_path):
        os.remove(screenshot_path)

else:
    messagebox.showerror("错误", "选择的区域无效，请重试！")
