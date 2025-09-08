import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import os
import time

from dhl import handle_dhl
from ups import handle_ups
from k5 import handle_k5
from wlyups import handle_wlyups

def select_file(*args):
    # 单个文件选择
    selected_file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", ("*.xlsx", "*.xlsm"))])  # 使用askopenfilename函数选择单个文件
    select_path.set(selected_file_path)
    global select_filename
    select_filename = os.path.split(str(select_path.get()))[1]
    entry1.configure(state='normal')
    entry1.delete(0, tk.END)
    entry1.insert(0, select_filename)
    entry1.configure(state='readonly')

def select_dhl_price_file(*args):
    # 单个文件选择
    selected_dhl_price_file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", ("*.xlsx", "*.xlsm"))])  # 使用askopenfilename函数选择单个文件
    select_dhl_price_path.set(selected_dhl_price_file_path)
    global select_dhl_price_filename
    select_dhl_price_filename = os.path.split(str(select_dhl_price_path.get()))[1]
    entry2.configure(state='normal')
    entry2.delete(0, tk.END)
    entry2.insert(0, select_dhl_price_filename)
    entry2.configure(state='readonly')

def handle():
    #  更新界面processbar
    progressbar['value'] = 0
    root.update()

    fullpath = str(select_path.get())
    dhl_price_fullpath = str(select_dhl_price_path.get())

    if fullpath == '':
        tipps_value.set('提示:请选择文件')
        root.update()
        return
    else:
        tipps_value.set('提示:【文件正在读取，请耐心等待】')
        root.update()

    handle_button['state'] = 'disable'
    file_button['state'] = 'disable'
    price_button['state'] = 'disable'
    radio1['state'] = 'disable'
    radio2['state'] = 'disable'
    radio3['state'] = 'disable'
    radio4['state'] = 'disable'

    # print("start")
    start_time = time.time()
    result = ''

    if radio.get() == 1:
        result = handle_ups(fullpath, select_filename, tipps_value, root, progressbar)
    elif radio.get() == 2:
        result = handle_k5(fullpath, select_filename, tipps_value, root, progressbar)
    elif radio.get() == 3:
        result = handle_wlyups(fullpath, select_filename, tipps_value, root, progressbar)
    elif radio.get() == 4:
        result = handle_dhl(fullpath, select_filename, tipps_value, root, progressbar, dhl_price_fullpath, select_dhl_price_filename)

    if result == 'finish':
        end_time = time.time()
        tipps_value.set('处理成功，请进行核对，耗时：' + str(round(end_time - start_time, 2)) + '秒')
        root.update()
    # print("end")

    handle_button['state'] = 'normal'
    file_button['state'] = 'normal'
    price_button['state'] = 'normal'
    radio1['state'] = 'normal'
    radio2['state'] = 'normal'
    radio3['state'] = 'normal'
    radio4['state'] = 'normal'


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    root = tk.Tk()
    root.title("Excel账单处理程序")
    root.geometry("500x230")
    root.resizable(width=False, height=False)

    # 初始化Entry控件的textvariable属性值
    select_path = tk.StringVar()
    select_filename = ''

    select_dhl_price_path = tk.StringVar()
    select_dhl_price_filename = ''

    radio = tk.IntVar()
    radio.set(1)

    lines = 100
    tipps_value = tk.StringVar()
    tipps_value.set("提示: 请确保Excel文件处于关闭状态")
    # 布局控件

    tk.Label(root, text="选择处理类型：").grid(row=0, column=0, padx=(20, 0), pady=(10, 0))

    radio1 = tk.Radiobutton(root, text="UPS", variable=radio, value=1)
    radio1.grid(row=0, column=1, padx=(10, 0), pady=(10, 0))

    radio2 = tk.Radiobutton(root, text="K5", variable=radio, value=2)
    radio2.grid(row=0, column=2, pady=(10, 0))

    radio3 = tk.Radiobutton(root, text="WLY-UPS", variable=radio, value=3)
    radio3.grid(row=0, column=3, pady=(10, 0))

    radio4 = tk.Radiobutton(root, text="DHL", variable=radio, value=4)
    radio4.grid(row=0, column=4, pady=(10, 0))

    file_button = tk.Button(root, text="选择文件", command=select_file, width=10)
    file_button.grid(row=1, column=0, padx=(10, 0), pady=10)

    entry1 = tk.Entry(root, textvariable=select_filename, width=25, state='readonly')
    entry1.grid(row=1, column=1, columnspan=5, padx=(10, 0), pady=10, sticky="n,s")

    price_button = tk.Button(root, text="DHL价格", command=select_dhl_price_file, width=10)
    price_button.grid(row=2, column=0, padx=(10, 0), pady=10)

    entry2 = tk.Entry(root, textvariable=select_dhl_price_filename, width=25, state='readonly')
    entry2.grid(row=2, column=1, columnspan=5, padx=(10, 0), pady=10, sticky="n,s")

    handle_button = tk.Button(root, text="开始处理", command=handle, width=10)
    handle_button.grid(row=3, column=0, padx=(10, 0))

    progressbar = ttk.Progressbar(root, value=0, length=180, maximum=lines, mode="determinate",
                                  orient=tk.HORIZONTAL)
    progressbar.grid(row=3, column=1, columnspan=5, padx=(10, 0), pady=10, ipady=4)

    tipps = tk.Label(root, textvariable=tipps_value)
    tipps.grid(row=4, column=0, columnspan=6, padx=(20, 0), pady=10, sticky="W")

    # tk.Button(root, text="选择多个文件", command=select_files).grid(row=1, column=2)
    # tk.Button(root, text="选择文件夹", command=select_folder).grid(row=2, column=2)

    root.mainloop()
