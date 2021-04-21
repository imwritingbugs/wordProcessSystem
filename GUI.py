from tkinter import *
from tkinter.filedialog import askopenfilename, askdirectory
from wordProcess import *
import tkinter.messagebox as box
import os

file_path = ''
dir_path = ''
is_file = True


def select_file_path():
    global file_path
    global is_file
    is_file = True
    path_ = askopenfilename()
    path.set(path_)
    file_path = path_
    # print(file_path)


def select_dir_path():
    global is_file
    global dir_path
    is_file = False
    dir_path = askdirectory()
    path.set(dir_path)


def confirm_check():
    global file_path
    global is_file
    global dir_path
    if is_file:
        # 选择了文件
        # print(file_path)
        file_dir_list = file_path.split("/")[:-1]
        txt_file_dir = ''
        # print("filedir", file_dir_list)
        for item in file_dir_list:
            txt_file_dir += item
            txt_file_dir += '/'
        # print(txt_file_dir)
        # 检查是否存在改文件s
        if os.path.exists(file_path):
            if file_path.endswith(".docx"):
                err_list, info_list, file_total_cnt = parse_file(file_path)
                err_message = ""
                for e in err_list:
                    err_message += e
                    err_message += "\n"
                # print("Err message", err_message)
                if len(err_message) != 0:
                    box.showerror(title="错误", message=err_message)
                    try:
                        fh = open(txt_file_dir + "./single_error.txt", 'w', encoding="utf-8")
                        fh.writelines(file_path + '\n')
                        fh.writelines(err_message + '\n\n')
                    except IOError:
                        box.showerror(title="错误", message="错误信息文件打开失败")
                    else:
                        print(file_path + "处理完成！\n")
                        box.showinfo(title="信息保存成功", message="错误信息已保存在" + txt_file_dir + "single_error.txt中")
                        fh.close()
                    # print("err", err_list)
                    # print("info", info_list)
                else:
                    print(file_path + " 处理完成！\n")
                    box.showinfo(title="成功", message="系统未检测到错误")
            else:
                box.showerror(title="错误", message="未选择.docx文件")
        else:
            box.showerror(title="错误", message="未找到该文件，可能已重命名。请重新选择")

    else:
        # 选择了目录
        # 统计整个文件夹中的总字数
        folder_total_num = 0
        file_list = []
        files = os.listdir(dir_path)
        for file in files:
            if file.endswith(".docx") and "~$" not in file:
                file_list.append(dir_path + '/' + file)
        # print(file_list)
        if len(file_list) == 0:
            box.showerror(title="错误", message="文件夹中未检测到.docx文件")
        else:
            err_message_list = []
            file_total_num = 0
            for docx_file in file_list:
                # 如果文件被改名了，此时记录的是老名字，可能会找不到该文件
                if os.path.exists(docx_file):
                    err_list, info_list, file_total_num = parse_file(docx_file)
                    err_message = ""
                    for e in err_list:
                        err_message += e
                        err_message += "\n"
                    # print("err", err_message)
                    if len(err_message) != 0:
                        err_message_list.append(err_message)
                    else:
                        err_message_list.append("未检测到错误")
                    folder_total_num += file_total_num
                    print(docx_file + " 处理完成！\n")
                else:
                    box.showerror(title="错误", message="文件" + docx_file + '不存在，请重新选择文件夹以刷新')

            #   填写错误信息文件
            try:
                # print("file_list", file_list, len(file_list))
                # print("err_list", err_message_list, len(err_message_list))
                fh = open(dir_path + "/group_error.txt", 'w', encoding="utf-8")
                fh.write(f"整个文件夹中的docx文件总字数为：{folder_total_num}\n")
                for i in range(len(file_list)):
                    # print(i)
                    fh.write(file_list[i] + "\n")
                    fh.write(err_message_list[i] + "\n\n")
            except IOError:
                box.showerror(title="错误", message="错误信息文件打开失败")
            else:
                box.showinfo(title="信息保存成功", message="错误信息已保存在" + dir_path + "/group_error.txt中")
                fh.close()


root = Tk()
path = StringVar()
root.title("会议纪要自动评判系统v1.3")
Label(root, text="目标路径:").grid(row=0, column=0)
Entry(root, textvariable=path, width=50).grid(row=0, column=1)
Button(root, text="文件路径选择", command=select_file_path).grid(row=0, column=2)
Button(root, text="文件夹路径选择", command=select_dir_path).grid(row=0, column=3)
Button(root, text="确认检测", command=confirm_check).grid(row=0, column=4)
root.mainloop()
