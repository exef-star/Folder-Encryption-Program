# -*- coding: utf-8 -*-
from colorama import Fore, Back, Style
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from cryptography.fernet import Fernet, InvalidToken
import psutil
import os
import zipfile
from tkinter import filedialog, messagebox, simpledialog, PhotoImage
import datetime
import sys
from PIL import Image, ImageTk
import traceback
import platform
import getpass
import webbrowser
import base64
import win32api
import win32con
import random
import string
import json
import ctypes

FolderEncryptorVersion = "v3.11-Alpha"

print("Folder Encryptor")
print(FolderEncryptorVersion)
print("MIT License")
print("By HantaFrog")
print("-------------------------------")

class FolderEncryptor(ttk.Window):
    def __init__(self):
        super().__init__(themename="lumen")
        self.title("文件夹加密程序&By HantaFrog | 2025年1月27日更新")
        self.geometry("973x480")
        self.resizable(False, False)
        self.key = None

        # 设置背景图片
        self.set_background_image()

        # 创建顶部导航栏
        self.create_menu_bar()

        self.create_widgets()

        def load_sentences_from_json(file_path):
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)
                return data.get('sentences', [])

        def show_random_sentence(sentences):
            if not sentences:
                messagebox.showinfo("提示", "没有找到句子。")
                return
            random_sentence = random.choice(sentences)
            messagebox.showinfo("你知道吗（*゜ー゜*）", random_sentence)

        # JSON文件路径
        json_file_path = 'icon/areyouknow.json'
        # 加载句子
        sentences = load_sentences_from_json(json_file_path)
        # 显示随机句子
        show_random_sentence(sentences)

    def set_background_image(self):
        # 加载图片
        imgpath = 'fepimg/background5.png'
        img = Image.open(imgpath)
        img = img.resize((973, 580))
        self.background_image = ImageTk.PhotoImage(img)

        # 创建Canvas并设置背景图片
        self.canvas = ttk.Canvas(self, width=img.width, height=img.height)
        self.canvas.create_image(0, 0, anchor=ttk.NW, image=self.background_image)
        self.canvas.pack(fill=ttk.BOTH, expand=True)

    def create_menu_bar(self):
        self.menu_bar = ttk.Menu(self)
        self.config(menu=self.menu_bar)
        self.file_icon = PhotoImage(file='icon/fileicon.png')
        self.file_choose_icon = PhotoImage(file='icon/filechicon.png')
        self.file_view_icon = PhotoImage(file='icon/filevw.png')
        self.author_icon = PhotoImage(file='icon/authoricon.png')
        self.programm_icon = PhotoImage(file='icon/programmicon.png')
        self.exit_icon = PhotoImage(file='icon/exiticon.png')
        self.update_icon = PhotoImage(file='icon/update.png')
        self.help_icon = PhotoImage(file='icon/helpdc.png')
        self.debug_icon = PhotoImage(file='icon/ainfo.png')
        self.jiami_icon = PhotoImage(file='icon/jiami.png')
        self.jiemi_icon = PhotoImage(file='icon/jiemi.png')
        self.defolder_icon = PhotoImage(file='icon/deletefolder.png')
        self.crfolder_icon = PhotoImage(file='icon/createfolder.png')
        self.refolder_icon = PhotoImage(file='icon/renamefolder.png')
        self.ziptotxt_icon = PhotoImage(file='icon/ziptotxt.png')
        self.changefolder_icon = PhotoImage(file='icon/chfolic.png')
        self.printuuFFFDencodeutf82decodegbk_icon = PhotoImage(file='icon/printuuFFFDencodeutf82decodegbk.png')
        self.areyouknow_icon = PhotoImage(file='icon/ainfo.png')
        self.github_icon = PhotoImage(file='icon/github.png')

        # 创建文件菜单
        file_menu = ttk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="生成密钥文件", command=self.generate_key, image=self.file_icon, compound='left')
        file_menu.add_command(label="选择密钥文件", command=self.load_key, image=self.file_choose_icon, compound='left')
        file_menu.add_command(label="查看密钥文件", command=self.view_key, image=self.file_view_icon, compound='left')
        file_menu.add_separator()
        file_menu.add_command(label="退出该程序", command=self.quit, image=self.exit_icon, compound='left')

        # 创建操作菜单
        action_menu = ttk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="菜单", menu=action_menu)
        action_menu.add_command(label="加密文件夹", command=self.encrypt_folder, image=self.jiami_icon, compound='left')
        action_menu.add_command(label="解密文件夹", command=self.decrypt_folder, image=self.jiemi_icon, compound='left')
        action_menu.add_command(label="删除文件夹", command=self.delete_folder, image=self.defolder_icon, compound='left')
        action_menu.add_command(label="新建文件夹", command=self.create_folder, image=self.crfolder_icon, compound='left')
        action_menu.add_command(label="重命名文件夹", command=self.rename_folder, image=self.refolder_icon, compound='left')
        action_menu.add_command(label="打包文件夹为ZIP并转换为TXT", command=self.zip_and_convert_to_txt, image=self.ziptotxt_icon, compound='left')
        action_menu.add_command(label="更改文件夹图标", command=self.change_folder_icon, image=self.changefolder_icon, compound='left')
        action_menu.add_command(label="更改文件夹内文件（不含后缀名）为乱码", command=self.rename_filename_random, image=self.printuuFFFDencodeutf82decodegbk_icon, compound='left')

        # 慎重启用
        warning_menu = ttk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="慎重启用", menu=warning_menu)
        warning_menu.add_command(label="生成Debug文件", command=self.generate_Debug, image=self.debug_icon, compound='left')

        # 创建关于菜单
        about_menu = ttk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="关于 · · ·", menu=about_menu)
        about_menu.add_command(label="关于本程序", command=self.about_program, image=self.programm_icon, compound='left')
        about_menu.add_command(label="关于作者", command=self.about_author, image=self.author_icon, compound='left')
        about_menu.add_command(label="更新日志", command=self.update_log, image=self.update_icon, compound='left')
        about_menu.add_command(label="查看帮助文档", command=self.help, image=self.help_icon, compound='left')
        about_menu.add_command(label="查看作者的博客", command=self.blog, image=self.author_icon, compound='left')
        about_menu.add_command(label="查看作者的Github", command=self.github_another_way, image=self.author_icon, compound='left')
        about_menu.add_command(label="访问Github仓库", command=self.github_way, image=self.github_icon, compound='left')

    def blog(self):
        webbrowser.open("https://hanta.us.kg/")

    def open_github_url(self):
        webbrowser.open_new("https://github.com/exef-star/Folder-Encryption-Program")

    def MITLicense(self):
        # 创建子窗口
        child = ttk.Toplevel(self)
        child.title("MIT协议：GitHub仓库License文件全文")
        child.geometry("900x600")
        child.resizable(False, False)
        child.iconbitmap("icon/icon.ico")

        # 创建一个 Label 标签
        label = ttk.Label(child, text="MIT License\n\nCopyright (c) 2024 The hanta\n\nPermission is hereby granted, free of charge, to any person obtaining a copy\nof this software and associated documentation files (the \"Software\"), to deal\nin the Software without restriction, including without limitation the rights\nto use, copy, modify, merge, publish, distribute, sublicense, and/or sell\ncopies of the Software, and to permit persons to whom the Software is\nfurnished to do so, subject to the following conditions:\n\nThe above copyright notice and this permission notice shall be included in all\ncopies or substantial portions of the Software.\n\nTHE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR\nIMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,\nFITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE\nAUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER\nLIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,\nOUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE\nSOFTWARE.\n", font=("Consolas", 10))

        # 将 Label 标签放置到窗口中
        label.pack(pady=0)

        # 创建一个 Label 标签
        label = ttk.Label(child, text="最后一次编辑：2025年二月十号", font=("Microsoft YaHei UI Light", 10))

        # 将 Label 标签放置到窗口中
        label.pack(padx=10, pady=10)

    def about_program(self):
        # 创建子窗口
        child = ttk.Toplevel(self)
        child.title("关于本程序(❤ ω ❤)")
        child.geometry("600x400")
        child.resizable(False, False)
        child.iconbitmap("icon/icon.ico")

        #加载第一个图片
        image1 = Image.open("2022-10-07_12.jpg")  # 请确保你的图片路径正确
        photo1 = ImageTk.PhotoImage(image1)

        # 加载第二个图片
        image2 = Image.open("fepimg/whitebg.png")  # 请确保你的图片路径正确
        photo2 = ImageTk.PhotoImage(image2)

        # 加载第三个图片
        image3 = Image.open("fepimg/whitejianbian.png")  # 请确保你的图片路径正确
        photo3 = ImageTk.PhotoImage(image3)

        # 加载第四个图片
        image4 = Image.open("fepimg/image14.png")  # 请确保你的图片路径正确
        photo4 = ImageTk.PhotoImage(image4)

        # 创建Canvas
        canvas = ttk.Canvas(child, width=600, height=400)
        canvas.pack(expand=YES, fill=BOTH)

        # 将第一个图片放置在Canvas的左上角
        canvas.create_image(150, 100, image=photo1, anchor=CENTER)
        canvas.image1 = photo1  # 保持对第一个图片的引用，防止被垃圾回收

        # 将第二个图片放置在Canvas的右下角
        canvas.create_image(300, 280, image=photo2, anchor=CENTER)
        canvas.image2 = photo2  # 保持对第二个图片的引用，防止被垃圾回收

        # 将第三个图片放置在Canvas的右下角
        canvas.create_image(300, 0, image=photo3, anchor=CENTER)
        canvas.image3 = photo3  # 保持对第二个图片的引用，防止被垃圾回收

        # 将第四个图片放置在Canvas的右下角
        canvas.create_image(200, 45, image=photo4, anchor=CENTER)
        canvas.image4 = photo4  # 保持对第二个图片的引用，防止被垃圾回收

        canvas.create_text(300, 210, text="适用于 Windows 10 及 Windows 11 系统的 x64 程序（不支持arm64）\n使用 Python 3.11 编写，使用 AES 256位 加密算法对文件夹进行加密和解密。\n加密和解密的过程都不需要用户输入密码。\n\n程序使用 tkinter 库进行界面GUI设计，\n使用了 ttkbootstrap 库进行主题设置，并使用 PIL 库加载背景图片。\n\nCopyright © 2022-2025 HantaFrog\n\n版本：3.11-Alpha\n在更改、分发该程序时需遵守MIT开源协议！", font=("Microsoft YaHei UI Light", 8), fill="black")

        # 创建一个按钮
        button = ttk.Button(child, text="MIT开源协议", command=self.MITLicense, bootstyle=LIGHT)

        # 将按钮放置在Canvas上
        canvas.create_window(80, 370, window=button, anchor=CENTER)

        # 创建一个按钮
        button1 = ttk.Button(child, text="Github仓库", command=self.open_github_url, bootstyle=LIGHT)

        # 将按钮放置在Canvas上
        canvas.create_window(210, 370, window=button1, anchor=CENTER)

        # 创建一个按钮
        button2 = ttk.Button(child, text="确定", command=child.destroy, bootstyle=DARK)

        # 将按钮放置在Canvas上
        canvas.create_window(560, 370, window=button2, anchor=CENTER)

    def about_author(self):
        # 创建子窗口
        child = ttk.Toplevel(self)
        child.title("关于作者(●ˇ∀ˇ●)")
        child.geometry("600x400")
        child.resizable(False, False)
        child.iconbitmap("icon/icon.ico")

        # 加载第一个图片
        image1 = Image.open("2022-10-07_12.jpg")  # 请确保你的图片路径正确
        photo1 = ImageTk.PhotoImage(image1)

        # 加载第二个图片
        image2 = Image.open("fepimg/whitebg.png")  # 请确保你的图片路径正确
        photo2 = ImageTk.PhotoImage(image2)

        # 加载第三个图片
        image3 = Image.open("fepimg/whitejianbian.png")  # 请确保你的图片路径正确
        photo3 = ImageTk.PhotoImage(image3)

        # 加载第四个图片
        image4 = Image.open("fepimg/image14.png")  # 请确保你的图片路径正确
        photo4 = ImageTk.PhotoImage(image4)

        # 加载第四个图片
        image5 = Image.open("fepimg/68e9-c5dcd365735a8978a4e02016edf69918.png")  # 请确保你的图片路径正确
        photo5 = ImageTk.PhotoImage(image5)

        # 创建Canvas
        canvas = ttk.Canvas(child, width=600, height=400)
        canvas.pack(expand=YES, fill=BOTH)

        # 将第一个图片放置在Canvas的左上角
        canvas.create_image(150, 100, image=photo1, anchor=CENTER)
        canvas.image1 = photo1  # 保持对第一个图片的引用，防止被垃圾回收

        # 将第二个图片放置在Canvas的右下角
        canvas.create_image(300, 280, image=photo2, anchor=CENTER)
        canvas.image2 = photo2  # 保持对第二个图片的引用，防止被垃圾回收

        # 将第三个图片放置在Canvas的右下角
        canvas.create_image(300, 0, image=photo3, anchor=CENTER)
        canvas.image3 = photo3  # 保持对第二个图片的引用，防止被垃圾回收

        # 将第四个图片放置在Canvas的右下角
        canvas.create_image(200, 45, image=photo4, anchor=CENTER)
        canvas.image4 = photo4  # 保持对第二个图片的引用，防止被垃圾回收

        # 将第五个图片放置在Canvas的右下角
        canvas.create_image(100, 160, image=photo5, anchor=CENTER)
        canvas.image5 = photo5  # 保持对第二个图片的引用，防止被垃圾回收

        canvas.create_text(380, 210, text="QQ: 2667174454\n\nGithub: exef-star\n\nQQ邮箱: 2667174454@qq.com\n\n本程序仅供学习交流使用，请勿用于商业用途！", font=("Microsoft YaHei UI Light", 10), fill="black")

        # 创建一个按钮
        button = ttk.Button(child, text="确定", command=child.destroy, bootstyle=DARK)

        # 将按钮放置在Canvas上
        canvas.create_window(560, 370, window=button, anchor=CENTER)

    def create_widgets(self):
        # 加载图片
        img = Image.open("fepimg/image11.png")
        self.photo = ImageTk.PhotoImage(img)

        # 在Canvas上创建图片
        self.canvas.create_image(70, 20, anchor=ttk.NW, image=self.photo)

        # 加载图片
        img = Image.open("fepimg/image7.png")
        self.photo2 = ImageTk.PhotoImage(img)

        # 在Canvas上创建图片
        self.canvas.create_image(750, 452, anchor=ttk.NW, image=self.photo2)

        # 在Canvas上创建按钮
        self.generate_key_button = ttk.Button(self, text="生成密钥文件", command=self.generate_key, bootstyle=INFO)
        self.canvas.create_window(100, 200, window=self.generate_key_button)

        self.load_key_button = ttk.Button(self, text="选择密钥文件", command=self.load_key, bootstyle=LIGHT)
        self.canvas.create_window(100, 300, window=self.load_key_button)

        self.encrypt_button = ttk.Button(self, text="加密文件夹", command=self.encrypt_folder, bootstyle=WARNING)
        self.canvas.create_window(250, 200, window=self.encrypt_button)

        self.decrypt_button = ttk.Button(self, text="解密文件夹", command=self.decrypt_folder, bootstyle=SUCCESS)
        self.canvas.create_window(250, 300, window=self.decrypt_button)

        self.separator = ttk.Frame(self, width=4, height=130, bootstyle=LIGHT)
        self.separator.place(x=350, y=185)

        self.delete_button = ttk.Button(self, text="删除文件夹", command=self.delete_folder, bootstyle=DANGER)
        self.canvas.create_window(450, 200, window=self.delete_button)

        self.create_folder_button = ttk.Button(self, text="新建文件夹", command=self.create_folder, bootstyle=LIGHT)
        self.canvas.create_window(450, 250, window=self.create_folder_button)

        self.rename_folder_button = ttk.Button(self, text="重命名文件夹", command=self.rename_folder, bootstyle=LIGHT)
        self.canvas.create_window(450, 300, window=self.rename_folder_button)

        self.zip_and_convert_button = ttk.Button(self, text="打包文件夹为ZIP转换为TXT", command=self.zip_and_convert_to_txt, bootstyle=LIGHT)
        self.canvas.create_window(650, 200, window=self.zip_and_convert_button)

        # 将文件修改为乱码
        self.rename_filename_random_button = ttk.Button(self, text="更改文件夹内文件（不含后缀名）为乱码", command=self.rename_filename_random, bootstyle=LIGHT)
        self.canvas.create_window(702, 300, window=self.rename_filename_random_button)

        self.change_icon_button = ttk.Button(self, text="更改文件夹图标", command=self.change_folder_icon, bootstyle=LIGHT)
        self.canvas.create_window(603, 250, window=self.change_icon_button)

        self.help_button = ttk.Button(self, text="帮助文档", command=self.help, bootstyle=INFO, width=7)
        self.canvas.create_window(726, 250, window=self.help_button)

        #在右上角添加版本号
        self.version_label = ttk.Label(self, text="version: 3.11-Alpha, 2025/01/27 HantaFrog", font=("consolas", 10), foreground="green", anchor=ttk.NW)
        self.version_label.place(x=518, y=0)

        #在左下角添加image12.png
        img = Image.open("fepimg/image12.png")
        self.photo3 = ImageTk.PhotoImage(img)

        # 在Canvas上创建图片
        self.canvas.create_image(0, 460, anchor=ttk.NW, image=self.photo3)

        #添加image13.png
        img = Image.open("fepimg/image13.png")
        self.photo4 = ImageTk.PhotoImage(img)

        # 在Canvas上创建图片
        self.canvas.create_image(135, 460, anchor=ttk.NW, image=self.photo4)

    def github_way(self):
        answer = messagebox.askyesno("访问Github仓库", "是否访问Github仓库？\n需要科学上网")
        if answer:
            webbrowser.open("https://github.com/exef-star/Folder-Encryption-Program/")

    def github_another_way(self):
        answer = messagebox.askyesno("访问Github", "是否访问作者的Github？\n需要科学上网")
        if answer:
            webbrowser.open("https://github.com/exef-star")

    def help(self):
        answer = messagebox.askyesno("跳转确认", "在3.8版本及以后，作者已经在压缩包里内置了帮助文档的html文件\n是否要跳转到帮助文档？\n如果访问不了，请使用科学上网或者VPN。")
        if answer:
            webbrowser.open("https://exef-star.github.io/lighthouse/help-folder.html")

    def update_log(self):
        # 创建子窗口
        child = ttk.Toplevel(self)
        child.title("更新日志")
        child.geometry("600x400")
        child.resizable(False, False)
        child.iconbitmap("icon/icon.ico")

        # 加载第一个图片
        image1 = Image.open("2022-10-07_12.jpg")  # 请确保你的图片路径正确
        photo1 = ImageTk.PhotoImage(image1)

        # 加载第二个图片
        image2 = Image.open("fepimg/whitebg.png")  # 请确保你的图片路径正确
        photo2 = ImageTk.PhotoImage(image2)

        # 加载第三个图片
        image3 = Image.open("fepimg/whitejianbian.png")  # 请确保你的图片路径正确
        photo3 = ImageTk.PhotoImage(image3)

        # 加载第四个图片
        image4 = Image.open("fepimg/image14.png")  # 请确保你的图片路径正确
        photo4 = ImageTk.PhotoImage(image4)

        # 加载第四个图片
        image5 = Image.open("fepimg/68e9-c5dcd365735a8978a4e02016edf69918.png")  # 请确保你的图片路径正确
        photo5 = ImageTk.PhotoImage(image5)

        # 创建Canvas
        canvas = ttk.Canvas(child, width=600, height=400)
        canvas.pack(expand=YES, fill=BOTH)

        # 将第一个图片放置在Canvas的左上角
        canvas.create_image(150, 100, image=photo1, anchor=CENTER)
        canvas.image1 = photo1  # 保持对第一个图片的引用，防止被垃圾回收

        # 将第二个图片放置在Canvas的右下角
        canvas.create_image(300, 280, image=photo2, anchor=CENTER)
        canvas.image2 = photo2  # 保持对第二个图片的引用，防止被垃圾回收

        # 将第三个图片放置在Canvas的右下角
        canvas.create_image(300, 0, image=photo3, anchor=CENTER)
        canvas.image3 = photo3  # 保持对第二个图片的引用，防止被垃圾回收

        # 将第四个图片放置在Canvas的右下角
        canvas.create_image(200, 45, image=photo4, anchor=CENTER)
        canvas.image4 = photo4  # 保持对第二个图片的引用，防止被垃圾回收

        canvas.create_text(300, 210, text="版本号：v3.11-Alpha\n\n更新内容：\n1. 为所有的菜单添加\n2. 优化代码结构\n3. 整理文件目录\n4. 添加“你知道吗？”\n\n该版本为预发布版本，仅供测试，请勿使用该版本进行任何生产环境的部署。", font=("Microsoft YaHei UI Light", 8), fill="black")

        # 创建一个按钮
        button = ttk.Button(child, text="确定", command=child.destroy, bootstyle=DARK)

        # 将按钮放置在Canvas上
        canvas.create_window(560, 370, window=button, anchor=CENTER)

    def generate_key(self):
        self.key = Fernet.generate_key()
        with open("secret.key", "wb") as key_file:
            key_file.write(self.key)
        messagebox.showinfo("密钥生成(*^_^*)", "密钥已生成并保存为 secret.key")
        print(Fore.GREEN + "[SUCCESS] Successfully generated key file and saved it on local disk, key: ", self.key)

    def load_key(self):
        file_path = filedialog.askopenfilename(filetypes=[("Key Files", "*.key")])
        if file_path:
            try:
                with open(file_path, "rb") as key_file:
                    self.key = key_file.read()
                messagebox.showinfo("密钥加载o(*￣▽￣*)o", "密钥已加载")
                print(Fore.GREEN + "[SUCCESS] Successfully imported key: ", self.key)
            except Exception as e:
                messagebox.showerror("错误（＞人＜；）", f"加载密钥时出错: {e}")
                print(Fore.RED + f"[ERROR] Failed to import key file: {e}")

    def process_folder(self, action):
        if not self.key:
            messagebox.showwarning("警告(⊙o⊙)？", "请先生成或加载密钥")
            print(Fore.YELLOW + "[WARNING] Please create or import the key file to continue, Self.key is None")
            return

        folder_path = filedialog.askdirectory()
        if not folder_path:
            return

        fernet = Fernet(self.key)
        try:
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    with open(file_path, "rb") as f:
                        data = f.read()
                    if action == "encrypt":
                        processed_data = fernet.encrypt(data)
                    else:
                        processed_data = fernet.decrypt(data)
                    with open(file_path, "wb") as f:
                        f.write(processed_data)
            messagebox.showinfo(f"{action.capitalize()}完成(￣▽￣)", f"文件夹已{action}")
            print(Fore.GREEN + f"[SUCCESS] Successfully {action}ed folder: {folder_path}")
        except InvalidToken:
            messagebox.showerror("错误ヽ(*。>Д<)o゜", "解密失败，密钥可能不正确")
            print(Fore.RED + "[ERROR] Failed to decrypt folder, invalid key")
        except Exception as e:
            messagebox.showerror("错误＞﹏＜", f"{action.capitalize()}文件夹时出错: {e}")
            print(Fore.RED + f"[ERROR] Failed to {action} folder: {folder_path}, error: {e}")

    def encrypt_folder(self):
        self.process_folder("encrypt")

    def decrypt_folder(self):
        self.process_folder("decrypt")

    def delete_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            try:
                for root, dirs, files in os.walk(folder_path, topdown=False):
                    for name in files:
                        os.remove(os.path.join(root, name))
                    for name in dirs:
                        os.rmdir(os.path.join(root, name))
                os.rmdir(folder_path)
                messagebox.showinfo("删除完成", "文件夹已删除")
                print(Fore.GREEN + f"[SUCCESS] Successfully deleted folder: {folder_path}")
            except Exception as e:
                messagebox.showerror("错误", f"删除文件夹时出错: {e}")
                print(Fore.RED + f"[ERROR] Failed to delete folder: {folder_path}, error: {e}")

    def create_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            new_folder_name = simpledialog.askstring("新建文件夹", "请输入新文件夹的名称:")
            if new_folder_name:
                new_folder_path = os.path.join(folder_path, new_folder_name)
                try:
                    os.mkdir(new_folder_path)
                    messagebox.showinfo("新建文件夹", f"文件夹 '{new_folder_name}' 已创建")
                    print(Fore.GREEN +  f"[SUCCESS] Successfully created folder: {new_folder_path}")
                except Exception as e:
                    messagebox.showerror("错误＞︿＜", f"创建文件夹时出错: {e}")
                    print(Fore.RED + f"[ERROR] Failed to create folder: {new_folder_path}, error: {e}")

    def rename_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            new_folder_name = simpledialog.askstring("重命名文件夹", "请输入新文件夹的名称:")
            print(Fore.BLUE + "[INFO]New folder Path: ", folder_path)
            if new_folder_name:
                new_folder_path = os.path.join(os.path.dirname(folder_path), new_folder_name)
                try:
                    os.rename(folder_path, new_folder_path)
                    messagebox.showinfo("重命名文件夹", f"文件夹已重命名为 '{new_folder_name}'")
                    print(Fore.GREEN + f"[SUCCESS] Successfully renamed folder: {folder_path} to {new_folder_path}")
                except Exception as e:
                    messagebox.showerror("错误≧ ﹏ ≦", f"重命名文件夹时出错: {e}")
                    print(Fore.RED + f"[ERROR] Failed to rename folder: {folder_path} to {new_folder_path}, error: {e}")

    def zip_and_convert_to_txt(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            try:
                # 创建ZIP文件
                zip_path = os.path.join(os.path.dirname(folder_path), os.path.basename(folder_path) + ".zip")
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, _, files in os.walk(folder_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, folder_path)
                            zipf.write(file_path, arcname)

                # 将ZIP文件转换为Base64编码的文本文件
                with open(zip_path, "rb") as zip_file:
                    zip_data = zip_file.read()
                    base64_data = base64.b64encode(zip_data).decode('utf-8')

                txt_path = os.path.join(os.path.dirname(folder_path), os.path.basename(folder_path) + ".txt")
                with open(txt_path, "w") as txt_file:
                    txt_file.write(base64_data)

                messagebox.showinfo("打包完成", f"文件夹已打包为ZIP并转换为TXT文件，保存路径为: {txt_path}")
                print(Fore.GREEN + f"[SUCCESS] Successfully zipped folder: {folder_path} and converted it to txt file: {txt_path}")
            except Exception as e:
                messagebox.showerror("错误இ௰இ", f"打包文件夹时出错: {e}")
                print(Fore.RED + f"[ERROR] Failed to zip folder: {folder_path}, error: {e}")

    def change_folder_icon(self):
        folder_path = filedialog.askdirectory()
        if not folder_path:
            return

        icon_path = filedialog.askopenfilename(filetypes=[("Icon Files", "*.ico")])
        if not icon_path:
            return

        try:
            self.set_folder_icon(folder_path, icon_path)
            messagebox.showinfo("成功(✿◡‿◡)", "文件夹图标已更改")
            print(Fore.GREEN + f"[SUCCESS] Successfully changed folder icon: {folder_path} to {icon_path}")
        except Exception as e:
            messagebox.showerror("错误::>_<::", f"更改文件夹图标时出错: {e}")
            print(Fore.RED + f"[ERROR] Failed to change folder icon: {folder_path} to {icon_path}, error: {e}")

    def set_folder_icon(self, folder_path, icon_path):
        desktop_ini_path = os.path.join(folder_path, "desktop.ini")

        # 创建desktop.ini文件
        with open(desktop_ini_path, "w") as ini_file:
            ini_file.write("[.ShellClassInfo]\n")
            ini_file.write(f"IconFile={icon_path}\n")
            ini_file.write("IconIndex=0\n")

        print(Fore.GREEN + f"[SUCCESS] Successfully created desktop.ini file: {desktop_ini_path}")

        # 设置文件属性
        os.chmod(desktop_ini_path, 0o644)
        win32api.SetFileAttributes(desktop_ini_path, win32con.FILE_ATTRIBUTE_HIDDEN | win32con.FILE_ATTRIBUTE_SYSTEM)
        win32api.SetFileAttributes(folder_path, win32con.FILE_ATTRIBUTE_SYSTEM)
        print(Fore.GREEN + f"[SUCCESS] Successfully set file attributes for: {desktop_ini_path} and {folder_path}")

        # 刷新文件夹图标
        shell32 = ctypes.windll.shell32
        shell32.SHChangeNotify(shell32.SHCNE_ASSOCCHANGED, shell32.SHCNF_IDLIST, None, None)
        print(Fore.GREEN + "[SUCCESS] Successfully refreshed folder icon")

    def rename_filename_random(self):
        folder_path = filedialog.askdirectory()
        if not folder_path:
            return

        try:
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    new_file_name = "".join(random.choices(string.ascii_letters + string.digits, k=16)) + os.path.splitext(file)[1]
                    new_file_path = os.path.join(root, new_file_name)
                    os.rename(file_path, new_file_path)
            messagebox.showinfo("完成(●'◡'●)", "文件名已更改为随机字符串")
            print(Fore.GREEN + f"Successfully renamed file names in folder: {folder_path}")
        except Exception as e:
            messagebox.showerror("错误(๑•̀ㅂ•́)و✧", f"更改文件名时出错: {e}")
            print(Fore.GREEN + f"Failed to rename file names in folder: {folder_path}, error: {e}")

    def view_key(self):
        if not self.key:
            messagebox.showwarning("警告(⊙o⊙)？", "请先生成或加载密钥")
            print(Fore.YELLOW + "[WARNING] Please create or import the key file to continue, Self.key is None")
            return

        messagebox.showinfo("密钥查看", f"密钥: {self.key.decode('utf-8')}")
        print(Fore.GREEN + f"[SUCCESS] Key: {self.key.decode('utf-8')}")

    def generate_Debug(self):
        messagebox.showwarning("生成日志文件警告(＃°Д°)", f"只有在程序出现错误时才需要生成日志文件，单击确定开始生成Debug.txt文件")
        # 创建Debug.txt文件
        debug_path = os.path.join(os.path.expanduser("~"), "Desktop", "Debug.txt")
        with open(debug_path, "w") as debug_file:
            debug_file.write("Debug.txt")
            debug_file.write(f"\nTime: {datetime.datetime.now()}")
            debug_file.write(f"\nPython version: {sys.version}")
            debug_file.write(f"\nFolderEncryptor version: {FolderEncryptorVersion}")
            debug_file.write(f"\nFolderEncryptor path: {os.path.abspath(__file__)}")
            debug_file.write(f"\nProcessor architecture: {platform.architecture()}")
            debug_file.write(f"\nSystem version: {platform.version()}")
            debug_file.write(f"\nSystem platform: {platform.platform()}")
            debug_file.write(f"\nSystem machine: {platform.machine()}")
            debug_file.write(f"\nSystem processor: {platform.processor()}")
            debug_file.write(f"\nCurrent user: {getpass.getuser()}")
            debug_file.write(f"\nNumber of loaded libraries: {len(sys.modules)}")
            #打印当前程序占用CPU占比
            debug_file.write(f"\nCPU usage: {psutil.cpu_percent()}%")
            debug_file.write(f"\n------------------------------------------")
            #打印控制台错误
            debug_file.write(f"\nConsole output:\n{traceback.format_exc()}")
            debug_file.write(f"\n该文件为非正式版，请勿用在任何重要生产环境中，如出现问题，概不负责！")
        messagebox.showinfo("Debug文件创建成功", f"Debug文件已创建: {debug_path}")
        print(Fore.GREEN + f"[SUCCESS] Debug file created: {debug_path}")

if __name__ == "__main__":
    app = FolderEncryptor()
    app.iconbitmap("icon/icon.ico")
    app.mainloop()

print("-------------------------------")
print("success run the program")
