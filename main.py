import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from cryptography.fernet import Fernet, InvalidToken
import os
import zipfile
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
import webbrowser
import base64
import win32api
import win32con

class FolderEncryptor(ttk.Window):
    def __init__(self):
        super().__init__(themename="lumen")
        self.title("文件夹加密程序&By HantaFrog (*^▽^*)")
        self.geometry("973x480")
        self.resizable(False, False)
        self.key = None

        # 设置背景图片
        self.set_background_image()

        # 创建顶部导航栏
        self.create_menu_bar()

        self.create_widgets()

    def set_background_image(self):
        # 加载图片
        imgpath = 'background1.png'
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

        # 创建文件菜单
        file_menu = ttk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="生成密钥文件", command=self.generate_key)
        file_menu.add_command(label="选择密钥文件", command=self.load_key)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.quit)

        # 创建操作菜单
        action_menu = ttk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="操作", menu=action_menu)
        action_menu.add_command(label="加密文件夹", command=self.encrypt_folder)
        action_menu.add_command(label="解密文件夹", command=self.decrypt_folder)
        action_menu.add_command(label="删除文件夹", command=self.delete_folder)
        action_menu.add_command(label="新建文件夹", command=self.create_folder)
        action_menu.add_command(label="重命名文件夹", command=self.rename_folder)
        action_menu.add_command(label="打包文件夹为ZIP并转换为TXT", command=self.zip_and_convert_to_txt)
        action_menu.add_command(label="更改文件夹图标", command=self.change_folder_icon)

        # 创建关于菜单
        about_menu = ttk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="关于", menu=about_menu)
        about_menu.add_command(label="关于本程序", command=self.about_program)
        about_menu.add_command(label="关于作者", command=self.about_author)

    def about_program(self):
        messagebox.showinfo("关于本程序(❤ ω ❤)", "本程序使用 Python 3.x 编写，使用 Fernet 加密算法对文件夹进行加密和解密。\n\n加密和解密的过程都不需要用户输入密码，加密后的文件名后缀会被修改为.enc，解密后的文件名后缀会被修改为原来的后缀。\n\n程序使用 tkinter 库进行界面GUI设计，\n使用了 ttkbootstrap 库进行主题设置，并使用 PIL 库加载背景图片。")

    def about_author(self):
        messagebox.showinfo("关于作者(●ˇ∀ˇ●)", "作者：HantaFrog\n\n邮箱：<2667174454@qq.com>\n\nGitHub：https://github.com/exef-star\n\n欢迎访问我的博客鸭：\n\nhttps://exef-star.github.io/")

    def create_widgets(self):
        # 加载图片
        img = Image.open("image6.png")
        self.photo = ImageTk.PhotoImage(img)

        # 在Canvas上创建图片
        self.canvas.create_image(-150, 0, anchor=ttk.NW, image=self.photo)

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

        self.change_icon_button = ttk.Button(self, text="更改文件夹图标", command=self.change_folder_icon, bootstyle=LIGHT)
        self.canvas.create_window(603, 300, window=self.change_icon_button)

        self.help_button = ttk.Button(self, text="帮助", command=self.help, bootstyle=INFO)
        self.canvas.create_window(740, 300, window=self.help_button)

    def help(self):
        webbrowser.open("https://exef-star.github.io/lighthouse/help-folder.html")

    def generate_key(self):
        self.key = Fernet.generate_key()
        with open("secret.key", "wb") as key_file:
            key_file.write(self.key)
        messagebox.showinfo("密钥生成(*^_^*)", "密钥已生成并保存为 secret.key")
        print("Successfully generated key file and saved it on local disk, key: ", self.key)

    def load_key(self):
        file_path = filedialog.askopenfilename(filetypes=[("Key Files", "*.key")])
        if file_path:
            try:
                with open(file_path, "rb") as key_file:
                    self.key = key_file.read()
                messagebox.showinfo("密钥加载o(*￣▽￣*)o", "密钥已加载")
                print("Successfully imported key: ", self.key)
            except Exception as e:
                messagebox.showerror("错误（＞人＜；）", f"加载密钥时出错: {e}")
                print(f"Failed to import key file: {e}")

    def process_folder(self, action):
        if not self.key:
            messagebox.showwarning("警告(⊙o⊙)？", "请先生成或加载密钥")
            print("Please create or import the key file to continue, Self.key is None")
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
            print(f"Successfully {action}ed folder: {folder_path}")
        except InvalidToken:
            messagebox.showerror("错误ヽ(*。>Д<)o゜", "解密失败，密钥可能不正确")
            print("Failed to decrypt folder, invalid key")
        except Exception as e:
            messagebox.showerror("错误＞﹏＜", f"{action.capitalize()}文件夹时出错: {e}")
            print(f"Failed to {action} folder: {folder_path}, error: {e}")

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
                print(f"Successfully deleted folder: {folder_path}")
            except Exception as e:
                messagebox.showerror("错误", f"删除文件夹时出错: {e}")
                print(f"Failed to delete folder: {folder_path}, error: {e}")

    def create_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            new_folder_name = simpledialog.askstring("新建文件夹", "请输入新文件夹的名称:")
            if new_folder_name:
                new_folder_path = os.path.join(folder_path, new_folder_name)
                try:
                    os.mkdir(new_folder_path)
                    messagebox.showinfo("新建文件夹", f"文件夹 '{new_folder_name}' 已创建")
                except Exception as e:
                    messagebox.showerror("错误＞︿＜", f"创建文件夹时出错: {e}")

    def rename_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            new_folder_name = simpledialog.askstring("重命名文件夹", "请输入新文件夹的名称:")
            print("New folder Path: ", folder_path)
            if new_folder_name:
                new_folder_path = os.path.join(os.path.dirname(folder_path), new_folder_name)
                try:
                    os.rename(folder_path, new_folder_path)
                    messagebox.showinfo("重命名文件夹", f"文件夹已重命名为 '{new_folder_name}'")
                    print(f"Successfully renamed folder: {folder_path} to {new_folder_path}")
                except Exception as e:
                    messagebox.showerror("错误≧ ﹏ ≦", f"重命名文件夹时出错: {e}")
                    print(f"Failed to rename folder: {folder_path} to {new_folder_path}, error: {e}")

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
                print(f"Successfully zipped folder: {folder_path} and converted it to txt file: {txt_path}")
            except Exception as e:
                messagebox.showerror("错误இ௰இ", f"打包文件夹时出错: {e}")
                print(f"Failed to zip folder: {folder_path}, error: {e}")

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
            print(f"Successfully changed folder icon: {folder_path} to {icon_path}")
        except Exception as e:
            messagebox.showerror("错误::>_<::", f"更改文件夹图标时出错: {e}")
            print(f"Failed to change folder icon: {folder_path} to {icon_path}, error: {e}")

    def set_folder_icon(self, folder_path, icon_path):
        desktop_ini_path = os.path.join(folder_path, "desktop.ini")

        # 创建desktop.ini文件
        with open(desktop_ini_path, "w") as ini_file:
            ini_file.write("[.ShellClassInfo]\n")
            ini_file.write(f"IconFile={icon_path}\n")
            ini_file.write("IconIndex=0\n")

        print(f"Successfully created desktop.ini file: {desktop_ini_path}")

        # 设置文件属性
        os.chmod(desktop_ini_path, 0o644)
        win32api.SetFileAttributes(desktop_ini_path, win32con.FILE_ATTRIBUTE_HIDDEN | win32con.FILE_ATTRIBUTE_SYSTEM)
        win32api.SetFileAttributes(folder_path, win32con.FILE_ATTRIBUTE_SYSTEM)
        print(f"Successfully set file attributes for: {desktop_ini_path} and {folder_path}")

        # 刷新文件夹图标
        shell32 = ctypes.windll.shell32
        shell32.SHChangeNotify(shell32.SHCNE_ASSOCCHANGED, shell32.SHCNF_IDLIST, None, None)
        print("Successfully refreshed folder icon")

if __name__ == "__main__":
    app = FolderEncryptor()
    app.iconbitmap("icon.ico")
    app.mainloop()
