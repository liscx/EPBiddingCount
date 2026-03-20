import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
import datetime
import sys

# --- 关键补丁：确保打包后能找到 main.py ---
if hasattr(sys, '_MEIPASS'):
    sys.path.append(sys._MEIPASS)

try:
    import main  # 引入逻辑文件
except ImportError:
    # 这里的提示是为了方便你在开发环境调试
    print("错误：未找到 main.py 文件，请确保它与 GUI.py 在同一目录。")

# 资源路径处理（用于单文件打包后的图标加载）
def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# 主题配色
THEME_DATA = {
    "Dark": {"bg": "#0d1117", "card": "#161b22", "border": "#30363d", "text": "#f0f6fc", "input_bg": "#010409",
             "icon_color": "#8b949e"},
    "Light": {"bg": "#ffffff", "card": "#f6f8fa", "border": "#d0d7de", "text": "#1f2328", "input_bg": "#ffffff",
              "icon_color": "#6e7781"}
}

class BiddingCountGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("BiddingCount v1.0")
        self.geometry("500x650")
        self.resizable(False, False)

        icon_path = get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except:
                pass

        self.appearance_mode = "Dark"
        ctk.set_appearance_mode(self.appearance_mode)

        self.setup_ui()
        self.apply_theme_styles()

    def setup_ui(self):
        # Header
        self.header = ctk.CTkFrame(self, fg_color="transparent", height=40)
        self.header.pack(fill="x", padx=20, pady=(20, 10))
        ctk.CTkLabel(self.header, text="BiddingCount", font=("Inter", 20, "bold")).pack(side="left")

        self.theme_btn = ctk.CTkButton(self.header, text="🌙", width=35, height=35, fg_color="transparent",
                                       command=self.toggle_theme)
        self.theme_btn.pack(side="right")

        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=20)

        # 1. 输入文件卡片
        self.in_card = self.create_card("SOURCE EXCEL FILE")
        self.in_entry = ctk.CTkEntry(self.in_card, height=32, border_width=1)
        # 默认寻找当前目录下的汇总表
        if os.path.exists("专区信息汇总表.xlsx"):
            self.in_entry.insert(0, os.path.abspath("专区信息汇总表.xlsx"))
            # 如果默认文件存在，同步设置输出路径
            self.out_entry_default = os.path.dirname(os.path.abspath("专区信息汇总表.xlsx"))
        else:
            self.out_entry_default = ""

        self.in_entry.pack(side="left", expand=True, fill="x", padx=(10, 5), pady=10)
        ctk.CTkButton(self.in_card, text="Browse", width=70, command=lambda: self.browse('file')).pack(side="right",
                                                                                                       padx=(0, 10))

        # 2. 输出路径卡片
        self.out_card = self.create_card("OUTPUT DIRECTORY")
        self.out_entry = ctk.CTkEntry(self.out_card, height=32, border_width=1)
        if self.out_entry_default:
            self.out_entry.insert(0, self.out_entry_default)
        self.out_entry.pack(side="left", expand=True, fill="x", padx=(10, 5), pady=10)
        ctk.CTkButton(self.out_card, text="Select", width=70, command=lambda: self.browse('dir')).pack(side="right",
                                                                                                       padx=(0, 10))

        # 3. 日志框标签
        ctk.CTkLabel(self.main_container, text="EXECUTION LOG", font=("Inter", 11, "bold"), text_color="#6e7781").pack(
            anchor="w", padx=5, pady=(15, 0))

        # 4. 日志输出
        self.log_output = ctk.CTkTextbox(self.main_container, corner_radius=8, border_width=1, font=("Consolas", 13))
        self.log_output.pack(fill="both", expand=True, pady=(5, 15))

        # 5. 运行按钮
        self.run_btn = ctk.CTkButton(self, text="Run Analysis", font=("Inter", 15, "bold"), height=45,
                                     command=self.start_process)
        self.run_btn.pack(fill="x", padx=20, pady=(0, 20))

    def create_card(self, title):
        card = ctk.CTkFrame(self.main_container, corner_radius=10, border_width=1)
        card.pack(fill="x", pady=6)
        ctk.CTkLabel(card, text=title, font=("Inter", 10, "bold"), text_color="#6e7781").pack(anchor="w", padx=12,
                                                                                              pady=(6, 0))
        return card

    def apply_theme_styles(self):
        c = THEME_DATA[self.appearance_mode]
        self.configure(fg_color=c["bg"])
        self.theme_btn.configure(text="🌙" if self.appearance_mode == "Dark" else "☀️", text_color=c["icon_color"])
        self.run_btn.configure(fg_color="#238636" if self.appearance_mode == "Dark" else "#2ea44f")
        self.log_output.configure(fg_color=c["card"], border_color=c["border"], text_color="#8b949e")

    def toggle_theme(self):
        self.appearance_mode = "Light" if self.appearance_mode == "Dark" else "Dark"
        ctk.set_appearance_mode(self.appearance_mode)
        self.apply_theme_styles()

    def browse(self, mode):
        if mode == 'file':
            f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
            if f:
                self.in_entry.delete(0, "end")
                self.in_entry.insert(0, f)
                # 核心逻辑：选完文件，默认输出目录指向该文件所在文件夹（实现原地修改）
                self.out_entry.delete(0, "end")
                self.out_entry.insert(0, os.path.dirname(os.path.abspath(f)))
        else:
            d = filedialog.askdirectory()
            if d:
                self.out_entry.delete(0, "end")
                self.out_entry.insert(0, d)

    def log(self, text):
        now = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_output.insert("end", f"[{now}] {text}\n")
        self.log_output.see("end")

    def start_process(self):
        in_path = self.in_entry.get()
        out_dir = self.out_entry.get()
        if not in_path or not os.path.exists(in_path):
            messagebox.showerror("Error", "输入文件无效")
            return
        if not out_dir:
            messagebox.showerror("Error", "请选择输出路径")
            return

        self.run_btn.configure(state="disabled", text="Processing...")
        self.log_output.delete("1.0", "end")
        threading.Thread(target=self.work_logic, args=(in_path, out_dir), daemon=True).start()

    def work_logic(self, in_path, out_dir):
        try:
            # 这里的 process_excel 是 main.py 里的函数
            success_file = main.process_excel(in_path, out_dir, self.log)

            if success_file:
                messagebox.showinfo("BiddingCount", f"任务成功完成！\n文件处理路径：\n{success_file}")
        except Exception as e:
            self.log(f"GUI调用错误: {str(e)}")
        finally:
            self.run_btn.configure(state="normal", text="Run Analysis")

if __name__ == "__main__":
    app = BiddingCountGUI()
    app.mainloop()