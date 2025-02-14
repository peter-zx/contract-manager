# 你的主代码文件
import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

# 自定义带复选框的滚动列表组件
class CheckboxListbox(tk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner_frame = tk.Frame(self.canvas)
        
        # 配置画布和滚动条
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # 创建可滚动区域
        self.canvas.create_window((0,0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        # 存储组件和状态
        self.checkboxes = []
        self.labels = []
        self.vars = []
        
    def add_items(self, items):
        """加载带复选框的列表项"""
        # 清除旧项
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        self.checkboxes.clear()
        self.labels.clear()
        self.vars.clear()
        
        # 添加新项
        for idx, text in enumerate(items):
            var = tk.BooleanVar(value=True)  # 默认选中
            self.vars.append(var)
            
            # 创建复选框
            check = ttk.Checkbutton(
                self.inner_frame,
                variable=var,
                style="Custom.TCheckbutton",
                command=lambda i=idx: self._update_label(i)
            )
            check.grid(row=idx, column=0, sticky="w", padx=5)
            
            # 创建文件标签
            label = ttk.Label(
                self.inner_frame,
                text=text,
                padding=(0, 2),
                style="Custom.TLabel"
            )
            label.grid(row=idx, column=1, sticky="w")
            label.bind("<Double-1>", lambda e, i=idx: self._toggle_item(i))
            
            self.checkboxes.append(check)
            self.labels.append(label)
            
    def _toggle_item(self, index):
        """双击切换选中状态"""
        current = self.vars[index].get()
        self.vars[index].set(not current)
        self._update_label(index)
        
    def _update_label(self, index):
        """更新标签样式"""
        if self.vars[index].get():
            self.labels[index].configure(style="Selected.TLabel")
        else:
            self.labels[index].configure(style="Unselected.TLabel")
            
    def get_selected_indices(self):
        """获取选中项的索引"""
        return [i for i, var in enumerate(self.vars) if var.get()]

# 主应用程序类
class ContractApp:
    def __init__(self, master):
        self.master = master
        master.title("智能合同生成工具 v3.0")
        master.geometry("1000x800")
        
        # 初始化路径
        self.desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.default_output = os.path.join(self.desktop_path, "合同生成结果")
        self.available_files = []
        
        # 配置样式
        self._configure_styles()
        
        # 创建界面组件
        self._create_widgets()
        
    def _configure_styles(self):
        """配置GUI样式"""
        style = ttk.Style()
        
        # 复选框样式
        style.configure("Custom.TCheckbutton", padding=5)
        style.map("Custom.TCheckbutton",
            background=[("active", "white")],
            indicatormargin=[("pressed", 5), ("!pressed", 3)]
        )
        
        # 标签样式
        style.configure("Selected.TLabel", foreground="black")
        style.configure("Unselected.TLabel", foreground="gray60")
        
    def _create_widgets(self):
        """创建界面组件"""
        # 源文件夹选择
        ttk.Label(self.master, text="源文件夹路径:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.src_entry = ttk.Entry(self.master, width=70)
        self.src_entry.grid(row=1, column=0, padx=10, sticky="ew")
        ttk.Button(self.master, text="选择文件夹", command=self._select_source).grid(row=1, column=1, padx=10)

        # 文件选择框
        self.file_frame = ttk.LabelFrame(self.master, text="选择要复制的文件（双击切换选择）")
        self.file_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        
        self.file_list = CheckboxListbox(self.file_frame)
        self.file_list.pack(fill="both", expand=True, padx=5, pady=5)

        # 目标路径选择
        ttk.Label(self.master, text="输出文件夹路径:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.dest_entry = ttk.Entry(self.master, width=70)
        self.dest_entry.insert(0, self.default_output)
        self.dest_entry.grid(row=4, column=0, padx=10, sticky="ew")
        ttk.Button(self.master, text="选择位置", command=self._select_destination).grid(row=4, column=1, padx=10)

        # 控制按钮
        self.run_btn = ttk.Button(self.master, text="开始生成", command=self._start_process)
        self.run_btn.grid(row=5, column=0, columnspan=2, pady=10)

        # 日志显示区
        self.log_text = tk.Text(self.master, height=15, state='disabled')
        self.log_text.grid(row=6, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

        # 布局配置
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(6, weight=1)

    def _select_source(self):
        """选择源文件夹"""
        path = filedialog.askdirectory(title="选择包含names.txt的文件夹")
        if path:
            self.src_entry.delete(0, tk.END)
            self.src_entry.insert(0, path)
            self._load_files(path)

    def _load_files(self, path):
        """加载源文件夹文件"""
        try:
            all_files = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
            self.available_files = [f for f in all_files if f.lower() != "names.txt"]
            self.file_list.add_items(self.available_files)
            self._log(f"找到 {len(self.available_files)} 个可用文件")
        except Exception as e:
            self._log(f"错误：加载文件失败 - {str(e)}")

    def _select_destination(self):
        """选择目标位置"""
        path = filedialog.askdirectory(title="选择保存结果的文件夹")
        if path:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, path)

    def _log(self, message):
        """记录日志信息"""
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, f"• {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')
        self.master.update_idletasks()

    def _validate(self):
        """验证输入有效性"""
        errors = []
        src_path = self.src_entry.get()
        dest_path = self.dest_entry.get()

        if not src_path:
            errors.append("请选择源文件夹")
        elif not os.path.exists(src_path):
            errors.append("源文件夹不存在")
        elif not os.path.isfile(os.path.join(src_path, "names.txt")):
            errors.append("源文件夹必须包含names.txt文件")

        if not dest_path:
            errors.append("请选择输出位置")

        if not self.file_list.get_selected_indices():
            errors.append("请至少选择一个要复制的文件")

        return errors

    def _process_files(self):
        """处理文件的核心逻辑"""
        src_path = self.src_entry.get()
        dest_path = self.dest_entry.get()
        selected_indices = self.file_list.get_selected_indices()
        selected_files = [self.available_files[i] for i in selected_indices]

        # 读取姓名列表
        try:
            with open(os.path.join(src_path, "names.txt"), "r", encoding="utf-8") as f:
                names = [line.strip() for line in f if line.strip()]
        except Exception as e:
            messagebox.showerror("错误", f"读取names.txt失败：{str(e)}")
            return

        total = len(names)
        success_count = 0

        self._log(f"开始处理，共 {total} 人...")
        self._log(f"将复制以下文件：{', '.join(selected_files)}")

        # 处理每个人员
        for idx, name in enumerate(names, 1):
            # 清洗姓名
            clean_name = ''.join(c for c in name if c not in r'<>:"/\|?*').strip()
            if not clean_name:
                self._log(f"({idx}/{total}) 跳过无效姓名：{name}")
                continue

            # 创建个人文件夹
            person_dir = os.path.join(dest_path, clean_name)
            try:
                os.makedirs(person_dir, exist_ok=True)
            except Exception as e:
                self._log(f"({idx}/{total}) 创建文件夹失败：{clean_name} - {str(e)}")
                continue

            # 复制文件
            copied = 0
            try:
                for filename in selected_files:
                    src = os.path.join(src_path, filename)
                    dst = os.path.join(person_dir, filename)

                    if not os.path.exists(src):
                        self._log(f"({idx}/{total}) 文件不存在：{filename}")
                        continue

                    if not os.path.exists(dst):
                        shutil.copy2(src, dst)
                        copied += 1
                        self._log(f"({idx}/{total}) 已复制：{filename} → {clean_name}")
                    else:
                        self._log(f"({idx}/{total}) 文件已存在：{filename}")

                if copied == len(selected_files):
                    success_count += 1
                    status = "✓ 成功"
                else:
                    status = "⚠ 部分成功"
                
                self._log(f"({idx}/{total}) {clean_name.ljust(10)} {status}")

            except Exception as e:
                self._log(f"({idx}/{total}) 处理失败：{clean_name} - {str(e)}")

        # 完成处理
        self._log(f"\n处理完成！成功处理 {success_count}/{total} 人")
        messagebox.showinfo("完成", f"处理完成！\n成功处理人数：{success_count}/{total}")
        self.run_btn.config(state="normal")

    def _start_process(self):
        """启动处理线程"""
        # 验证输入
        errors = self._validate()
        if errors:
            messagebox.showerror("输入错误", "\n".join(errors))
            return

        # 初始化界面状态
        self.run_btn.config(state="disabled")
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        
        # 启动处理线程
        threading.Thread(target=self._process_files, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = ContractApp(root)
    root.mainloop()
