import sys
import os
import warnings
import time
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from pywinauto.application import Application
from tfls_pdt import show_pdt_dialog
from tfls_metadata import show_metadata_setup_dialog
from tfls_init_pgm import run_initial_pgm
from pywinauto.keyboard import send_keys

# 忽略 UserWarning 警告
warnings.simplefilter("ignore", UserWarning)
# 设置 COM 线程模型
sys.coinit_flags = 2


class SearchableCombobox(ttk.Combobox):
    """支持搜索功能的Combobox"""
    def __init__(self, parent, **kwargs):
        # 设置默认值
        kwargs.setdefault('state', 'normal')
        super().__init__(parent, **kwargs)
        
        # 存储所有选项
        self.all_values = []
        
        # 绑定事件
        self.bind('<KeyRelease>', self._on_key_release)
        self.bind('<FocusOut>', self._on_focus_out)
        self.bind('<Return>', self._on_return)
    
    def set_values(self, values):
        """设置所有选项"""
        self.all_values = sorted(values) if values else []
        self['values'] = self.all_values
    
    def _on_key_release(self, event):
        """当用户输入时，过滤选项"""
        # 如果下拉框被禁用，不处理
        if self['state'] == 'disabled':
            return
        
        if event.keysym in ['Up', 'Down', 'Return', 'Tab']:
            return
        
        current_value = self.get().lower()
        
        if current_value:
            # 过滤匹配的选项（不区分大小写）
            filtered = [v for v in self.all_values if current_value in v.lower()]
            self['values'] = filtered
        else:
            # 如果输入为空，显示所有选项
            self['values'] = self.all_values
        
        # 如果只有一个匹配项且完全匹配，自动选择
        if len(self['values']) == 1 and self['values'][0].lower() == current_value:
            self.set(self['values'][0])
            # 触发选择事件
            self.event_generate('<<ComboboxSelected>>')
    
    def _on_focus_out(self, event):
        """失去焦点时，如果输入的值不在列表中，清空或恢复"""
        # 如果下拉框被禁用，不处理
        if self['state'] == 'disabled':
            return
        
        current_value = self.get()
        if current_value and current_value not in self.all_values:
            # 尝试找到最匹配的项（以输入值开头的）
            current_lower = current_value.lower()
            matches = [v for v in self.all_values if v.lower().startswith(current_lower)]
            if matches:
                self.set(matches[0])
                # 触发选择事件
                self.event_generate('<<ComboboxSelected>>')
            else:
                # 如果没有匹配，清空
                self.set('')
                # 恢复所有选项
                self['values'] = self.all_values
    
    def _on_return(self, event):
        """按回车键时，选择第一个匹配项"""
        # 如果下拉框被禁用，不处理
        if self['state'] == 'disabled':
            return 'break'
        
        current_value = self.get().lower()
        if current_value and self['values']:
            # 选择第一个匹配项
            self.set(self['values'][0])
            # 触发选择事件
            self.event_generate('<<ComboboxSelected>>')
        return 'break'


class SASEGGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SASEG Autoexec")
        self.root.geometry("1200x700")
        self.root.resizable(True, True)
        
        # Z盘路径
        self.z_drive = "Z:\\"
        
        # 存储6个下拉框
        self.comboboxes = []
        # 存储当前选中的路径
        self.selected_paths = [""] * 6
        
        # 存储网格框架和标签
        self.grid_frame = None
        self.grid_labels = {}  # 存储列标题和内容标签
        self.current_subfolders = []  # 当前网格列顺序（用于快捷跳转）
        
        # 创建界面
        self.create_widgets()
        
        # 初始化第一个下拉框
        self.refresh_first_dropdown()
        # 测试阶段：默认选中 projects、HRS2129、HRS2129_test、csr_01
        self.root.after(100, self._apply_test_defaults)
    
    def create_widgets(self):
        # 主布局：左侧导航 + 右侧内容
        content_row = tk.Frame(self.root)
        content_row.pack(fill=tk.BOTH, expand=True)
        
        # ========== 左侧导航栏（仿 Dizal 风格）：浅灰背景 + 文字菜单 + 当前项左侧蓝条 ==========
        SIDEBAR_BG = "#e5e5e5"
        SIDEBAR_SEL_BG = "#d5d5d5"
        SIDEBAR_LEFT_BAR = "#5B9BD5"
        
        self.sidebar = tk.Frame(content_row, bg=SIDEBAR_BG, width=140)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)
        
        self.page_names = ["主页", "aCRF", "SDTM", "ADaM", "TFLs", "M5", "pm"]
        self.sidebar_items = []  # [(frame, left_bar, text_lbl, page_id), ...]
        self.current_page = "主页"
        
        def make_sidebar_item(parent, text, page_id):
            frame = tk.Frame(parent, bg=SIDEBAR_BG, cursor="hand2")
            frame.pack(fill=tk.X, pady=1)
            left_bar = tk.Frame(frame, width=4, bg=SIDEBAR_BG)
            left_bar.pack(side=tk.LEFT, fill=tk.Y)
            left_bar.pack_propagate(False)
            text_lbl = tk.Label(
                frame,
                text=text,
                font=("Microsoft YaHei UI", 10),
                bg=SIDEBAR_BG,
                fg="#444444",
                anchor="w",
                padx=12,
                pady=10,
                cursor="hand2"
            )
            text_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            def on_click(e, pid=page_id):
                self._switch_page(pid)
            def on_enter(ev, f=frame, lb=left_bar, tl=text_lbl, pid=page_id):
                if pid != self.current_page:
                    f.config(bg="#dddddd")
                    lb.config(bg=SIDEBAR_BG)
                    tl.config(bg="#dddddd")
            def on_leave(ev, f=frame, lb=left_bar, tl=text_lbl, pid=page_id):
                if pid == self.current_page:
                    f.config(bg=SIDEBAR_SEL_BG)
                    lb.config(bg=SIDEBAR_LEFT_BAR)
                    tl.config(bg=SIDEBAR_SEL_BG)
                else:
                    f.config(bg=SIDEBAR_BG)
                    lb.config(bg=SIDEBAR_BG)
                    tl.config(bg=SIDEBAR_BG)
            
            for w in (frame, text_lbl, left_bar):
                w.bind("<Button-1>", on_click)
                w.bind("<Enter>", on_enter)
                w.bind("<Leave>", on_leave)
            return (frame, left_bar, text_lbl, page_id)
        
        self.sidebar_items.append(make_sidebar_item(self.sidebar, "Home", "主页"))
        for name in ["aCRF", "SDTM", "ADaM", "TFLs", "M5", "pm"]:
            self.sidebar_items.append(make_sidebar_item(self.sidebar, name, name))
        
        def update_sidebar_style():
            for frame, left_bar, text_lbl, page_id in self.sidebar_items:
                if page_id == self.current_page:
                    frame.config(bg=SIDEBAR_SEL_BG)
                    left_bar.config(bg=SIDEBAR_LEFT_BAR)
                    text_lbl.config(bg=SIDEBAR_SEL_BG, fg="#333333")
                else:
                    frame.config(bg=SIDEBAR_BG)
                    left_bar.config(bg=SIDEBAR_BG)
                    text_lbl.config(bg=SIDEBAR_BG, fg="#444444")
        
        self._update_sidebar_style = update_sidebar_style
        update_sidebar_style()
        
        # ========== 右侧主内容区（多页容器） ==========
        self.main_container = tk.Frame(content_row)
        self.main_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # ----- 顶部栏：6 个下拉框 + Launch / Refresh（所有页面均展示） -----
        top_frame = tk.Frame(self.main_container, bg="#f0f0f0", padx=10, pady=10)
        top_frame.pack(fill=tk.X)
        
        # 创建6个水平排列的下拉框（支持搜索）
        for i in range(6):
            combobox = SearchableCombobox(
                top_frame,
                width=15
            )
            combobox.pack(side=tk.LEFT, padx=5)
            
            # 绑定选择事件
            combobox.bind("<<ComboboxSelected>>", lambda e, idx=i: self.on_dropdown_selected(idx))
            
            self.comboboxes.append(combobox)
        
        # Launch SAS EG按钮
        self.open_btn = tk.Button(
            top_frame,
            text="Launch SAS EG",
            command=self.open_seguide,
            width=15,
            font=("Arial", 9)
        )
        self.open_btn.pack(side=tk.LEFT, padx=10)
        
        # Refresh Access按钮（放在Launch按钮旁边）
        self.refresh_btn = tk.Button(
            top_frame,
            text="Refresh Access",
            command=self.refresh_access,
            width=15,
            font=("Arial", 9)
        )
        self.refresh_btn.pack(side=tk.LEFT, padx=5)
        
        # ----- 欢迎区：在下拉框下面，所有页面均展示（问候语 + 6 个快捷按钮） -----
        self.welcome_frame = tk.Frame(self.main_container, bg="#f5f5f5", padx=10, pady=8)
        self.welcome_frame.pack(fill=tk.X)
        welcome_lbl = tk.Label(
            self.welcome_frame,
            text="您好, 从哪里开始您的工作呢?",
            font=("Microsoft YaHei UI", 11),
            bg="#f5f5f5",
            fg="#333333"
        )
        welcome_lbl.pack(anchor="w", pady=(0, 6))
        self.quick_jump_map = [
            ("aCRF", "04_crt"),
            ("SDTM", "01_sdtm"),
            ("ADaM", "02_adam"),
            ("TFLs", "03_reports"),
            ("M5", "05_pkpd"),
            ("pm", "06_programs"),
        ]
        self.quick_buttons = []
        btn_frame = tk.Frame(self.welcome_frame, bg="#f5f5f5")
        btn_frame.pack(anchor="w")
        for label_text, folder_key in self.quick_jump_map:
            btn = tk.Button(
                btn_frame,
                text=label_text,
                command=lambda page_id=label_text: self._switch_page(page_id),
                width=8,
                font=("Arial", 9),
                bg="#9e9e9e",
                fg="white",
                activebackground="#757575",
                activeforeground="white",
                relief=tk.FLAT,
                cursor="hand2",
                padx=8,
                pady=4
            )
            btn.pack(side=tk.LEFT, padx=4)
            self.quick_buttons.append(btn)
        
        # ----- 主页：仅 Subfolders 网格（仅 Home 页展示） -----
        self.home_page_frame = tk.Frame(self.main_container, bg="#f5f5f5")
        
        # 标签页控件（在主页内）
        self.notebook = ttk.Notebook(self.home_page_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # "Hyperlinks to Sub Folders"标签页
        self.hyperlinks_frame = tk.Frame(self.notebook, bg="#f5f5f5")
        self.notebook.add(self.hyperlinks_frame, text="Hyperlinks to Sub Folders")
        
        # 创建可滚动的网格容器（支持水平和垂直滚动）
        self.canvas = tk.Canvas(self.hyperlinks_frame, bg="#f5f5f5")
        v_scrollbar = ttk.Scrollbar(self.hyperlinks_frame, orient="vertical", command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(self.hyperlinks_frame, orient="horizontal", command=self.canvas.xview)
        self.grid_frame = tk.Frame(self.canvas, bg="#f5f5f5")
        
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # 布局滚动条和canvas
        self.canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        self.hyperlinks_frame.grid_rowconfigure(0, weight=1)
        self.hyperlinks_frame.grid_columnconfigure(0, weight=1)
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        
        def update_scroll_region(event=None):
            # 更新滚动区域（包括水平和垂直）
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        def on_canvas_configure(event):
            # 更新列宽为固定宽度（所有列保持一致）
            if hasattr(self, 'grid_frame') and self.grid_frame:
                column_width = 120  # 固定列宽120像素，所有列保持一致
                for col_idx in range(20):  # 假设最多20列
                    try:
                        self.grid_frame.columnconfigure(col_idx, minsize=column_width, weight=0)
                    except:
                        pass
            # 更新滚动区域
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        def on_frame_configure(event):
            # 当grid_frame大小改变时，更新canvas窗口宽度和滚动区域
            # 让canvas_window的宽度等于grid_frame的实际宽度，实现水平滚动
            frame_width = event.width
            canvas_width = self.canvas.winfo_width()
            # 设置canvas_window宽度为两者中的较大值，确保可以水平滚动
            self.canvas.itemconfig(self.canvas_window, width=max(frame_width, canvas_width))
            # 更新滚动区域
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        self.canvas.bind('<Configure>', on_canvas_configure)
        self.grid_frame.bind('<Configure>', on_frame_configure)
        
        # "Autoexec"标签页
        self.autoexec_frame = tk.Frame(self.notebook, bg="#f5f5f5")
        self.notebook.add(self.autoexec_frame, text="Autoexec")
        
        # Autoexec内容（暂时为空，可以后续添加）
        autoexec_label = tk.Label(
            self.autoexec_frame,
            text="Autoexec 配置区域",
            font=("Arial", 12),
            bg="#f5f5f5"
        )
        autoexec_label.pack(pady=50)
        
        # ----- 6 个独立页面（占位，可后续扩展内容） -----
        self.page_frames = {}
        for page_id in ["aCRF", "SDTM", "ADaM", "TFLs", "M5", "pm"]:
            f = tk.Frame(self.main_container, bg="#f5f5f5")
            if page_id == "TFLs":
                btn_row = tk.Frame(f, bg="#f5f5f5")
                btn_row.pack(anchor="w", padx=16, pady=16)
                btn_width = 10  # 两按钮列宽相同且较窄，文字各两行显示
                btn_metadata = tk.Button(
                    btn_row,
                    text="Metadata\nSetup",
                    command=lambda: show_metadata_setup_dialog(self),
                    width=btn_width,
                    font=("Microsoft YaHei UI", 10),
                    bg="#9e9e9e",
                    fg="white",
                    relief=tk.FLAT,
                    cursor="hand2",
                    padx=10,
                    pady=6
                )
                btn_metadata.pack(side=tk.LEFT, padx=(0, 8))
                btn_pdt = tk.Button(
                    btn_row,
                    text="PDT\nGen",
                    command=lambda: show_pdt_dialog(self),
                    width=btn_width,
                    font=("Microsoft YaHei UI", 10),
                    bg="#9e9e9e",
                    fg="white",
                    relief=tk.FLAT,
                    cursor="hand2",
                    padx=10,
                    pady=6
                )
                btn_pdt.pack(side=tk.LEFT, padx=(0, 8))
                btn_pgm_init = tk.Button(
                    btn_row,
                    text="Initial\nPGM",
                    command=lambda: run_initial_pgm(self),
                    width=btn_width,
                    font=("Microsoft YaHei UI", 10),
                    bg="#9e9e9e",
                    fg="white",
                    relief=tk.FLAT,
                    cursor="hand2",
                    padx=10,
                    pady=6
                )
                btn_pgm_init.pack(side=tk.LEFT, padx=(0, 8))
                btn_batch_run = tk.Button(
                    btn_row,
                    text="Batch\nRun",
                    command=lambda: None,  # 稍后补充
                    width=btn_width,
                    font=("Microsoft YaHei UI", 10),
                    bg="#9e9e9e",
                    fg="white",
                    relief=tk.FLAT,
                    cursor="hand2",
                    padx=10,
                    pady=6
                )
                btn_batch_run.pack(side=tk.LEFT, padx=(0, 8))
                btn_tfls_combine = tk.Button(
                    btn_row,
                    text="TFLs\nCombine",
                    command=lambda: None,  # 稍后补充
                    width=btn_width,
                    font=("Microsoft YaHei UI", 10),
                    bg="#9e9e9e",
                    fg="white",
                    relief=tk.FLAT,
                    cursor="hand2",
                    padx=10,
                    pady=6
                )
                btn_tfls_combine.pack(side=tk.LEFT)
            else:
                lbl = tk.Label(
                    f,
                    text=f"{page_id} 页面\n（内容可在此扩展）",
                    font=("Microsoft YaHei UI", 14),
                    bg="#f5f5f5",
                    fg="#666666"
                )
                lbl.pack(expand=True, pady=80)
            self.page_frames[page_id] = f
        
        # 状态栏（必须在 _switch_page 之前创建，否则切换页面时 update_status 会报错）
        self.status_label = tk.Label(
            self.root,
            text="就绪",
            relief=tk.SUNKEN,
            anchor=tk.W,
            padx=5,
            pady=2
        )
        self.status_label.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 默认显示主页
        self._switch_page("主页")
    
    def get_directories(self, path):
        """获取指定路径下的所有目录"""
        try:
            if not os.path.exists(path):
                return []
            
            items = os.listdir(path)
            directories = []
            
            for item in items:
                item_path = os.path.join(path, item)
                if os.path.isdir(item_path):
                    directories.append(item)
            
            # 按字母顺序排序
            directories.sort()
            return directories
        
        except PermissionError:
            self.show_error(f"没有权限访问: {path}")
            return []
        except Exception as e:
            self.show_error(f"访问路径时出错: {str(e)}")
            return []
    
    def refresh_first_dropdown(self):
        """刷新第一个下拉框（根目录）"""
        self.update_status("正在扫描Z盘根目录...")
        
        directories = self.get_directories(self.z_drive)
        
        if directories:
            # 使用set_values方法设置选项（支持搜索）
            if hasattr(self.comboboxes[0], 'set_values'):
                self.comboboxes[0].set_values(directories)
            else:
                self.comboboxes[0]['values'] = directories
            self.update_status(f"找到 {len(directories)} 个目录")
        else:
            if hasattr(self.comboboxes[0], 'set_values'):
                self.comboboxes[0].set_values([])
            else:
                self.comboboxes[0]['values'] = []
            self.update_status("Z盘根目录为空或无法访问")

    def _apply_test_defaults(self):
        """测试阶段：将前4个下拉框默认设置为 projects、HRS2129、HRS2129_test、csr_01"""
        defaults = ["projects", "HRS2129", "HRS2129_test", "csr_01"]
        for i in range(4, 6):
            self.comboboxes[i].set("")
            if hasattr(self.comboboxes[i], 'set_values'):
                self.comboboxes[i].set_values([])
            else:
                self.comboboxes[i]['values'] = []
            self.selected_paths[i] = ""
        current_path = self.z_drive
        for i, val in enumerate(defaults):
            if i >= 6:
                break
            dirs = self.get_directories(current_path)
            if not dirs or val not in dirs:
                break
            self.comboboxes[i].set(val)
            self.selected_paths[i] = val
            current_path = os.path.join(current_path, val)
            if i < 5:
                self.update_next_dropdown(i + 1, current_path)
        self.root.after(10, self.update_grid_display)

    def on_dropdown_selected(self, index):
        """当下拉框选择改变时的处理"""
        selected_value = self.comboboxes[index].get()
        
        if not selected_value:
            return
        
        # 更新当前路径
        self.selected_paths[index] = selected_value
        
        # 构建当前完整路径
        current_path = self.z_drive
        for i in range(index + 1):
            if self.selected_paths[i]:
                current_path = os.path.join(current_path, self.selected_paths[i])
        
        # 更新后续下拉框
        if index < 5:  # 如果不是最后一个下拉框
            self.update_next_dropdown(index + 1, current_path)
            
            # 清空更后面的下拉框
            for i in range(index + 2, 6):
                self.comboboxes[i].set("")
                if hasattr(self.comboboxes[i], 'set_values'):
                    self.comboboxes[i].set_values([])
                else:
                    self.comboboxes[i]['values'] = []
                self.selected_paths[i] = ""
        
        # 更新网格显示（只有选择完第4个下拉框后才显示）
        # 确保在更新网格之前，selected_paths已经正确更新
        self.root.after(10, self.update_grid_display)
    
    def update_next_dropdown(self, index, path):
        """更新指定索引的下拉框"""
        if index >= 6:
            return
        
        self.update_status(f"正在扫描: {path}")
        
        directories = self.get_directories(path)
        
        if directories:
            # 使用set_values方法设置选项（支持搜索）
            if hasattr(self.comboboxes[index], 'set_values'):
                self.comboboxes[index].set_values(directories)
            else:
                self.comboboxes[index]['values'] = directories
            self.update_status(f"找到 {len(directories)} 个子目录")
        else:
            if hasattr(self.comboboxes[index], 'set_values'):
                self.comboboxes[index].set_values([])
            else:
                self.comboboxes[index]['values'] = []
            self.update_status("该目录下没有子目录")
    
    def refresh_access(self):
        """刷新访问 - 清空所有下拉框并重新扫描"""
        # 清空所有下拉框
        for i in range(6):
            self.comboboxes[i].set("")
            if hasattr(self.comboboxes[i], 'set_values'):
                self.comboboxes[i].set_values([])
            else:
                self.comboboxes[i]['values'] = []
            self.selected_paths[i] = ""
            # 重置所有下拉框为可编辑状态（支持搜索）
            if hasattr(self.comboboxes[i], 'set_values'):
                # SearchableCombobox默认就是normal状态，无需修改
                pass
            else:
                self.comboboxes[i].config(state="readonly")
        
        # 清空网格显示
        self.clear_grid()
        
        # 重新扫描根目录
        self.refresh_first_dropdown()
        self.update_status("已刷新访问")
    
    def get_current_path(self):
        """获取当前选中的完整路径"""
        current_path = self.z_drive
        for path in self.selected_paths:
            if path:
                current_path = os.path.join(current_path, path)
        return current_path
    
    def clear_grid(self):
        """清空网格显示"""
        if self.grid_frame:
            for widget in self.grid_frame.winfo_children():
                widget.destroy()
        self.grid_labels = {}
    
    def _jump_to_column(self, folder_key):
        """快捷按钮：滚动到 Subfolders 中对应的列（页面）"""
        if not self.current_subfolders or not hasattr(self, 'canvas') or not self.canvas.winfo_exists():
            self.update_status("请先选择路径并等待 Subfolders 显示")
            return
        col_idx = None
        for i, name in enumerate(self.current_subfolders):
            if name == folder_key or folder_key.lower() in name.lower():
                col_idx = i
                break
        if col_idx is None:
            self.update_status(f"未找到对应列: {folder_key}")
            return
        column_width = 120
        padx = 2
        total_width = len(self.current_subfolders) * (column_width + padx * 2)
        if total_width <= 0:
            return
        x_start = col_idx * (column_width + padx * 2)
        try:
            self.canvas.xview_moveto(max(0.0, min(1.0, x_start / total_width)))
        except Exception:
            pass
        self.update_status(f"已跳转至 {self.current_subfolders[col_idx]}")
    
    def _switch_page(self, page_id):
        """切换左侧导航对应的页面（主页 / aCRF / SDTM / ADaM / TFLs / M5 / pm）"""
        self.current_page = page_id
        # 隐藏所有内容页
        self.home_page_frame.pack_forget()
        for f in self.page_frames.values():
            f.pack_forget()
        # 显示当前页
        if page_id == "主页":
            self.home_page_frame.pack(fill=tk.BOTH, expand=True)
        else:
            self.page_frames[page_id].pack(fill=tk.BOTH, expand=True)
        self._update_sidebar_style()
        self.update_status(f"当前页面: {page_id}")
    
    def update_grid_display(self):
        """更新网格显示，显示当前路径下的子文件夹（选择至少4级路径后显示）"""
        # 清空现有网格
        self.clear_grid()
        
        # 检查是否选择了至少4级路径（任意连续层级）
        selected_count = sum(1 for path in self.selected_paths if path)
        if selected_count < 4:
            return
        
        # 获取当前完整路径
        current_path = self.get_current_path()
        
        if not os.path.exists(current_path):
            self.update_status(f"路径不存在: {current_path}")
            return
        
        # 获取当前路径下的所有子文件夹
        try:
            items = os.listdir(current_path)
            subfolders = []
            
            for item in items:
                item_path = os.path.join(current_path, item)
                if os.path.isdir(item_path):
                    subfolders.append(item)
            
            # 任何时候都过滤掉指定的文件夹
            excluded_folders = ["00_source_data", "91_export", "92_import", "99_archive"]
            subfolders = [folder for folder in subfolders if folder not in excluded_folders]
            
            subfolders.sort()
            
            if not subfolders:
                return
            
            # 保存当前列顺序，供快捷按钮跳转使用
            self.current_subfolders = list(subfolders)
            
            # 创建列标题（每个子文件夹作为一列）
            for col_idx, folder_name in enumerate(subfolders):
                # 列标题（按钮样式，可点击）
                header_path = os.path.join(current_path, folder_name)
                header = tk.Label(
                    self.grid_frame,
                    text=folder_name,
                    bg="#d0d0d0",
                    relief=tk.RAISED,
                    font=("Arial", 9, "bold"),
                    padx=5,
                    pady=3,
                    cursor="hand2",
                    fg="blue"
                )
                header.grid(row=0, column=col_idx, sticky="ew", padx=2, pady=2)
                # 绑定点击事件 - 第一层文件夹使用Windows形式打开
                header.bind("<Button-1>", lambda e, path=header_path: self._open_folder_with_windows(path))
                header.bind("<Enter>", lambda e, lbl=header: lbl.config(bg="#b0b0b0"))
                header.bind("<Leave>", lambda e, lbl=header: lbl.config(bg="#d0d0d0"))
                
                # 获取该文件夹下的子项（包括文件夹和文件）
                folder_path = os.path.join(current_path, folder_name)
                try:
                    sub_items = os.listdir(folder_path)
                    # 分别获取文件夹和文件
                    sub_dirs = []
                    sub_files = []
                    for item in sub_items:
                        item_path = os.path.join(folder_path, item)
                        if os.path.isdir(item_path):
                            sub_dirs.append(item)
                        else:
                            sub_files.append(item)
                    sub_dirs.sort()
                    sub_files.sort()
                    sub_items = sub_dirs + sub_files  # 先显示文件夹，再显示文件
                except:
                    sub_items = []
                
                # 创建该列下的内容标签（可点击的超链接）
                col_labels = []
                for row_idx, sub_item in enumerate(sub_items, start=1):
                    item_path = os.path.join(folder_path, sub_item)
                    is_dir = os.path.isdir(item_path)
                    
                    # 检查文件扩展名
                    file_ext = os.path.splitext(sub_item)[1].lower() if not is_dir else ""
                    is_ps1 = file_ext == '.ps1'
                    
                    # 对于03_reports, 07_logs, 09_validation，文件名截断显示
                    display_text = sub_item
                    if folder_name in ["03_reports", "07_logs", "09_validation"]:
                        # 截断文件名，只显示前30个字符
                        max_length = 30
                        if len(sub_item) > max_length:
                            display_text = sub_item[:max_length] + "..."
                    
                    # 如果是.ps1文件，不显示为超链接样式
                    if is_ps1:
                        label = tk.Label(
                            self.grid_frame,
                            text=display_text,
                            bg="#ffffff",
                            anchor="w",
                            padx=5,
                            pady=2,
                            font=("Arial", 8),
                            cursor="arrow",  # 普通光标，不是手型
                            fg="black",  # 黑色文字，不是蓝色
                            # 不设置underline，即无下划线
                        )
                    else:
                        label = tk.Label(
                            self.grid_frame,
                            text=display_text,
                            bg="#ffffff",
                            anchor="w",
                            padx=5,
                            pady=2,
                            font=("Arial", 8),
                            cursor="hand2",
                            fg="blue",
                            underline=1
                        )
                    
                    label.grid(row=row_idx, column=col_idx, sticky="ew", padx=2, pady=1)
                    
                    # 如果不是.ps1文件，绑定点击事件
                    if not is_ps1:
                        label.bind("<Button-1>", lambda e, path=item_path: self.open_path(path))
                        
                        # 绑定悬停事件和tooltip
                        def make_hover_events(lbl, full_text, is_truncated):
                            def on_enter(event):
                                lbl.config(bg="#e0e0e0")
                                if is_truncated:
                                    # 显示tooltip
                                    tooltip = tk.Toplevel()
                                    tooltip.wm_overrideredirect(True)
                                    tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
                                    label_tooltip = tk.Label(tooltip, text=full_text, bg="#ffffe0", 
                                                            relief=tk.SOLID, borderwidth=1, padx=5, pady=2,
                                                            font=("Arial", 8))
                                    label_tooltip.pack()
                                    lbl.tooltip = tooltip
                            
                            def on_leave(event):
                                lbl.config(bg="#ffffff")
                                if hasattr(lbl, 'tooltip'):
                                    lbl.tooltip.destroy()
                                    del lbl.tooltip
                            
                            lbl.bind("<Enter>", on_enter)
                            lbl.bind("<Leave>", on_leave)
                        
                        make_hover_events(label, sub_item, display_text != sub_item)
                    
                    col_labels.append(label)
                
                # 存储列标题和内容标签
                self.grid_labels[folder_name] = {
                    'header': header,
                    'labels': col_labels
                }
            
            # 配置列宽
            # 03_reports, 07_logs, 09_validation的列宽和08_formats保持一致
            # 先找到08_formats的索引（如果存在）
            formats_idx = None
            for idx, folder in enumerate(subfolders):
                if folder == "08_formats":
                    formats_idx = idx
                    break
            
            # 如果找到08_formats，使用它的列宽；否则使用默认宽度
            base_column_width = 120  # 默认列宽120像素
            
            for col_idx, folder_name in enumerate(subfolders):
                # 对于03_reports, 07_logs, 09_validation，使用和08_formats相同的列宽
                if folder_name in ["03_reports", "07_logs", "09_validation"]:
                    if formats_idx is not None:
                        # 使用08_formats的列宽
                        column_width = base_column_width
                    else:
                        column_width = base_column_width
                else:
                    column_width = base_column_width
                
                self.grid_frame.columnconfigure(col_idx, minsize=column_width, weight=0)
            
            # 更新canvas滚动区域（包括水平和垂直）
            # 延迟更新，确保所有widget都已创建
            def update_scroll():
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
                # 更新canvas_window宽度以支持水平滚动
                frame_width = self.grid_frame.winfo_reqwidth()
                canvas_width = self.canvas.winfo_width()
                if frame_width > 0:
                    self.canvas.itemconfig(self.canvas_window, width=max(frame_width, canvas_width))
            
            self.root.after(50, update_scroll)
            
            self.update_status(f"显示 {len(subfolders)} 个文件夹")
        
        except Exception as e:
            self.update_status(f"更新网格时出错: {str(e)}")
    
    def _update_column_widths(self):
        """更新列宽为固定像素宽度（所有列保持一致）"""
        try:
            # 使用固定宽度，所有列保持一致
            column_width = 120  # 固定列宽120像素
            
            # 更新所有已存在的列
            if self.grid_frame:
                for col_idx in range(20):  # 假设最多20列
                    try:
                        self.grid_frame.columnconfigure(col_idx, minsize=column_width, weight=0)
                    except:
                        pass
        except Exception as e:
            print(f"更新列宽时出错: {e}")
    
    def open_path(self, path):
        """打开指定的文件夹或文件"""
        try:
            if os.path.exists(path):
                if os.path.isdir(path):
                    # 所有文件夹都使用Windows方式打开
                    folder_name = os.path.basename(path)
                    os.startfile(path)
                    self.update_status(f"已打开文件夹: {folder_name}")
                else:
                    # 检查文件扩展名和所在文件夹
                    file_ext = os.path.splitext(path)[1].lower()
                    file_dir = os.path.dirname(path)
                    folder_name = os.path.basename(file_dir)
                    
                    # 根据文件扩展名和所在文件夹选择打开方式
                    # .ps1文件不打开（已去除超链接）
                    if file_ext == '.ps1':
                        # .ps1文件不执行任何操作
                        self.update_status(f".ps1文件不可点击: {os.path.basename(path)}")
                    elif file_ext == '.xml' and folder_name in ['09_validation', '07_logs']:
                        # 09_validation和07_logs文件夹下的.xml文件使用Excel打开
                        self._open_with_excel(path)
                    elif file_ext == '.sas7bdat':
                        # 所有.sas7bdat文件都使用SAS EG打开
                        self._open_with_saseg(path)
                    elif file_ext == '.sas':
                        # 所有.sas文件都使用SAS EG打开
                        self._open_with_saseg(path)
                    else:
                        # 其他文件使用系统默认程序打开
                        os.startfile(path)
                        self.update_status(f"已打开文件: {os.path.basename(path)}")
            else:
                self.show_error(f"路径不存在: {path}")
        except Exception as e:
            error_msg = f"打开路径时出错: {str(e)}"
            self.show_error(error_msg)
            self.update_status(error_msg)
    
    def _open_with_powershell(self, file_path):
        """使用PowerShell打开.ps1文件"""
        try:
            # 确保路径是绝对路径
            abs_path = os.path.abspath(file_path)
            
            # 使用PowerShell执行.ps1文件
            # 方法1：直接使用subprocess.Popen，传递参数列表
            powershell_exe = "powershell.exe"
            args = [
                powershell_exe,
                "-ExecutionPolicy", "Bypass",
                "-File", abs_path
            ]
            
            # 尝试使用参数列表方式
            try:
                subprocess.Popen(args, shell=False)
                self.update_status(f"正在使用PowerShell打开: {os.path.basename(file_path)}")
            except Exception as e1:
                # 如果失败，尝试使用字符串命令方式
                try:
                    powershell_cmd = f'powershell.exe -ExecutionPolicy Bypass -File "{abs_path}"'
                    subprocess.Popen(powershell_cmd, shell=True)
                    self.update_status(f"正在使用PowerShell打开: {os.path.basename(file_path)}")
                except Exception as e2:
                    # 如果还是失败，尝试使用start命令
                    os.startfile(abs_path)
                    self.update_status(f"正在打开: {os.path.basename(file_path)}")
        except Exception as e:
            error_msg = f"使用PowerShell打开文件时出错: {str(e)}"
            self.show_error(error_msg)
            self.update_status(error_msg)
    
    def _open_with_excel(self, file_path):
        """使用Excel打开.xml文件"""
        try:
            # Excel的路径（通常可以通过startfile直接打开，但也可以指定Excel程序）
            # 方法1：使用os.startfile，系统会自动用Excel打开
            # 方法2：直接指定Excel程序路径
            excel_path = r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'
            
            # 先尝试使用指定的Excel路径
            if os.path.exists(excel_path):
                subprocess.Popen([excel_path, file_path], shell=False)
            else:
                # 如果找不到Excel，使用系统默认方式（通常也会用Excel打开）
                os.startfile(file_path)
            
            self.update_status(f"正在使用Excel打开: {os.path.basename(file_path)}")
        except Exception as e:
            # 如果出错，尝试使用系统默认方式
            try:
                os.startfile(file_path)
                self.update_status(f"正在打开: {os.path.basename(file_path)}")
            except Exception as e2:
                error_msg = f"使用Excel打开文件时出错: {str(e2)}"
                self.show_error(error_msg)
                self.update_status(error_msg)
    
    def _open_folder_with_windows(self, folder_path):
        """使用Windows资源管理器打开文件夹"""
        try:
            # 使用Windows资源管理器打开文件夹
            os.startfile(folder_path)
            self.update_status(f"已打开文件夹: {os.path.basename(folder_path)}")
        except Exception as e:
            error_msg = f"打开文件夹时出错: {str(e)}"
            self.show_error(error_msg)
            self.update_status(error_msg)
    
    def _convert_path_for_saseg(self, path):
        """将Windows路径转换为SAS EG使用的映射路径"""
        # 将路径标准化（处理反斜杠和正斜杠）
        normalized_path = os.path.normpath(path)
        
        # 统一转换为大写进行比较
        upper_path = normalized_path.upper()
        
        # 检查路径是否以Z:\或Z:/开头（不区分大小写）
        if upper_path.startswith('Z:\\') or upper_path.startswith('Z:/'):
            # 移除Z:\或Z:/前缀（3个字符）
            remaining_path = normalized_path[3:]
            
            # 将反斜杠转换为正斜杠
            remaining_path = remaining_path.replace('\\', '/')
            
            # 确保路径以正斜杠开头（移除可能的前导斜杠）
            if remaining_path.startswith('/'):
                remaining_path = remaining_path[1:]
            
            # 构建映射路径
            mapped_path = f"/u01/app/sas/sas9.4/DocumentRepository/DDT/{remaining_path}"
            return mapped_path
        else:
            # 如果不是Z盘路径，直接转换反斜杠为正斜杠
            return path.replace('\\', '/')
    
    def _open_folder_with_saseg(self, folder_path):
        """使用SAS EG打开文件夹"""
        try:
            # SAS EG的路径
            seguide_path = r'C:\Program Files\SaS\SASHome\SASEnterpriseGuide\8\SEGuide.exe'
            
            # 转换路径：将Z:\转换为映射路径
            mapped_path = self._convert_path_for_saseg(folder_path)
            
            # 启动SAS EG（使用转换后的路径作为参数）
            subprocess.Popen([seguide_path, mapped_path], shell=False)
            
            self.update_status(f"正在使用SAS EG打开文件夹: {os.path.basename(folder_path)}")
        except Exception as e:
            error_msg = f"使用SAS EG打开文件夹时出错: {str(e)}"
            self.show_error(error_msg)
            self.update_status(error_msg)
    
    def _open_with_saseg(self, file_path):
        """使用SAS EG打开文件（.sas7bdat, .sas等）"""
        try:
            # SAS EG的路径
            seguide_path = r'C:\Program Files\SaS\SASHome\SASEnterpriseGuide\8\SEGuide.exe'
            
            # 转换路径：将Z:\转换为映射路径
            mapped_path = self._convert_path_for_saseg(file_path)
            
            # 使用subprocess启动SAS EG并打开文件（使用转换后的路径）
            subprocess.Popen([seguide_path, mapped_path], shell=False)
            
            self.update_status(f"正在使用SAS EG打开: {os.path.basename(file_path)}")
        except Exception as e:
            error_msg = f"使用SAS EG打开文件时出错: {str(e)}"
            self.show_error(error_msg)
            self.update_status(error_msg)
    
    def open_seguide(self):
        """打开SEGuide.exe"""
        # 禁用按钮，防止重复点击
        self.open_btn.config(state=tk.DISABLED)
        self.update_status("正在启动SEGuide...")
        
        # 在后台线程中执行，避免界面冻结
        thread = threading.Thread(target=self._launch_seguide)
        thread.daemon = True
        thread.start()
    
    def _launch_seguide(self):
        """在后台线程中启动SEGuide并导航到选择的路径"""
        try:
            # 获取用户选择的完整路径
            target_path = self.get_current_path()
            
            # 启动应用程序
            app = Application(backend='uia').start(
                r'"C:\Program Files\SaS\SASHome\SASEnterpriseGuide\8\SEGuide.exe"'
            
            )
            
            # 等待对话框出现image.png
            app.Dialog.wait("exists ready", timeout=15, retry_interval=3)
            
            # 查找子窗口
            dlg = app.Dialog.child_window(
                title="服务器", 
                auto_id="服务器", 
                control_type="Pane"
            )
            dlg1 = dlg.child_window(title="服务器", control_type="TreeItem")
            
            # 先展开服务器节点
            dlg1.ensure_visible()
            dlg1.expand()
            time.sleep(2)
            
            # 再找到并展开SASApp节点
            SASApp = dlg1.child_window(title='SASApp', control_type="TreeItem")
            SASApp.ensure_visible()
            SASApp.expand()
            time.sleep(3)
            
            # 尝试多种方式查找文件节点
            file = None
            
            # 方法1：直接查找'文件'
            try:
                file = SASApp.child_window(title='文件', control_type="TreeItem")
                file.wait("exists", timeout=5)
            except:
                pass
            
            # 方法2：模糊匹配
            if file is None:
                try:
                    file = SASApp.child_window(
                        title_re=".*文件.*", 
                        control_type="TreeItem"
                    )
                    file.wait("exists", timeout=3)
                except:
                    pass
            
            # 方法3：查找英文"Files"
            if file is None:
                try:
                    file = SASApp.child_window(title='Files', control_type="TreeItem")
                    file.wait("exists", timeout=3)
                except:
                    pass
            
            # 方法4：遍历子元素查找
            if file is None:
                for child in SASApp.children():
                    try:
                        child_title = child.window_text().lower()
                        if "文件" in child_title or "file" in child_title:
                            file = child
                            break
                    except:
                        continue
            
            # 方法5：按索引查找
            if file is None:
                children = SASApp.children()
                if len(children) > 0:
                    try:
                        file = children[0]
                    except:
                        pass
            
            if file is not None:
                # 确保可见并双击
                file.ensure_visible()
                time.sleep(1)
                file.click_input(button='left', double=True)
                
                # 发送回车键
                send_keys("{VK_RETURN}")
                time.sleep(2)
                
                # 导航到目标路径
                self._navigate_to_path(app, target_path)
            else:
                self.root.after(0, lambda: self.update_status("SEGuide已启动，但无法找到文件节点"))
        
        except Exception as e:
            error_msg = f"启动SEGuide时出错: {str(e)}"
            self.root.after(0, lambda: self.update_status(error_msg))
            self.root.after(0, lambda: self.show_error(error_msg))
        
        finally:
            # 重新启用按钮
            self.root.after(0, lambda: self.open_btn.config(state=tk.NORMAL))
    
    def _navigate_to_path(self, app, target_path):
        """在SEGuide中导航到指定路径，从文件节点开始逐级展开"""
        try:
            # 获取路径的各个部分（6个下拉框的值）
            path_parts = []
            for path_part in self.selected_paths:
                if path_part:
                    path_parts.append(path_part)
            
            if not path_parts:
                return
            
            # 等待文件浏览器窗口出现
            time.sleep(3)
            
            try:
                # 查找对话框窗口（服务器连接对话框）
                dlg = app.Dialog
                
                # 查找"文件"节点（Files节点）
                file_node = None
                
                # 方法1：通过已知的路径查找文件节点
                try:
                    # 从之前的代码知道，文件节点在：服务器 > SASApp > 文件
                    dlg_pane = dlg.child_window(title="服务器", auto_id="服务器", control_type="Pane")
                    server_item = dlg_pane.child_window(title="服务器", control_type="TreeItem")
                    
                    # 确保SASApp已展开
                    try:
                        sasapp = server_item.child_window(title='SASApp', control_type="TreeItem")
                        if not sasapp.is_expanded():
                            sasapp.expand()
                            time.sleep(1)
                    except:
                        pass
                    
                    # 查找"文件"节点
                    try:
                        file_node = server_item.child_window(title='文件', control_type="TreeItem")
                        if not file_node.exists():
                            # 尝试查找英文"Files"
                            file_node = server_item.child_window(title='Files', control_type="TreeItem")
                    except:
                        pass
                    
                    # 如果还是找不到，尝试遍历子节点
                    if file_node is None or not file_node.exists():
                        try:
                            sasapp = server_item.child_window(title='SASApp', control_type="TreeItem")
                            children = sasapp.children()
                            for child in children:
                                try:
                                    child_text = child.window_text().lower()
                                    if "文件" in child_text or "file" in child_text:
                                        file_node = child
                                        break
                                except:
                                    continue
                        except:
                            pass
                
                except Exception as e:
                    print(f"查找文件节点时出错: {e}")
                
                if file_node and file_node.exists():
                    # 确保文件节点可见并已展开
                    file_node.ensure_visible()
                    if not file_node.is_expanded():
                        file_node.expand()
                    time.sleep(1)
                    
                    # 从文件节点开始，逐级展开路径
                    current_node = file_node
                    
                    for i, part in enumerate(path_parts):
                        found = False
                        try:
                            # 确保当前节点已展开
                            if not current_node.is_expanded():
                                current_node.expand()
                                time.sleep(0.8)  # 等待展开完成
                            
                            # 获取当前节点的所有子节点
                            children = current_node.children()
                            
                            # 查找匹配的子节点
                            for child in children:
                                try:
                                    child_text = child.window_text()
                                    # 精确匹配或忽略大小写匹配
                                    if child_text == part or child_text.lower() == part.lower():
                                        current_node = child
                                        current_node.ensure_visible()
                                        time.sleep(0.5)
                                        
                                        # 展开这个节点（包括最后一个，以便打开文件夹）
                                        if not current_node.is_expanded():
                                            current_node.expand()
                                            time.sleep(0.8)
                                        
                                        found = True
                                        break
                                except Exception as e:
                                    print(f"检查子节点时出错: {e}")
                                    continue
                            
                            if not found:
                                print(f"无法找到路径部分: {part}")
                                self.root.after(0, lambda p=part: self.update_status(f"无法找到文件夹: {p}"))
                                break
                        
                        except Exception as e:
                            print(f"展开路径部分 {part} 时出错: {e}")
                            break
                    
                    # 选中并展开最终节点（打开最后一个文件夹）
                    if current_node:
                        current_node.ensure_visible()
                        # 确保最后一个文件夹也被展开（打开）
                        if not current_node.is_expanded():
                            current_node.expand()
                            time.sleep(0.8)
                        current_node.select()
                        time.sleep(0.5)
                        self.root.after(0, lambda: self.update_status(f"已导航到并打开: {' > '.join(path_parts)}"))
                        return
                    else:
                        self.root.after(0, lambda: self.update_status("导航完成，但未找到最终节点"))
                
                else:
                    print("无法找到文件节点")
                    self.root.after(0, lambda: self.update_status("无法找到文件节点"))
            
            except Exception as e:
                print(f"导航过程出错: {e}")
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: self.update_status(f"导航出错: {str(e)}"))
        
        except Exception as e:
            print(f"导航过程出错: {e}")
            import traceback
            traceback.print_exc()
            self.root.after(0, lambda: self.update_status(f"导航出错: {str(e)}"))
    
    def update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def show_error(self, message):
        """显示错误消息框"""
        messagebox.showerror("错误", message)


def main():
    root = tk.Tk()
    app = SASEGGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
