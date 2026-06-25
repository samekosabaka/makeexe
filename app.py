import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from datetime import datetime, timedelta
import os
import sys
import ctypes 
import warnings
import threading

# --- TẮT CÁC CẢNH BÁO (WARNINGS) KHÔNG CẦN THIẾT TRÊN TERMINAL ---
warnings.filterwarnings('ignore', category=UserWarning)

# --- GIỮ KẾT NỐI MUTEX ĐỂ KHÔNG BỊ RÁC BỘ NHỚ ---
global_mutex = None

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- LỚP TÙY CHỈNH: DROP DOWN CÓ CHECKBOX VÀ THANH TÌM KIẾM ---
class MultiSelect(tk.Frame):
    def __init__(self, parent, title="Tất cả", width=15, command=None):
        super().__init__(parent)
        self.command = command
        self.values = []
        self.selected = []
        self.display_var = tk.StringVar(value=title)
        self.title_default = title
        
        self.btn = ttk.Button(self, textvariable=self.display_var, width=width, command=self.show_popup)
        self.btn.pack(fill=tk.BOTH, expand=True)
        
    def set_values(self, values):
        if set(self.values) != set(values):
            self.values = values
            self.selected = [s for s in self.selected if s in values]
            self._update_display()
            
    def _update_display(self):
        if not self.selected or len(self.selected) == len(self.values):
            self.display_var.set(self.title_default)
        elif len(self.selected) == 1:
            self.display_var.set(self.selected[0])
        else:
            self.display_var.set(f"Đã chọn {len(self.selected)} mục")

    def get(self):
        return self.selected if self.selected else ["Tất cả"]
        
    def show_popup(self):
        popup = tk.Toplevel(self)
        popup.title(self.title_default)
        popup.geometry("250x320")
        popup.transient(self.winfo_toplevel())
        popup.grab_set()
        
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        popup.geometry(f"+{x}+{y}")
        
        search_var = tk.StringVar()
        search_entry = ttk.Entry(popup, textvariable=search_var)
        search_entry.pack(fill=tk.X, padx=5, pady=5)
        
        f_canvas = tk.Frame(popup)
        f_canvas.pack(fill=tk.BOTH, expand=True, padx=5)
        
        canvas = tk.Canvas(f_canvas)
        scroll = ttk.Scrollbar(f_canvas, orient="vertical", command=canvas.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas.configure(yscrollcommand=scroll.set)
        
        inner_frame = tk.Frame(canvas)
        canvas_window = canvas.create_window((0,0), window=inner_frame, anchor='nw')
        
        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
        inner_frame.bind('<Configure>', on_configure)
        
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind('<Configure>', on_canvas_configure)
        
        cb_vars = {}
        for val in self.values:
            cb_vars[val] = tk.BooleanVar(value=(val in self.selected))
            
        def populate_list(*args):
            for widget in inner_frame.winfo_children():
                widget.destroy()
            search_term = search_var.get().lower()
            def toggle_all():
                new_state = var_all.get()
                for val, var in cb_vars.items():
                    if search_term in str(val).lower():
                        var.set(new_state)
            var_all = tk.BooleanVar(value=False)
            chk_all = tk.Checkbutton(inner_frame, text="-- Chọn / Bỏ chọn tất cả --", variable=var_all, command=toggle_all, fg="blue", font=("Arial", 8, "bold"))
            chk_all.pack(anchor='w', pady=(0, 5))
            
            for val in self.values:
                if search_term in str(val).lower():
                    chk = tk.Checkbutton(inner_frame, text=str(val), variable=cb_vars[val])
                    chk.pack(anchor='w')
                    
        search_var.trace("w", populate_list)
        populate_list()
        
        def on_confirm():
            self.selected = [val for val, var in cb_vars.items() if var.get()]
            self._update_display()
            popup.destroy()
            if self.command:
                self.command()
        
        btn_ok = ttk.Button(popup, text="Xác nhận", command=on_confirm)
        btn_ok.pack(pady=5, fill=tk.X, padx=10)
        search_entry.focus_set()


class TimekeepApp:
    def __init__(self, root):
        self.root = root
        self.root.title("iamhisu")
        self.root.geometry("1250x700") 
        
        try:
            icon_path = resource_path('noface.ico')
            self.root.iconbitmap(icon_path)
        except Exception:
            pass
        
        self.root.update_idletasks()
        width = 1250
        height = 700
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        self.file1_path = ""
        self.file2_path = ""
        self.file_xp_path = "" 
        self.df1_clean = None
        self.df2_clean = None
        self.df_result = None
        self.df_xp_result = None
        self.current_theme = "Sáng"
        self.is_pinned = False

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="white", foreground="black", rowheight=25, fieldbackground="white")
        style.map('Treeview', background=[('selected', '#347083')])
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'), background="#d3d3d3")
        
        style.configure("TNotebook", background="#f0f0f0")
        style.configure("TNotebook.Tab", padding=[15, 5], font=('Arial', 10, 'bold'))

        # =========================================================================
        # --- UI MỚI: ĐẶT NOTEBOOK LÊN TRƯỚC ĐỂ LÀM NỀN ---
        # =========================================================================
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tab_chamcong = tk.Frame(self.notebook)
        self.tab_xinphep = tk.Frame(self.notebook)

        self.notebook.add(self.tab_chamcong, text="🕒 Chấm Công")
        self.notebook.add(self.tab_xinphep, text="📝 Xin Phép")

        # =========================================================================
        # --- UI MỚI: KHUNG TRÔI TỰ DO (FLOATING CONTROLS) CÙNG HÀNG VỚI TAB ---
        # =========================================================================
        self.control_frame = tk.Frame(self.root)
        self.control_frame.place(relx=1.0, y=6, anchor="ne", x=-10)

        # Nút Ghim (Pin on top)
        self.btn_pin = tk.Button(self.control_frame, text="📌 Ghim: TẮT", width=12, command=self.toggle_pin, bg="#e0e0e0", font=("Arial", 8, "bold"), cursor="hand2")
        self.btn_pin.pack(side=tk.RIGHT, padx=(5, 0))

        # Nút Đổi Theme
        self.btn_theme = tk.Button(self.control_frame, text="🎨 Giao diện: Sáng", width=16, command=self.toggle_theme, bg="#e0e0e0", font=("Arial", 8, "bold"), cursor="hand2")
        self.btn_theme.pack(side=tk.RIGHT, padx=5)

        # Thanh Opacity
        frame_opacity = tk.Frame(self.control_frame)
        frame_opacity.pack(side=tk.RIGHT, padx=5)
        self.lbl_opacity = tk.Label(frame_opacity, text="Độ mờ:", font=("Arial", 8, "bold"))
        self.lbl_opacity.pack(side=tk.LEFT)
        self.opacity_slider = ttk.Scale(frame_opacity, from_=0.2, to=1.0, value=1.0, command=self.change_opacity)
        self.opacity_slider.pack(side=tk.LEFT, padx=2)


        # ==================== TAB 1: CHẤM CÔNG ====================
        frame_top = tk.Frame(self.tab_chamcong, pady=5)
        frame_top.pack(fill=tk.X, pady=(15, 0))

        tk.Label(frame_top, text="1. File zkt:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.lbl_file1 = tk.Label(frame_top, text="Chưa chọn file", width=35, anchor="w", bg="white", relief="sunken")
        self.lbl_file1.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(frame_top, text="Chọn File", command=self.select_file1).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(frame_top, text="2. Dữ liệu HR:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.lbl_file2 = tk.Label(frame_top, text="Chưa chọn file", width=35, anchor="w", bg="white", relief="sunken")
        self.lbl_file2.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(frame_top, text="Chọn File", command=self.select_file2).grid(row=1, column=2, padx=5, pady=5)

        tk.Button(frame_top, text="Bắt đầu Xử lý", command=self.process_data, bg="lightblue", font=("Arial", 10, "bold")).grid(row=0, column=3, rowspan=2, padx=20, ipadx=10, ipady=15)
        
        frame_hints = tk.Frame(frame_top)
        frame_hints.grid(row=0, column=4, rowspan=2, padx=10, sticky="w")
        tk.Label(frame_hints, text="(MẸO: Nhấp đúp chuột vào cột Timetable để sửa tên Ca và tính lại dữ liệu)", fg="red", font=("Arial", 9, "italic")).pack(anchor="w")
        tk.Label(frame_hints, text="(HỆ THỐNG: Đã tích hợp bộ lõi AI tự động đọc file Excel ZKTime giả mạo.)", fg="blue", font=("Arial", 9, "italic")).pack(anchor="w")

        frame_filter = tk.Frame(self.tab_chamcong, pady=5)
        frame_filter.pack(fill=tk.X, padx=10)
        
        tk.Label(frame_filter, text="Ngày:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(5, 2))
        self.cb_date = MultiSelect(frame_filter, title="Tất cả Ngày", width=15, command=self.apply_filter)
        self.cb_date.pack(side=tk.LEFT, padx=2)
        
        tk.Label(frame_filter, text="MSNV:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(10, 2))
        self.cb_msnv = MultiSelect(frame_filter, title="Tất cả MSNV", width=15, command=self.apply_filter)
        self.cb_msnv.pack(side=tk.LEFT, padx=2)
        
        tk.Label(frame_filter, text="Ca làm:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(10, 2))
        self.cb_timetable = ttk.Combobox(frame_filter, state="readonly", width=15, font=("Arial", 9))
        self.cb_timetable.pack(side=tk.LEFT, padx=2)
        
        tk.Label(frame_filter, text="Work Time:", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(10, 2))
        self.cb_worktime = ttk.Combobox(frame_filter, state="readonly", values=["Tất cả", "Bù vân tay"], width=12, font=("Arial", 9))
        self.cb_worktime.pack(side=tk.LEFT, padx=2)
        self.cb_worktime.current(0)
        
        btn_filter = tk.Button(frame_filter, text="↻ Xóa Lọc / Tải Lại", command=self.reset_filter, bg="#ffcc00", font=("Arial", 9, "bold"))
        btn_filter.pack(side=tk.LEFT, padx=(15, 5))

        self.cb_timetable.bind("<<ComboboxSelected>>", self.apply_filter)
        self.cb_worktime.bind("<<ComboboxSelected>>", self.apply_filter)

        frame_middle = tk.Frame(self.tab_chamcong)
        frame_middle.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scroll_y = tk.Scrollbar(frame_middle, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(frame_middle, orient=tk.HORIZONTAL)
        
        self.tree = ttk.Treeview(frame_middle, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.tree.yview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.config(command=self.tree.xview)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.tag_configure('oddrow', background="white", foreground="black")
        self.tree.tag_configure('evenrow', background="#f0f0f0", foreground="black")

        self.tree.bind('<Double-1>', self.on_double_click)

        frame_bottom = tk.Frame(self.tab_chamcong, pady=10)
        frame_bottom.pack(fill=tk.X)
        tk.Button(frame_bottom, text="Xuất Bảng Dữ Liệu", command=self.export_excel, bg="lightgreen", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=20, pady=5)
        tk.Button(frame_bottom, text="Xuất Theo Form Cố Định", command=self.export_fixed_form, bg="#FF9999", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=20, pady=5)
        
        # ==================== TAB 2: XIN PHÉP ====================
        frame_top_xp = tk.Frame(self.tab_xinphep, pady=20)
        frame_top_xp.pack(fill=tk.X, pady=(15, 0))

        tk.Label(frame_top_xp, text="1. Dữ liệu đầu vào:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.lbl_file_xp = tk.Label(frame_top_xp, text="Chưa chọn file", width=40, anchor="w", bg="white", relief="sunken")
        self.lbl_file_xp.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(frame_top_xp, text="Chọn File", command=self.select_file_xp).grid(row=0, column=2, padx=5, pady=5)

        tk.Button(frame_top_xp, text="Xử lý dữ liệu", command=self.process_xp, bg="lightblue", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=10, ipadx=10, ipady=5)
        tk.Button(frame_top_xp, text="Xuất kết quả", command=self.export_xp, bg="lightgreen", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=10, ipadx=10, ipady=5)

        frame_middle_xp = tk.Frame(self.tab_xinphep)
        frame_middle_xp.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scroll_y_xp = tk.Scrollbar(frame_middle_xp, orient=tk.VERTICAL)
        scroll_x_xp = tk.Scrollbar(frame_middle_xp, orient=tk.HORIZONTAL)
        
        self.tree_xp = ttk.Treeview(frame_middle_xp, yscrollcommand=scroll_y_xp.set, xscrollcommand=scroll_x_xp.set)
        scroll_y_xp.config(command=self.tree_xp.yview)
        scroll_y_xp.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x_xp.config(command=self.tree_xp.xview)
        scroll_x_xp.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_xp.pack(fill=tk.BOTH, expand=True)

        self.tree_xp.tag_configure('oddrow', background="white", foreground="black")
        self.tree_xp.tag_configure('evenrow', background="#f0f0f0", foreground="black")

        self.apply_theme("Sáng")

    # =========================================================================
    # --- LOGIC GIAO DIỆN: OPACITY, THEME & PIN ---
    # =========================================================================
    def change_opacity(self, val):
        self.root.attributes('-alpha', float(val))

    def toggle_pin(self):
        self.is_pinned = not self.is_pinned
        self.root.attributes('-topmost', self.is_pinned)
        
        if self.is_pinned:
            self.btn_pin.config(text="📌 Ghim: BẬT", bg="#ff9999", fg="black")
        else:
            self.btn_pin.config(text="📌 Ghim: TẮT")
            self.apply_theme(self.current_theme)

    def toggle_theme(self):
        themes = ["Sáng", "Tối", "Trong suốt"]
        idx = themes.index(self.current_theme)
        self.current_theme = themes[(idx + 1) % len(themes)]
        self.apply_theme(self.current_theme)

    def apply_theme(self, theme):
        self.btn_theme.config(text=f"🎨 Giao diện: {theme}")
        style = ttk.Style()
        
        if theme == "Sáng":
            bg_main = "#f0f0f0"
            fg_main = "black"
            bg_btn = "#e0e0e0"
            bg_lbl = "white"
            self.root.attributes('-alpha', self.opacity_slider.get())
            style.configure("TNotebook", background=bg_main)
            style.configure("TNotebook.Tab", background="#d3d3d3", foreground="black")
            style.map("TNotebook.Tab", background=[("selected", "white")])
        elif theme == "Tối":
            bg_main = "#2b2b2b"
            fg_main = "white"
            bg_btn = "#555555"
            bg_lbl = "#444444"
            self.root.attributes('-alpha', self.opacity_slider.get())
            style.configure("TNotebook", background=bg_main)
            style.configure("TNotebook.Tab", background="#3c3f41", foreground="white")
            style.map("TNotebook.Tab", background=[("selected", bg_main)])
        elif theme == "Trong suốt":
            bg_main = "#0a0a0a"
            fg_main = "#00ffcc"
            bg_btn = "#222222"
            bg_lbl = "#1a1a1a"
            self.opacity_slider.set(0.85)
            self.root.attributes('-alpha', 0.85)
            style.configure("TNotebook", background=bg_main)
            style.configure("TNotebook.Tab", background="#1a1a1a", foreground="#00ffcc")
            style.map("TNotebook.Tab", background=[("selected", bg_main)])

        def _update_colors(widget):
            try:
                wtype = widget.winfo_class()
                if wtype in ('Frame', 'Tk', 'Toplevel'):
                    widget.configure(bg=bg_main)
                elif wtype == 'Label':
                    if widget in [self.lbl_file1, self.lbl_file2, self.lbl_file_xp]:
                        widget.configure(bg=bg_lbl, fg=fg_main)
                    else:
                        c_fg = widget.cget('fg')
                        if c_fg in ['red', 'blue', 'gray']:
                            widget.configure(bg=bg_main) 
                        else:
                            widget.configure(bg=bg_main, fg=fg_main)
                elif wtype == 'Button' or 'Button' in wtype:
                    c_text = widget.cget('text') if hasattr(widget, 'cget') else ''
                    if c_text in ["Bắt đầu Xử lý", "Xuất Bảng Dữ Liệu", "Xuất Theo Form Cố Định", "↻ Xóa Lọc / Tải Lại", "Xuất kết quả", "Xử lý dữ liệu"]:
                        pass 
                    elif "📌 Ghim: BẬT" in c_text:
                        widget.configure(bg="#ff9999", fg="black")
                    else:
                        widget.configure(bg=bg_btn, fg=fg_main)
            except Exception:
                pass
            for child in widget.winfo_children():
                _update_colors(child)
                
        _update_colors(self.root)
        
        style.configure("Treeview", background=bg_lbl, fieldbackground=bg_lbl, foreground=fg_main)
        style.configure("Treeview.Heading", background=bg_btn, foreground=fg_main)
        self.tree.tag_configure('oddrow', background=bg_lbl, foreground=fg_main)
        self.tree.tag_configure('evenrow', background=bg_main, foreground=fg_main)
        self.tree_xp.tag_configure('oddrow', background=bg_lbl, foreground=fg_main)
        self.tree_xp.tag_configure('evenrow', background=bg_main, foreground=fg_main)

    # =========================================================================
    # --- LOGIC CHO TAB 2: XIN PHÉP ---
    # =========================================================================
    def select_file_xp(self):
        self.file_xp_path = filedialog.askopenfilename(filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv")])
        if self.file_xp_path:
            self.lbl_file_xp.config(text=self.file_xp_path.split('/')[-1])
            
    def process_xp(self):
        import pandas as pd
        if not self.file_xp_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file dữ liệu đầu vào trước!")
            return

        try:
            # SỬ DỤNG HÀM READ_FILE NỘI SOI MỚI NÂNG CẤP
            df = self.read_file(self.file_xp_path)

            req_cols = ['Bất thường chấm công', 'Mã số nhân viên', 'Ngày', 'Thời gian quẹt thẻ lên ca', 'Thời gian quẹt thẻ xuống ca', 'Số ca làm việc trong ngày']
            missing = [col for col in req_cols if col not in df.columns]
            if missing:
                messagebox.showerror("Lỗi dữ liệu", f"File đầu vào thiếu các cột sau:\n{', '.join(missing)}")
                return

            shift_dict = {}
            for idx, row in df.iterrows():
                m_val = str(row.get('Mã số nhân viên', '')).strip()
                d_val = row.get('Ngày')
                s_val = str(row.get('Số ca làm việc trong ngày', '')).upper()
                
                if m_val and pd.notna(d_val):
                    try:
                        d_str = str(d_val).split(' ')[0]
                        if '-' in d_str: d_p = d_str.split('-')
                        elif '/' in d_str: d_p = d_str.split('/')
                        else: d_p = []
                        
                        if len(d_p) == 3:
                            if len(d_p[0]) == 4:
                                d_obj = datetime(int(d_p[0]), int(d_p[1]), int(d_p[2])).date()
                            else:
                                d_obj = datetime(int(d_p[2]), int(d_p[1]), int(d_p[0])).date()
                            shift_dict[(m_val, d_obj)] = s_val
                    except:
                        pass

            df = df[df['Bất thường chấm công'].astype(str).str.strip().str.lower() == 'phải'].copy()

            def get_shift_times(base_date, shift_name):
                dt = datetime.combine(base_date, datetime.min.time())
                sn = str(shift_name).upper()
                if "7H30-17H" in sn or "7:30-17:00" in sn:
                    return dt.replace(hour=7, minute=30), dt.replace(hour=17, minute=0)
                elif "7H30-15H30" in sn or "7:30-15:30" in sn:
                    return dt.replace(hour=7, minute=30), dt.replace(hour=15, minute=30)
                elif "SÁNG" in sn or "CA 1" in sn:
                    return dt.replace(hour=6, minute=0), dt.replace(hour=14, minute=0)
                elif "TRƯA" in sn or "CA 2" in sn:
                    return dt.replace(hour=14, minute=0), dt.replace(hour=22, minute=0)
                elif "ĐÊM" in sn or "CA 3" in sn:
                    return dt.replace(hour=22, minute=0), (dt + timedelta(days=1)).replace(hour=6, minute=0)
                else:
                    return None, None

            def ceil_30(dt):
                mins = dt.minute
                if mins == 0: return dt
                if mins <= 30: return dt.replace(minute=30, second=0, microsecond=0)
                return (dt + timedelta(hours=1)).replace(minute=0, second=0, microsecond=0)

            def floor_30(dt):
                mins = dt.minute
                if mins == 0: return dt
                if mins < 30: return dt.replace(minute=0, second=0, microsecond=0)
                return dt.replace(minute=30, second=0, microsecond=0)

            def parse_time_with_temp_memory(t_val):
                if pd.isna(t_val): return None, ""
                s = str(t_val).strip()
                if s in ['nan', 'None', 'NaT', '']: return None, ""
                
                try:
                    temp_memory = {
                        'year': '', 'month': '', 'day': '',
                        'hour': '', 'minute': '', 'second': ''
                    }
                    
                    parts = s.split(' ')
                    if len(parts) >= 2:
                        d_part = parts[0]
                        t_part = parts[1]
                        
                        if '-' in d_part: d_split = d_part.split('-')
                        elif '/' in d_part: d_split = d_part.split('/')
                        else: d_split = []
                            
                        if len(d_split) == 3:
                            if len(d_split[0]) == 4: 
                                temp_memory['year'] = d_split[0]
                                temp_memory['month'] = d_split[1].zfill(2)
                                temp_memory['day'] = d_split[2].zfill(2)
                            elif len(d_split[2]) == 4: 
                                temp_memory['day'] = d_split[0].zfill(2)
                                temp_memory['month'] = d_split[1].zfill(2)
                                temp_memory['year'] = d_split[2]
                                
                        t_split = t_part.split(':')
                        if len(t_split) >= 2:
                            temp_memory['hour'] = t_split[0].zfill(2)
                            temp_memory['minute'] = t_split[1].zfill(2)
                                
                        if temp_memory['year'] and temp_memory['hour']:
                            final_str = f"{temp_memory['year']}-{temp_memory['month']}-{temp_memory['day']} {temp_memory['hour']}:{temp_memory['minute']}"
                            dt_obj = datetime(
                                int(temp_memory['year']), int(temp_memory['month']), int(temp_memory['day']), 
                                int(temp_memory['hour']), int(temp_memory['minute'])
                            )
                            return dt_obj, final_str
                except Exception:
                    pass
                return None, ""


            new_rows = []
            for idx, row in df.iterrows():
                msnv = str(row['Mã số nhân viên']).strip()
                raw_date = row['Ngày']
                shift_name = row['Số ca làm việc trong ngày']
                
                t_in_dt, t_in_str = parse_time_with_temp_memory(row['Thời gian quẹt thẻ lên ca'])
                t_out_dt, t_out_str = parse_time_with_temp_memory(row['Thời gian quẹt thẻ xuống ca'])

                date_out = str(raw_date).strip()
                if date_out in ['nan', 'NaT', 'None']: date_out = ""
                if date_out.endswith(" 00:00:00"): date_out = date_out.replace(" 00:00:00", "")

                true_date = None
                if t_in_dt:
                    true_date = t_in_dt.date()
                elif t_out_dt:
                    true_date = t_out_dt.date()
                    if "ĐÊM" in str(shift_name).upper() or "CA 3" in str(shift_name).upper():
                        if t_out_dt.hour < 12: 
                            true_date = true_date - timedelta(days=1)
                else:
                    try:
                        d_str_temp = str(raw_date).strip().split(' ')[0]
                        if '-' in d_str_temp: d_parts = d_str_temp.split('-')
                        elif '/' in d_str_temp: d_parts = d_str_temp.split('/')
                        else: d_parts = []
                        
                        if len(d_parts) == 3:
                            if len(d_parts[0]) == 4:
                                true_date = datetime(int(d_parts[0]), int(d_parts[1]), int(d_parts[2])).date()
                            elif len(d_parts[2]) == 4:
                                true_date = datetime(int(d_parts[2]), int(d_parts[1]), int(d_parts[0])).date()
                    except:
                        pass
                        
                if not true_date:
                    continue

                shift_start, shift_end = get_shift_times(true_date, shift_name)
                if not shift_start or not shift_end:
                    continue

                is_empty_in = t_in_dt is None
                is_empty_out = t_out_dt is None

                records_for_this_row = []

                if is_empty_in and is_empty_out:
                    records_for_this_row.append(('ACTR PHÉP NĂM', shift_start.strftime("%Y-%m-%d %H:%M"), shift_end.strftime("%Y-%m-%d %H:%M"), False))
                    
                elif is_empty_in and not is_empty_out:
                    records_for_this_row.append(('ACTR VIỆC RIÊNG', "Kiểm tra xếp ca/Bù vân tay", t_out_str, False))
                    
                elif not is_empty_in and is_empty_out:
                    records_for_this_row.append(('ACTR VIỆC RIÊNG', t_in_str, "Kiểm tra xếp ca/Bù vân tay", False))
                    
                else:
                    is_late = t_in_dt > shift_start
                    is_early = t_out_dt < shift_end

                    if is_late and is_early:
                        records_for_this_row.append(('ACTR VIỆC RIÊNG', shift_start.strftime("%Y-%m-%d %H:%M"), ceil_30(t_in_dt).strftime("%Y-%m-%d %H:%M"), False))
                        records_for_this_row.append(('ACTR VIỆC RIÊNG', floor_30(t_out_dt).strftime("%Y-%m-%d %H:%M"), shift_end.strftime("%Y-%m-%d %H:%M"), True))
                    elif is_late:
                        records_for_this_row.append(('ACTR VIỆC RIÊNG', shift_start.strftime("%Y-%m-%d %H:%M"), ceil_30(t_in_dt).strftime("%Y-%m-%d %H:%M"), False))
                    elif is_early:
                        records_for_this_row.append(('ACTR VIỆC RIÊNG', floor_30(t_out_dt).strftime("%Y-%m-%d %H:%M"), shift_end.strftime("%Y-%m-%d %H:%M"), False))
                    else:
                        records_for_this_row.append(('ACTR VIỆC RIÊNG', t_in_str, t_out_str, False))

                curr_shift = str(shift_name).upper()
                is_target_hc = "HC" in curr_shift or "HÀNH CHÍNH" in curr_shift
                next_date = true_date + timedelta(days=1)
                next_shift = shift_dict.get((msnv, next_date), "")
                is_next_ca3 = "CA 3" in next_shift or "ĐÊM" in next_shift

                if (is_empty_in or is_empty_out) and is_target_hc and is_next_ca3:
                    records_for_this_row = [('ACTR PHÉP NĂM', shift_start.strftime("%Y-%m-%d %H:%M"), shift_end.strftime("%Y-%m-%d %H:%M"), False)]

                for loai_phep, start_val, end_val, is_extra in records_for_this_row:
                    new_rows.append({
                        'MSNV': msnv,
                        'Ngày': date_out, 
                        '*Loại hình xin phép': loai_phep,
                        '*Thời gian bắt đầu xin phép': start_val,
                        '*Thời gian kết thúc xin phép': end_val,
                        'Lý do nghỉ phép': loai_phep,
                        'IsExtra': is_extra
                    })

            self.df_xp_result = pd.DataFrame(new_rows)
            
            if self.df_xp_result.empty:
                messagebox.showinfo("Thông báo", "Không tìm thấy dữ liệu hợp lệ cần xin phép!")
                return

            gui_headers = ['MSNV', 'Ngày', '*Loại hình xin phép', '*Thời gian bắt đầu xin phép', '*Thời gian kết thúc xin phép', 'Lý do nghỉ phép']
            
            self.tree_xp.delete(*self.tree_xp.get_children())
            self.tree_xp["column"] = gui_headers
            self.tree_xp["show"] = "headings"
            for col in self.tree_xp["columns"]:
                self.tree_xp.heading(col, text=col)
                self.tree_xp.column(col, width=150, anchor="center")

            for index, row in self.df_xp_result.iterrows():
                tag = 'evenrow' if index % 2 == 0 else 'oddrow'
                gui_row = [row[h] for h in gui_headers]
                self.tree_xp.insert("", "end", iid=str(index), values=gui_row, tags=(tag,))

            messagebox.showinfo("Thành công", f"Đã lọc và xử lý xong {len(self.df_xp_result)} trường hợp xin phép!")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra trong lúc xử lý:\n{str(e)}")

    def export_xp(self):
        import pandas as pd
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Border, Side, Font
        from openpyxl.worksheet.datavalidation import DataValidation
        from openpyxl.utils import get_column_letter

        if self.df_xp_result is None or self.df_xp_result.empty:
            messagebox.showwarning("Cảnh báo", "Bạn chưa Xử lý dữ liệu. Không có gì để xuất!")
            return
        
        try:
            now = datetime.now()
            date_str = now.strftime("%d_%m_%Y")
            default_name = f"Ket_Qua_Xin_Phep_{date_str}.xlsx"
            
            save_path = filedialog.asksaveasfilename(
                initialfile=default_name,
                defaultextension=".xlsx", 
                title="Lưu kết quả xin phép",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not save_path:
                return
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Báo Cáo"
            
            headers = [
                "MSNV", "Ngày", "*Loại hình xin phép", "*Thời gian bắt đầu xin phép", 
                "*Thời gian kết thúc xin phép", "Lý do nghỉ phép", "Ngày sinh con", 
                "Có phải là con một không", "结婚登记日期", "Sinh nhật của người già được chăm sóc", "Ghi chú"
            ]
            
            for c_idx, h_name in enumerate(headers, 1):
                cell = ws.cell(row=1, column=c_idx, value=h_name)
                cell.font = Font(bold=True)
                
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            
            start_row = 2
            for r_idx, row in self.df_xp_result.iterrows():
                current_row = start_row + r_idx
                is_extra = row.get('IsExtra', False)
                
                col_data = [
                    row['MSNV'], row['Ngày'], row['*Loại hình xin phép'], 
                    row['*Thời gian bắt đầu xin phép'], row['*Thời gian kết thúc xin phép'], 
                    row['Lý do nghỉ phép']
                ]
                
                for c_idx, val in enumerate(col_data):
                    val_str = str(val).strip()
                    if pd.isna(val) or val_str in ['nan', 'None', 'NaT', '']:
                        final_val = None
                    else:
                        final_val = val_str

                    cell = ws.cell(row=current_row, column=c_idx + 1, value=final_val)
                    
                    if final_val and "Kiểm tra xếp ca/Bù vân tay" in final_val:
                        cell.fill = yellow_fill
                    elif is_extra:
                        cell.fill = yellow_fill
            
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).border = thin_border
            
            max_r = ws.max_row
            if max_r >= 2:
                dv = DataValidation(
                    type="list", 
                    formula1='"ACTR VIỆC RIÊNG,ACTR PHÉP NĂM,ACTR NGHỈ BÙ,ACTR PHÉP BỆNH,ACTR KẾT HÔN,ACTR THAI SẢN NAM,ACTR PHÉP TANG"', 
                    allow_blank=True
                )
                ws.add_data_validation(dv)
                dv.add(f'C2:C{max_r}')
            
            for idx, col in enumerate(ws.columns, 1):
                max_length = 0
                column_letter = get_column_letter(idx)
                for cell in col:
                    try:
                        if cell.value is not None and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                ws.column_dimensions[column_letter].width = max_length + 2
                
            wb.save(save_path)
            messagebox.showinfo("Thành công", "Đã xuất file Xin Phép thành công với Form tự động hoàn hảo (đã gắn menu Drop-down)!")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi khi lưu file:\n{str(e)}")

    # =========================================================================
    # --- CÁC HÀM CŨ CHO TAB 1: CHẤM CÔNG ---
    # =========================================================================
    def select_file1(self):
        self.file1_path = filedialog.askopenfilename(filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv")])
        if self.file1_path:
            self.lbl_file1.config(text=self.file1_path.split('/')[-1])

    def select_file2(self):
        self.file2_path = filedialog.askopenfilename(filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv")])
        if self.file2_path:
            self.lbl_file2.config(text=self.file2_path.split('/')[-1])

    def read_file(self, path):
        import pandas as pd
        path_lower = path.lower()

        # 1. ĐỌC FILE CSV CHUẨN
        if path_lower.endswith('.csv'):
            encodings = ['utf-8', 'utf-8-sig', 'utf-16', 'utf-16le', 'cp1258', 'latin1']
            for enc in encodings:
                try:
                    return pd.read_csv(path, dtype=str, encoding=enc)
                except:
                    continue
            return pd.read_csv(path, dtype=str, encoding='latin1')

        # 2. XỬ LÝ FILE EXCEL VÀ "EXCEL GIẢ MẠO" TỪ ZKTIME
        try:
            # Thử đọc như file Excel bình thường trước
            return pd.read_excel(path, dtype=str)
        except Exception as e:
            # Rơi vào đây tức là Excel báo lỗi cấu trúc -> File giả mạo
            # Bật hệ thống NỘI SOI tàng hình: Đọc nó như một file Text/CSV
            encodings_to_try = ['utf-16le', 'utf-8', 'utf-8-sig', 'cp1258', 'latin1']
            seps_to_try = ['\t', ',', ';'] # Tab hoặc phẩy
            
            for enc in encodings_to_try:
                for sep in seps_to_try:
                    try:
                        df = pd.read_csv(path, sep=sep, encoding=enc, dtype=str, on_bad_lines='skip')
                        # Nếu đọc ra được nhiều hơn 3 cột -> Chắc chắn đúng định dạng
                        if len(df.columns) > 3:
                            return df
                    except:
                        continue
                        
            raise ValueError("Cấu trúc file ZKTeco quá lạ, hệ thống không thể tự động giải mã. Bạn hãy mở file bằng Excel và bấm Save As thành .xlsx chuẩn nhé!")


    def on_double_click(self, event):
        if self.df_result is None or self.df_result.empty: return
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = self.tree.identify_column(event.x)
        col_idx = int(column[1:]) - 1
        col_name = self.tree["columns"][col_idx]
        
        if col_name == 'Timetable':
            item_id = self.tree.focus()
            if not item_id: return
            
            idx = int(item_id)
            x, y, w, h = self.tree.bbox(item_id, column)
            current_val = self.tree.set(item_id, column)
            
            cb = ttk.Combobox(self.tree, values=["HC 7H30-17H00", "CA 1", "CA 2", "CA 3"], state="normal")
            cb.set(current_val)
            cb.place(x=x, y=y, width=w, height=h)
            cb.focus_set()
            
            def save_edit(event=None):
                new_val = cb.get()
                cb.destroy()
                if new_val != current_val:
                    self.df_result.at[idx, 'Timetable'] = new_val
                    self.process_single_row(idx, force_recalc=True)
                    self.tree.item(item_id, values=list(self.df_result.loc[idx]))
                    
                    unique_timetables = sorted([str(t) for t in self.df_result['Timetable'].unique() if str(t).strip() != '' and str(t) != 'nan'])
                    self.cb_timetable['values'] = ["Tất cả"] + unique_timetables
            cb.bind("<Return>", save_edit)
            cb.bind("<FocusOut>", save_edit)

    def reset_filter(self):
        if self.df_result is None or self.df_result.empty: return
        self.cb_date.selected = []
        self.cb_msnv.selected = []
        self.cb_timetable.current(0)
        self.cb_worktime.current(0)
        self.apply_filter()

    def apply_filter(self, event=None):
        if self.df_result is None or self.df_result.empty: return
        
        selected_dates = self.cb_date.get()
        selected_msnvs = self.cb_msnv.get()
        selected_timetable = self.cb_timetable.get()
        selected_worktime = self.cb_worktime.get()
        
        df_to_show = self.df_result.copy()
        
        if selected_dates != ["Tất cả"]:
            df_to_show = df_to_show[df_to_show['Date'].isin(selected_dates)]
        if selected_msnvs != ["Tất cả"]:
            df_to_show = df_to_show[df_to_show['MSNV'].isin(selected_msnvs)]
            
        if selected_timetable != "Tất cả" and selected_timetable != "":
            df_to_show = df_to_show[df_to_show['Timetable'] == selected_timetable]
        if selected_worktime == "Bù vân tay":
            df_to_show = df_to_show[df_to_show['Work Time'] == "Bù vân tay"]
            
        self.display_data(df_to_show)

        def get_dynamic_opts(col_name):
            df_f = self.df_result.copy()
            if col_name != 'Date' and selected_dates != ["Tất cả"]:
                df_f = df_f[df_f['Date'].isin(selected_dates)]
            if col_name != 'MSNV' and selected_msnvs != ["Tất cả"]:
                df_f = df_f[df_f['MSNV'].isin(selected_msnvs)]
            if col_name != 'Timetable' and selected_timetable not in ["Tất cả", ""]:
                df_f = df_f[df_f['Timetable'] == selected_timetable]
            if col_name != 'Work Time' and selected_worktime == "Bù vân tay":
                df_f = df_f[df_f['Work Time'] == "Bù vân tay"]
                
            return sorted([str(x) for x in df_f[col_name].unique() if str(x).strip() != '' and str(x) != 'nan'])

        self.cb_date.set_values(get_dynamic_opts('Date'))
        self.cb_msnv.set_values(get_dynamic_opts('MSNV'))
        
        opts_timetable = ["Tất cả"] + get_dynamic_opts('Timetable')
        self.cb_timetable['values'] = opts_timetable
        if selected_timetable not in opts_timetable:
            self.cb_timetable.current(0)


    def process_data(self):
        import pandas as pd
        if not self.file1_path or not self.file2_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn đủ cả 2 file!")
            return

        try:
            df1 = self.read_file(self.file1_path)
            df2 = self.read_file(self.file2_path)

            if 'AC-No.' not in df1.columns:
                messagebox.showerror("Lỗi dữ liệu", "Không tìm thấy cột 'AC-No.' trong File 1!")
                return
                
            df1['AC-No.'] = df1['AC-No.'].astype(str).str.strip()
            df1['MSNV'] = 'S' + df1['AC-No.'].str[2:]
            
            cols_to_keep = ['MSNV', 'Name', 'Date', 'Timetable', 'On duty', 'Off duty', 'Clock In', 'Clock Out', 'Late', 'Early', 'OT Time', 'Work Time']
            for col in cols_to_keep:
                if col not in df1.columns: df1[col] = ''
            
            self.df1_clean = df1[cols_to_keep].copy()
            self.df1_clean.replace('nan', '', inplace=True)
            self.df1_clean.fillna('', inplace=True)
            
            def convert_to_dmy(date_str):
                try:
                    dt_obj = pd.to_datetime(date_str).date()
                    return dt_obj.strftime("%d/%m/%Y")
                except:
                    return str(date_str)
                    
            self.df1_clean['Date'] = self.df1_clean['Date'].apply(convert_to_dmy)

            if 'Mã số nhân viên' not in df2.columns or 'Thời gian quẹt thẻ' not in df2.columns:
                messagebox.showerror("Lỗi dữ liệu", "File 2 không đúng chuẩn (Thiếu cột 'Mã số nhân viên' hoặc 'Thời gian quẹt thẻ')!")
                return
                
            df2['Mã số nhân viên'] = df2['Mã số nhân viên'].astype(str).str.strip()
            df2['MSNV_Match'] = df2['Mã số nhân viên'].apply(lambda x: x if x.startswith('S') else 'S' + x.zfill(6))
            df2['Thời gian quẹt thẻ'] = pd.to_datetime(df2['Thời gian quẹt thẻ'], errors='coerce')
            self.df2_clean = df2

            self.df_result = self.df1_clean.copy()
            
            self.df_result['Kết Quả Công'] = pd.Series(dtype=object)
            self.df_result['Nghỉ phép'] = pd.Series(dtype=object)
            self.df_result['Tăng ca'] = pd.Series(dtype=object)

            for idx in self.df_result.index:
                self.process_single_row(idx, force_recalc=False)

            def sort_dmy(dmy_list):
                valid = []
                for d in dmy_list:
                    try: valid.append(datetime.strptime(d, "%d/%m/%Y").date())
                    except: pass
                return [d.strftime("%d/%m/%Y") for d in sorted(valid)]
                
            unique_dates = sort_dmy([str(d) for d in self.df_result['Date'].unique() if str(d).strip() != '' and str(d) != 'nan'])
            self.cb_date.set_values(unique_dates)
            
            unique_msnvs = sorted([str(m) for m in self.df_result['MSNV'].unique() if str(m).strip() != '' and str(m) != 'nan'])
            self.cb_msnv.set_values(unique_msnvs)
            
            unique_timetables = sorted([str(t) for t in self.df_result['Timetable'].unique() if str(t).strip() != '' and str(t) != 'nan'])
            self.cb_timetable['values'] = ["Tất cả"] + unique_timetables
            self.cb_timetable.current(0)
            
            self.cb_worktime.current(0)
            
            self.apply_filter()
            messagebox.showinfo("Thành công", "Xử lý dữ liệu hoàn tất!")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi trong quá trình xử lý:\n{str(e)}")

    def process_single_row(self, idx, force_recalc=False):
        orig_row = self.df1_clean.loc[idx]
        c_in = str(orig_row['Clock In']).strip()
        c_out = str(orig_row['Clock Out']).strip()
        msnv = str(orig_row['MSNV']).strip()
        
        timetable = str(self.df_result.at[idx, 'Timetable']).strip()
        tt_upper = timetable.upper()
        
        is_hc = "HC" in tt_upper or "HÀNH CHÍNH" in tt_upper
        is_ca1 = "CA 1" in tt_upper or "SÁNG" in tt_upper
        is_ca2 = "CA 2" in tt_upper or "TRƯA" in tt_upper
        is_ca3 = "CA 3" in tt_upper or "ĐÊM" in tt_upper

        if is_hc:
            if "15H30" in tt_upper or "15:30" in tt_upper:
                on_duty, off_duty = "07:30", "15:30"
            else:
                on_duty, off_duty = "07:30", "17:00"
        elif is_ca1:
            on_duty, off_duty = "06:00", "14:00"
        elif is_ca2:
            on_duty, off_duty = "14:00", "22:00"
        elif is_ca3:
            on_duty, off_duty = "22:00", "06:00"
        else:
            on_duty = str(orig_row['On duty']).strip()
            off_duty = str(orig_row['Off duty']).strip()
            
        self.df_result.at[idx, 'On duty'] = on_duty
        self.df_result.at[idx, 'Off duty'] = off_duty

        available_times = []
        if c_in != '' and c_in != 'Bù vân tay' and str(c_in) != 'nan': available_times.append(c_in)
        if c_out != '' and c_out != 'Bù vân tay' and str(c_out) != 'nan': available_times.append(c_out)

        c_in, c_out = '', ''
        if len(available_times) == 2:
            times = sorted(available_times)
            if is_ca3:
                c_in, c_out = times[1], times[0]
            else:
                c_in, c_out = times[0], times[1]
        elif len(available_times) == 1:
            t = available_times[0]
            try:
                fmt = "%H:%M"
                t_dt = datetime.strptime(t[:5], fmt)
                on_dt = datetime.strptime(on_duty, fmt)
                off_dt = datetime.strptime(off_duty, fmt)
                
                def get_dist(d1, d2):
                    return min(abs((d1 - d2).total_seconds()), 
                               abs((d1 + timedelta(days=1) - d2).total_seconds()), 
                               abs((d1 - timedelta(days=1) - d2).total_seconds()))
                
                if get_dist(t_dt, on_dt) <= get_dist(t_dt, off_dt):
                    c_in = t
                else:
                    c_out = t
            except:
                c_in = t

        self.df_result.at[idx, 'Clock In'] = c_in
        self.df_result.at[idx, 'Clock Out'] = c_out
        
        if force_recalc:
            self.df_result.at[idx, 'Work Time'] = ''
            self.df_result.at[idx, 'Late'] = ''
            self.df_result.at[idx, 'Early'] = ''
            self.df_result.at[idx, 'OT Time'] = ''
        else:
            wt = str(orig_row['Work Time']).strip()
            self.df_result.at[idx, 'Work Time'] = wt if wt != 'nan' else ''
            late_v = str(orig_row['Late']).strip() if 'Late' in orig_row else ''
            self.df_result.at[idx, 'Late'] = late_v if late_v != 'nan' else ''
            early_v = str(orig_row['Early']).strip() if 'Early' in orig_row else ''
            self.df_result.at[idx, 'Early'] = early_v if early_v != 'nan' else ''
            ot_v = str(orig_row['OT Time']).strip() if 'OT Time' in orig_row else ''
            self.df_result.at[idx, 'OT Time'] = ot_v if ot_v != 'nan' else ''
            
        self.df_result.at[idx, 'Kết Quả Công'] = ''
        self.df_result.at[idx, 'Nghỉ phép'] = ''
        self.df_result.at[idx, 'Tăng ca'] = ''
        
        try:
            date_val = datetime.strptime(orig_row['Date'], "%d/%m/%Y").date()
        except:
            date_val = None

        if date_val and (c_in == '' or c_out == ''):
            df2 = self.df2_clean
            if c_in == '' and c_out == '':
                pass
            else:
                try:
                    fmt = "%H:%M"
                    on_dt_time = datetime.strptime(on_duty, fmt).time()
                    off_dt_time = datetime.strptime(off_duty, fmt).time()
                    
                    if is_ca3:
                        target_in_dt = datetime.combine(date_val - timedelta(days=1), on_dt_time)
                        target_out_dt = datetime.combine(date_val, off_dt_time)
                        if target_out_dt <= target_in_dt: 
                            target_out_dt += timedelta(days=1)
                    else:
                        target_in_dt = datetime.combine(date_val, on_dt_time)
                        target_out_dt = datetime.combine(date_val, off_dt_time)
                        if target_out_dt <= target_in_dt: 
                            target_out_dt += timedelta(days=1)
                        
                    emp_swipes = df2.loc[df2['MSNV_Match'] == msnv, 'Thời gian quẹt thẻ'].dropna()
                    
                    if c_in == '' and c_out != '':
                        if is_ca3:
                            valid_ins = emp_swipes[(emp_swipes >= target_in_dt - timedelta(hours=4)) & (emp_swipes <= target_out_dt)]
                        else:
                            start_bound = datetime.combine(date_val, datetime.min.time())
                            valid_ins = emp_swipes[(emp_swipes >= max(target_in_dt - timedelta(hours=4), start_bound)) & (emp_swipes <= target_out_dt)]
                            
                        if not valid_ins.empty:
                            new_in = valid_ins.min().strftime("%H:%M")
                            if new_in != c_out: self.df_result.at[idx, 'Clock In'] = new_in
                            
                    elif c_out == '' and c_in != '':
                        if is_ca3:
                            valid_outs = emp_swipes[(emp_swipes >= target_in_dt) & (emp_swipes <= target_out_dt + timedelta(hours=8))]
                        else:
                            end_bound = datetime.combine(date_val, datetime.max.time())
                            valid_outs = emp_swipes[(emp_swipes >= target_in_dt) & (emp_swipes <= min(target_out_dt + timedelta(hours=8), end_bound))]
                            
                        if not valid_outs.empty:
                            new_out = valid_outs.max().strftime("%H:%M")
                            if new_out != c_in: self.df_result.at[idx, 'Clock Out'] = new_out
                except Exception as e:
                    pass

        c_in = self.df_result.at[idx, 'Clock In']
        c_out = self.df_result.at[idx, 'Clock Out']
        
        if c_in == '' and c_out == '':
            self.df_result.at[idx, 'Nghỉ phép'] = "PLP"
        elif c_in == '' and c_out != '':
            self.df_result.at[idx, 'Clock In'] = "Bù vân tay"
            self.df_result.at[idx, 'Work Time'] = "Bù vân tay"
        elif c_in != '' and c_out == '':
            self.df_result.at[idx, 'Clock Out'] = "Bù vân tay"
            self.df_result.at[idx, 'Work Time'] = "Bù vân tay"

        w_time = str(self.df_result.at[idx, 'Work Time']).strip()
        needs_recalc = force_recalc
        if w_time == '' or w_time == 'Bù vân tay':
            needs_recalc = True
        elif is_hc and c_in != '' and c_out != '' and c_in != 'Bù vân tay' and c_out != 'Bù vân tay':
            try:
                fmt = "%H:%M"
                tin_dt = datetime.strptime(c_in[:5], fmt)
                tout_dt = datetime.strptime(c_out[:5], fmt)
                if tout_dt < tin_dt: tout_dt += timedelta(days=1)
                if (tout_dt - tin_dt).total_seconds() <= 5.5 * 3600:
                    needs_recalc = True
            except: pass

        if needs_recalc and c_in != '' and c_out != '' and c_in != 'Bù vân tay' and c_out != 'Bù vân tay':
            try:
                fmt = "%H:%M"
                tin_dt = datetime.strptime(c_in[:5], fmt)
                tout_dt = datetime.strptime(c_out[:5], fmt)
                if tout_dt < tin_dt: tout_dt += timedelta(days=1)
                
                total_sec = int((tout_dt - tin_dt).total_seconds())
                
                if is_hc and (off_duty == "17:00" or "17H00" in off_duty):
                    lunch_start = tin_dt.replace(hour=11, minute=30, second=0)
                    lunch_end = tin_dt.replace(hour=13, minute=0, second=0)
                    overlap_start = max(tin_dt, lunch_start)
                    overlap_end = min(tout_dt, lunch_end)
                    if overlap_end > overlap_start:
                        total_sec -= int((overlap_end - overlap_start).total_seconds())
                
                if total_sec >= 0:
                    if total_sec > 8 * 3600:
                        self.df_result.at[idx, 'Work Time'] = "08:00"
                        ot_sec = total_sec - 8 * 3600
                        self.df_result.at[idx, 'OT Time'] = f"{ot_sec // 3600:02d}:{(ot_sec % 3600) // 60:02d}"
                        self.df_result.at[idx, 'Late'] = ''
                        self.df_result.at[idx, 'Early'] = ''
                    else:
                        self.df_result.at[idx, 'Work Time'] = f"{total_sec // 3600:02d}:{(total_sec % 3600) // 60:02d}"
                        self.df_result.at[idx, 'OT Time'] = ''
                        late_val, early_val = "", ""
                        
                        if on_duty != '' and off_duty != '':
                            try:
                                dt_on = datetime.strptime(on_duty[:5], fmt)
                                dt_off = datetime.strptime(off_duty[:5], fmt)
                                dt_in = datetime.strptime(c_in[:5], fmt)
                                dt_out = datetime.strptime(c_out[:5], fmt)
                                
                                if dt_off < dt_on: dt_off += timedelta(days=1)
                                if dt_in < dt_on and (dt_on - dt_in).total_seconds() > 12 * 3600: dt_in += timedelta(days=1)
                                elif dt_in > dt_on and (dt_in - dt_on).total_seconds() > 12 * 3600: dt_in -= timedelta(days=1)
                                if dt_out < dt_on and (dt_on - dt_out).total_seconds() > 12 * 3600: dt_out += timedelta(days=1)
                                    
                                late_sec = int((dt_in - dt_on).total_seconds())
                                if late_sec > 0: late_val = f"{late_sec // 3600:02d}:{(late_sec % 3600) // 60:02d}"
                                    
                                early_sec = int((dt_off - dt_out).total_seconds())
                                if early_sec > 0: early_val = f"{early_sec // 3600:02d}:{(early_sec % 3600) // 60:02d}"
                            except: pass
                                
                        self.df_result.at[idx, 'Late'] = late_val
                        self.df_result.at[idx, 'Early'] = early_val
                else:
                    self.df_result.at[idx, 'Work Time'] = "00:00"
            except: pass
        
        w_time = str(self.df_result.at[idx, 'Work Time']).strip()
        cham_cong_val = ""
        
        if w_time != '' and w_time != 'Bù vân tay':
            try:
                h_str, m_str = w_time.split(":")
                h = int(h_str)
                m = int(m_str)
                
                phangio = str(h)
                phanphut = ".5" if m >= 30 else ""
                phanduoi = ""
                
                if is_ca1: phanduoi = "C1"
                elif is_ca2: phanduoi = "C2"
                elif is_ca3: phanduoi = "C3"
                elif is_hc: phanduoi = ""
                
                if phanduoi == "":
                    val = h + (0.5 if m >= 30 else 0)
                    cham_cong_val = str(int(val)) if val == int(val) else str(val)
                else:
                    cham_cong_val = phangio + phanphut + phanduoi
                
                if self.df_result.at[idx, 'Nghỉ phép'] != "PLP":
                    num_work = h + (0.5 if m >= 30 else 0)
                    if num_work < 8:
                        np_val = 8 - num_work
                        self.df_result.at[idx, 'Nghỉ phép'] = f"{np_val:g}VR"
            except Exception as e:
                pass
        
        if w_time == 'Bù vân tay':
            cham_cong_val = "Bù vân tay"
            
        self.df_result.at[idx, 'Kết Quả Công'] = cham_cong_val

    def display_data(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["column"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=90, anchor="center")

        for index, row in df.iterrows():
            tag = 'evenrow' if index % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", iid=str(index), values=list(row), tags=(tag,))

    def get_default_filename(self, prefix="Bao_Cao"):
        now = datetime.now()
        date_str = now.strftime("%d_%m_%Y")
        return f"{prefix}_{date_str}.xlsx"

    def export_excel(self):
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Border, Side
        
        if self.df_result is None or self.df_result.empty:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
            return
            
        default_name = self.get_default_filename("Bao_Cao_Cham_Cong")
        save_path = filedialog.asksaveasfilename(
            initialfile=default_name,
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        if save_path:
            try:
                df_export = self.df_result.copy()
                
                for col in df_export.columns:
                    df_export[col] = df_export[col].apply(lambda x: None if str(x).strip() in ['nan', 'None', 'NaT', ''] else x)
                
                df_export['Kết Quả Công'] = pd.to_numeric(df_export['Kết Quả Công'], errors='ignore')
                df_export.to_excel(save_path, index=False)
                
                wb = load_workbook(save_path)
                ws = wb.active
                
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                fill_color = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
                for r in range(1, ws.max_row + 1):
                    for c in range(1, ws.max_column + 1):
                        ws.cell(row=r, column=c).border = thin_border
                
                headers = [cell.value for cell in ws[1]]
                col_indices = {}
                for col_name in ['Late', 'Early', 'OT Time', 'Work Time']:
                    if col_name in headers:
                        col_indices[col_name] = headers.index(col_name) + 1
                        
                kq_col_idx = None
                if 'Kết Quả Công' in headers:
                    kq_col_idx = headers.index('Kết Quả Công') + 1
                
                for r in range(2, ws.max_row + 1):
                    for col_name, c_idx in col_indices.items():
                        cell = ws.cell(row=r, column=c_idx)
                        val = str(cell.value).strip() if cell.value is not None else ""
                        
                        if val != "" and val != "nan" and val != "None":
                            if col_name == 'Work Time':
                                if val != '08:00': cell.fill = fill_color
                            else:
                                cell.fill = fill_color
                                
                    if kq_col_idx:
                        kq_cell = ws.cell(row=r, column=kq_col_idx)
                        if kq_cell.value is not None:
                            v_str = str(kq_cell.value).strip()
                            if v_str not in ["", "nan", "None"]:
                                try:
                                    f_v = float(v_str)
                                    kq_cell.value = int(f_v) if f_v.is_integer() else f_v
                                except:
                                    pass
                            else:
                                kq_cell.value = None
                
                wb.save(save_path)
                messagebox.showinfo("Thành công", "Xuất bảng Excel và định dạng số chuẩn thành công!")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu và định dạng file:\n{str(e)}")

    def export_fixed_form(self):
        import pandas as pd
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
        from openpyxl.utils import get_column_letter

        if self.df_result is None or self.df_result.empty:
            messagebox.showwarning("Cảnh báo", "Bạn chưa xử lý dữ liệu!")
            return

        default_name = self.get_default_filename("Form_Bao_Cao_Tong_Hop")
        save_path = filedialog.asksaveasfilename(
            initialfile=default_name,
            defaultextension=".xlsx", 
            title="Lưu báo cáo form cố định", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not save_path:
            return
            
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Báo Cáo"
            
            raw_dates = [d for d in self.df_result['Date'].unique() if str(d).strip() != '' and str(d) != 'nan']
            
            parsed_dates = []
            for d in raw_dates:
                try:
                    dt_obj = datetime.strptime(d, "%d/%m/%Y").date()
                    parsed_dates.append(dt_obj)
                except Exception:
                    pass
            
            sorted_dates = sorted(list(set(parsed_dates)))
            
            emp_list = self.df_result[['MSNV', 'Name']].drop_duplicates().to_dict('records')
            
            ws.cell(row=1, column=1, value="MSNV").font = Font(bold=True)
            ws.cell(row=1, column=2, value="Họ Tên").font = Font(bold=True)
            ws.cell(row=2, column=1, value="").font = Font(bold=True)
            ws.cell(row=2, column=2, value="").font = Font(bold=True)
            
            col_idx = 3
            for dt_obj in sorted_dates:
                day_str = str(dt_obj.day)
                
                cell_date = ws.cell(row=1, column=col_idx, value=day_str)
                cell_date.font = Font(bold=True)
                cell_date.alignment = Alignment(horizontal="center")
                ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx+1)
                
                ws.cell(row=2, column=col_idx, value="Kết Quả Công").font = Font(bold=True)
                ws.cell(row=2, column=col_idx+1, value="Nghỉ phép").font = Font(bold=True)
                
                col_idx += 2
                
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
            row_idx = 3
            for emp in emp_list:
                msnv = emp['MSNV']
                name = emp['Name']
                
                df_emp = self.df_result[self.df_result['MSNV'] == msnv]
                
                ws.cell(row=row_idx, column=1, value=msnv)
                ws.cell(row=row_idx, column=2, value=name)
                
                c_idx = 3
                for dt_obj in sorted_dates:
                    target_date_str = dt_obj.strftime("%d/%m/%Y")
                    record = df_emp[df_emp['Date'] == target_date_str]
                    
                    if not record.empty:
                        kq_cong_raw = record.iloc[0]['Kết Quả Công']
                        kq_val = None
                        if pd.isna(kq_cong_raw) or str(kq_cong_raw).strip() in ['nan', 'None', 'NaT', '']:
                            kq_val = None
                        else:
                            try:
                                f_val = float(kq_cong_raw)
                                kq_val = int(f_val) if f_val.is_integer() else f_val
                            except:
                                kq_val = str(kq_cong_raw).strip()
                        
                        nghi_phep_raw = record.iloc[0]['Nghỉ phép']
                        if pd.isna(nghi_phep_raw) or str(nghi_phep_raw).strip() in ['nan', 'None', 'NaT', '']:
                            nghi_phep = None
                        else:
                            nghi_phep = str(nghi_phep_raw).strip()
                        
                        cell_kq = ws.cell(row=row_idx, column=c_idx, value=kq_val)
                        if str(kq_val) == "Bù vân tay":
                            cell_kq.fill = yellow_fill
                            
                        ws.cell(row=row_idx, column=c_idx+1, value=nghi_phep)
                    c_idx += 2
                row_idx += 1
                
            for idx, col in enumerate(ws.columns, 1):
                max_length = 0
                column_letter = get_column_letter(idx)
                for cell in col:
                    try:
                        if cell.value is not None and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                ws.column_dimensions[column_letter].width = max_length + 2
                
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).border = thin_border
                    
            wb.save(save_path)
            messagebox.showinfo("Thành công", "Đã xuất báo cáo theo Form cố định thành công!")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi xuất Form:\n{str(e)}")


# =========================================================================
# --- HÀM KIỂM TRA MUTEX (CHỐNG MỞ ỨNG DỤNG NHIỀU LẦN CÙNG LÚC) ---
# =========================================================================
def check_single_instance():
    if os.name == 'nt':
        mutex_name = "iamhisu_timekeep_mutex_xyz123"
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
        if ctypes.windll.kernel32.GetLastError() == 183:
            sys.exit(0)
        return mutex
    return None

# =========================================================================
# --- HÀM LUỒNG NGẦM ĐỂ TẢI THƯ VIỆN NẶNG (TRÁNH ĐƠ ẢNH GIF) ---
# =========================================================================
def load_heavy_libraries():
    try:
        import pandas
        import openpyxl
    except:
        pass

# =========================================================================
# --- HÀM TẠO MÀN HÌNH SPLASH ĐỘNG (CÓ HÌNH GIF & % LOADING) ---
# =========================================================================
def show_splash_and_start():
    global global_mutex
    global_mutex = check_single_instance()

    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.attributes('-topmost', True)

    width = 400
    height = 320
    x = (splash.winfo_screenwidth() // 2) - (width // 2)
    y = (splash.winfo_screenheight() // 2) - (height // 2)
    splash.geometry(f'{width}x{height}+{x}+{y}')
    splash.configure(bg="#2b2b2b")

    # 1. Đọc và Nạp khung hình của file loading.gif
    lbl_gif = tk.Label(splash, bg="#2b2b2b")
    lbl_gif.pack(pady=(20, 10))
    
    frames = []
    gif_path = resource_path("loading.gif")
    if os.path.exists(gif_path):
        try:
            i = 0
            while True:
                frames.append(tk.PhotoImage(file=gif_path, format=f"gif -index {i}"))
                i += 1
        except tk.TclError:
            pass 

    def animate_gif(ind):
        if frames:
            lbl_gif.configure(image=frames[ind])
            splash.after(50, animate_gif, (ind + 1) % len(frames))

    if frames:
        animate_gif(0)
    else:
        lbl_gif.config(text="[Đang xử lý hình ảnh...]", fg="white", font=("Arial", 10, "italic"))

    # 2. Trang trí Giao diện Splash (Chữ và Thanh Process)
    tk.Label(splash, text="TIMEKEEP (iamhisu)", font=("Arial", 16, "bold"), bg="#2b2b2b", fg="#00ffcc").pack(pady=5)
    lbl_status = tk.Label(splash, text="Đang khởi tạo hệ thống...", font=("Arial", 10), bg="#2b2b2b", fg="white")
    lbl_status.pack()

    frame_prog = tk.Frame(splash, bg="#2b2b2b")
    frame_prog.pack(pady=10)

    progress = ttk.Progressbar(frame_prog, orient="horizontal", length=250, mode="determinate")
    progress.pack(side=tk.LEFT, padx=5)

    lbl_pct = tk.Label(frame_prog, text="0%", font=("Arial", 10, "bold"), bg="#2b2b2b", fg="#00ffcc")
    lbl_pct.pack(side=tk.LEFT)

    splash.update() 

    # 3. Logic Đa luồng: Vừa quay GIF, vừa tải Pandas, vừa cộng %
    def check_loading_thread(thread):
        if thread.is_alive():
            current = progress['value']
            if current < 90:
                new_val = current + 1.5
                progress['value'] = new_val
                lbl_pct.config(text=f"{int(new_val)}%")
            splash.after(100, check_loading_thread, thread)
        else:
            progress['value'] = 100
            lbl_pct.config(text="100%")
            lbl_status.config(text="Hoàn tất! Đang mở ứng dụng...")
            splash.after(400, launch_main_app)

    def start_background_loading():
        lbl_status.config(text="Đang nạp dữ liệu lõi (có thể mất 15s)...")
        progress['value'] = 10
        lbl_pct.config(text="10%")
        
        thread = threading.Thread(target=load_heavy_libraries)
        thread.start()
        splash.after(100, check_loading_thread, thread)

    def launch_main_app():
        root = tk.Tk()
        root.withdraw()
        app = TimekeepApp(root)
        splash.destroy()
        root.deiconify()
        root.mainloop()

    splash.after(200, start_background_loading)
    splash.mainloop()

# =========================================================================
# --- ĐOẠN CODE KÍCH HOẠT NHẬP PASS (HIỆN ĐANG TẮT - BỎ COMMENT ĐỂ DÙNG) ---
# =========================================================================
# def show_login():
#     try:
#         myappid = 'iamhisu.timekeep.app.1.0'
#         ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
#     except:
#         pass
#
#     root = tk.Tk()
#     root.withdraw()
#     
#     login_win = tk.Toplevel(root)
#     login_win.title("iamhisu - Đăng nhập")
#     login_win.geometry("300x160")
#     login_win.resizable(False, False)
#     
#     try:
#         icon_path = resource_path('noface.ico')
#         login_win.iconbitmap(icon_path)
#     except Exception:
#         pass
#     
#     login_win.update_idletasks()
#     width = 300
#     height = 160
#     x = (login_win.winfo_screenwidth() // 2) - (width // 2)
#     y = (login_win.winfo_screenheight() // 2) - (height // 2)
#     login_win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
#     
#     login_win.protocol("WM_DELETE_WINDOW", root.destroy)
#     
#     tk.Label(login_win, text="ID đăng nhập:", font=("Arial", 10)).pack(pady=(15, 5))
#     entry_id = tk.Entry(login_win, font=("Arial", 10), justify="center")
#     entry_id.pack()
#     
#     tk.Label(login_win, text="Mật khẩu:", font=("Arial", 10)).pack(pady=(5, 5))
#     entry_pass = tk.Entry(login_win, show="*", font=("Arial", 10), justify="center")
#     entry_pass.pack()
#     
#     def check_login(event=None):
#         if entry_id.get() == "admin" and entry_pass.get() == "3042":
#             login_win.destroy()
#             show_splash_and_start() # Hiện màn hình load sau khi đăng nhập đúng
#         else:
#             messagebox.showerror("Lỗi", "Sai ID hoặc mật khẩu!", parent=login_win)
#             
#     login_win.bind('<Return>', check_login)
#     tk.Button(login_win, text="Đăng nhập", command=check_login, bg="lightblue", font=("Arial", 10, "bold")).pack(pady=10)
#     
#     root.mainloop()

if __name__ == "__main__":
    # NẾU MUỐN DÙNG CHỨC NĂNG ĐĂNG NHẬP: 
    # Thay dòng show_splash_and_start() bên dưới bằng chữ show_login()
    show_splash_and_start()
