import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
import os

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
            # Giữ lại các tùy chọn cũ nếu nó vẫn còn tồn tại trong danh sách mới
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
        
        # Căn chỉnh vị trí ngay dưới nút bấm
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        popup.geometry(f"+{x}+{y}")
        
        # Thanh tìm kiếm (Gõ để lọc nhanh)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(popup, textvariable=search_var)
        search_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # Khung chứa Checkbox có thanh cuộn
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
                self.command() # Tự động chạy apply_filter sau khi xác nhận
        
        btn_ok = ttk.Button(popup, text="Xác nhận", command=on_confirm)
        btn_ok.pack(pady=5, fill=tk.X, padx=10)
        
        search_entry.focus_set()


class TimekeepApp:
    def __init__(self, root):
        self.root = root
        self.root.title("iamhisu")
        self.root.geometry("1250x650")
        
        try:
            self.root.iconbitmap('fr.ico')
        except Exception:
            pass
        
        self.root.update_idletasks()
        width = 1250
        height = 680
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        self.file1_path = ""
        self.file2_path = ""
        self.df1_clean = None
        self.df2_clean = None
        self.df_result = None
        self.current_theme = "Sáng"

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
                        background="white",
                        foreground="black",
                        rowheight=25, 
                        fieldbackground="white")
        style.map('Treeview', background=[('selected', '#347083')])
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'), background="#d3d3d3")

        frame_top = tk.Frame(self.root, pady=5)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="1. File zkt:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.lbl_file1 = tk.Label(frame_top, text="Chưa chọn file", width=40, anchor="w", bg="white", relief="sunken")
        self.lbl_file1.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(frame_top, text="Chọn File", command=self.select_file1).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(frame_top, text="2. Dữ liệu HR:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.lbl_file2 = tk.Label(frame_top, text="Chưa chọn file", width=40, anchor="w", bg="white", relief="sunken")
        self.lbl_file2.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(frame_top, text="Chọn File", command=self.select_file2).grid(row=1, column=2, padx=5, pady=5)

        tk.Button(frame_top, text="Bắt đầu Xử lý", command=self.process_data, bg="lightblue", font=("Arial", 10, "bold")).grid(row=0, column=3, rowspan=2, padx=20, ipadx=10, ipady=10)
        
        frame_hints = tk.Frame(frame_top)
        frame_hints.grid(row=0, column=4, rowspan=2, padx=10, sticky="w")
        tk.Label(frame_hints, text="(MẸO: Nhấp đúp chuột vào cột Timetable để sửa tên Ca và tính lại dữ liệu)", fg="red", font=("Arial", 9, "italic")).pack(anchor="w")
        tk.Label(frame_hints, text="(HƯỚNG DẪN: Hãy mở file ZKT bằng Excel và chọn Save As thành định dạng .xlsx chuẩn)", fg="red", font=("Arial", 9, "italic")).pack(anchor="w")

        frame_settings = tk.Frame(frame_top)
        frame_settings.grid(row=0, column=5, rowspan=2, padx=20, sticky="e")
        
        self.btn_theme = tk.Button(frame_settings, text="🎨 Giao diện: Sáng", command=self.toggle_theme, bg="#e0e0e0", font=("Arial", 9, "bold"))
        self.btn_theme.pack(fill=tk.X, pady=2)
        
        frame_opacity = tk.Frame(frame_settings)
        frame_opacity.pack(fill=tk.X, pady=2)
        self.lbl_opacity = tk.Label(frame_opacity, text="Độ trong suốt:", font=("Arial", 8))
        self.lbl_opacity.pack(side=tk.LEFT)
        self.opacity_slider = ttk.Scale(frame_opacity, from_=0.2, to=1.0, value=1.0, command=self.change_opacity)
        self.opacity_slider.pack(side=tk.LEFT, padx=5)

        frame_filter = tk.Frame(self.root, pady=5)
        frame_filter.pack(fill=tk.X, padx=10)
        
        # SỬ DỤNG CLASS ĐA CHỌN MỚI TẠO CHO NGÀY VÀ MSNV
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

        # Ràng buộc chạy lọc khi chọn Ca và WorkTime bằng Combobox chuẩn
        self.cb_timetable.bind("<<ComboboxSelected>>", self.apply_filter)
        self.cb_worktime.bind("<<ComboboxSelected>>", self.apply_filter)

        frame_middle = tk.Frame(self.root)
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

        frame_bottom = tk.Frame(self.root, pady=10)
        frame_bottom.pack(fill=tk.X)
        tk.Button(frame_bottom, text="Xuất dữ liệu ra Excel", command=self.export_excel, bg="lightgreen", font=("Arial", 12, "bold")).pack(pady=5)
        
        self.apply_theme("Sáng")

    def change_opacity(self, val):
        self.root.attributes('-alpha', float(val))

    def toggle_theme(self):
        themes = ["Sáng", "Tối", "Trong suốt"]
        idx = themes.index(self.current_theme)
        self.current_theme = themes[(idx + 1) % len(themes)]
        self.apply_theme(self.current_theme)

    def apply_theme(self, theme):
        self.btn_theme.config(text=f"🎨 Giao diện: {theme}")
        
        if theme == "Sáng":
            bg_main = "#f0f0f0"
            fg_main = "black"
            bg_btn = "#e0e0e0"
            bg_lbl = "white"
            self.root.attributes('-alpha', self.opacity_slider.get())
        elif theme == "Tối":
            bg_main = "#2b2b2b"
            fg_main = "white"
            bg_btn = "#555555"
            bg_lbl = "#444444"
            self.root.attributes('-alpha', self.opacity_slider.get())
        elif theme == "Trong suốt":
            bg_main = "#0a0a0a"
            fg_main = "#00ffcc"
            bg_btn = "#222222"
            bg_lbl = "#1a1a1a"
            self.opacity_slider.set(0.85)
            self.root.attributes('-alpha', 0.85)

        def _update_colors(widget):
            try:
                wtype = widget.winfo_class()
                if wtype in ('Frame', 'Tk', 'Toplevel'):
                    widget.configure(bg=bg_main)
                elif wtype == 'Label':
                    if widget in [self.lbl_file1, self.lbl_file2]:
                        widget.configure(bg=bg_lbl, fg=fg_main)
                    else:
                        c_fg = widget.cget('fg')
                        if c_fg in ['red', 'blue']:
                            widget.configure(bg=bg_main) 
                        else:
                            widget.configure(bg=bg_main, fg=fg_main)
                elif wtype == 'Button' or 'Button' in wtype: # Hỗ trợ luôn ttk.Button
                    c_text = widget.cget('text') if hasattr(widget, 'cget') else ''
                    if c_text in ["Bắt đầu Xử lý", "Xuất dữ liệu ra Excel", "↻ Xóa Lọc / Tải Lại"]:
                        pass 
                    else:
                        widget.configure(bg=bg_btn, fg=fg_main)
            except Exception:
                pass
            for child in widget.winfo_children():
                _update_colors(child)
                
        _update_colors(self.root)
        
        style = ttk.Style()
        style.configure("Treeview", background=bg_lbl, fieldbackground=bg_lbl, foreground=fg_main)
        style.configure("Treeview.Heading", background=bg_btn, foreground=fg_main)
        self.tree.tag_configure('oddrow', background=bg_lbl, foreground=fg_main)
        self.tree.tag_configure('evenrow', background=bg_main, foreground=fg_main)

    def select_file1(self):
        self.file1_path = filedialog.askopenfilename(filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv")])
        if self.file1_path:
            # --- KHẮC PHỤC LỖI CONVERT FILE FAKE XLS CỦA ZKTECO ---
            if self.file1_path.lower().endswith('.xls'):
                try:
                    pd.read_excel(self.file1_path)
                except Exception:
                    # Nếu báo lỗi (bất cứ lỗi gì từ read_excel), chạy bộ chuyển đổi Text -> XLSX
                    converted = False
                    df_temp = None
                    # Bộ quét cường độ cao (Thử tất cả các chuẩn format)
                    encodings = ['utf-8', 'utf-8-sig', 'utf-16', 'utf-16-le', 'cp1258', 'latin1']
                    seps = ['\t', ',', ';'] 
                    
                    for enc in encodings:
                        for sep in seps:
                            try:
                                df_temp = pd.read_csv(self.file1_path, sep=sep, encoding=enc, dtype=str)
                                if len(df_temp.columns) > 3: # Đảm bảo file cắt cột thành công
                                    converted = True
                                    break
                            except: pass
                        if converted: break
                    
                    if converted and df_temp is not None:
                        new_path = self.file1_path + "x" # Chuyển thành đuôi .xlsx
                        df_temp.to_excel(new_path, index=False)
                        self.file1_path = new_path
                        messagebox.showinfo("Tự động sửa lỗi", "Phát hiện File ZKT là định dạng giả mạo (CSV).\nPhần mềm đã tự động chuyển đổi thành file chuẩn (.xlsx) và nạp vào hệ thống!")
                    else:
                        messagebox.showwarning("Cảnh báo", "File ZKTeco quá cũ hoặc bị lỗi nặng, vui lòng mở file lên và chọn Save As thành .xlsx thủ công.")

            self.lbl_file1.config(text=self.file1_path.split('/')[-1])

    def select_file2(self):
        self.file2_path = filedialog.askopenfilename(filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv")])
        if self.file2_path:
            self.lbl_file2.config(text=self.file2_path.split('/')[-1])

    def read_file(self, path):
        encodings = ['utf-8', 'utf-8-sig', 'cp1258', 'latin1', 'utf-16']
        if path.lower().endswith('.csv'):
            for enc in encodings:
                try:
                    return pd.read_csv(path, dtype=str, encoding=enc)
                except:
                    continue
            return pd.read_csv(path, dtype=str, encoding='latin1')
        else:
            try:
                return pd.read_excel(path, dtype=str)
            except ValueError as ve:
                if "Excel file format cannot be determined" in str(ve):
                    for enc in encodings:
                        try:
                            return pd.read_csv(path, dtype=str, sep=None, engine='python', encoding=enc)
                        except:
                            continue
                    for enc in encodings:
                        try:
                            return pd.read_csv(path, dtype=str, sep='\t', engine='python', encoding=enc)
                        except:
                            continue
                    raise ValueError("Không thể đọc được cấu trúc file. Hãy mở file bằng Excel và chọn Save As lại thành định dạng .xlsx chuẩn.")
                else:
                    raise ve

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
                    # Cập nhật danh sách thả xuống
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
        
        # 1. THI HÀNH LỌC DỮ LIỆU
        if selected_dates != ["Tất cả"]:
            df_to_show = df_to_show[df_to_show['Date'].isin(selected_dates)]
        if selected_msnvs != ["Tất cả"]:
            df_to_show = df_to_show[df_to_show['MSNV'].isin(selected_msnvs)]
            
        if selected_timetable != "Tất cả" and selected_timetable != "":
            df_to_show = df_to_show[df_to_show['Timetable'] == selected_timetable]
        if selected_worktime == "Bù vân tay":
            df_to_show = df_to_show[df_to_show['Work Time'] == "Bù vân tay"]
            
        self.display_data(df_to_show)

        # 2. CẬP NHẬT CASCADING OPTIONS (LOẠI BỎ RÁC Ở CÁC DROP DOWNS)
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

            if 'Mã số nhân viên' not in df2.columns or 'Thời gian quẹt thẻ' not in df2.columns:
                messagebox.showerror("Lỗi dữ liệu", "File 2 không đúng chuẩn (Thiếu cột 'Mã số nhân viên' hoặc 'Thời gian quẹt thẻ')!")
                return
                
            df2['Mã số nhân viên'] = df2['Mã số nhân viên'].astype(str).str.strip()
            df2['MSNV_Match'] = df2['Mã số nhân viên'].apply(lambda x: x if x.startswith('S') else 'S' + x.zfill(6))
            df2['Thời gian quẹt thẻ'] = pd.to_datetime(df2['Thời gian quẹt thẻ'], errors='coerce')
            self.df2_clean = df2

            self.df_result = self.df1_clean.copy()
            self.df_result['Kết Quả Công'] = ''
            self.df_result['Nghỉ phép'] = ''
            self.df_result['Tăng ca'] = ''

            for idx in self.df_result.index:
                self.process_single_row(idx, force_recalc=False)

            # Cấp phát list cho dropdown Date & MSNV
            unique_dates = sorted([str(d) for d in self.df_result['Date'].unique() if str(d).strip() != '' and str(d) != 'nan'])
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
        
        if "HC" in tt_upper:
            on_duty, off_duty = "07:30", "17:00"
        elif "CA 1" in tt_upper:
            on_duty, off_duty = "06:00", "14:00"
        elif "CA 2" in tt_upper:
            on_duty, off_duty = "14:00", "22:00"
        elif "CA 3" in tt_upper:
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
            if "CA 3" in tt_upper:
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
            date_val = pd.to_datetime(orig_row['Date']).date()
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
                    
                    is_ca3 = "CA 3" in tt_upper
                    if is_ca3:
                        target_in_dt = datetime.combine(date_val - timedelta(days=1), on_dt_time)
                    else:
                        target_in_dt = datetime.combine(date_val, on_dt_time)
                        
                    target_out_dt = datetime.combine(date_val, off_dt_time)
                    if target_out_dt <= target_in_dt: 
                        target_out_dt += timedelta(days=1)
                        
                    emp_swipes = df2.loc[df2['MSNV_Match'] == msnv, 'Thời gian quẹt thẻ'].dropna()
                    
                    if c_in == '' and c_out != '':
                        valid_ins = emp_swipes[(emp_swipes >= target_in_dt - timedelta(hours=4)) & (emp_swipes <= target_out_dt)]
                        if not valid_ins.empty:
                            new_in = valid_ins.min().strftime("%H:%M")
                            if new_in != c_out: self.df_result.at[idx, 'Clock In'] = new_in
                            
                    elif c_out == '' and c_in != '':
                        valid_outs = emp_swipes[(emp_swipes >= target_in_dt) & (emp_swipes <= target_out_dt + timedelta(hours=8))]
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
        elif "HC" in tt_upper and c_in != '' and c_out != '' and c_in != 'Bù vân tay' and c_out != 'Bù vân tay':
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
                
                if "HC" in tt_upper:
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
                
                if "HC" in tt_upper: phanduoi = ""
                elif "CA 1" in tt_upper: phanduoi = "C1"
                elif "CA 2" in tt_upper: phanduoi = "C2"
                elif "CA 3" in tt_upper: phanduoi = "C3"
                
                cham_cong_val = phangio + phanphut + phanduoi
                
                if self.df_result.at[idx, 'Nghỉ phép'] != "PLP":
                    num_work = h + (0.5 if m >= 30 else 0)
                    if num_work < 8:
                        np_val = 8 - num_work
                        self.df_result.at[idx, 'Nghỉ phép'] = f"{np_val:g}VR"
            except: pass
        
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

    def export_excel(self):
        if self.df_result is None or self.df_result.empty:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
            return
            
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            try:
                self.df_result.to_excel(save_path, index=False)
                
                wb = load_workbook(save_path)
                ws = wb.active
                
                thin_border = Border(left=Side(style='thin'), 
                                     right=Side(style='thin'), 
                                     top=Side(style='thin'), 
                                     bottom=Side(style='thin'))
                
                fill_color = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
                for r in range(1, ws.max_row + 1):
                    for c in range(1, ws.max_column + 1):
                        ws.cell(row=r, column=c).border = thin_border
                
                headers = [cell.value for cell in ws[1]]
                col_indices = {}
                for col_name in ['Late', 'Early', 'OT Time', 'Work Time']:
                    if col_name in headers:
                        col_indices[col_name] = headers.index(col_name) + 1
                
                for r in range(2, ws.max_row + 1):
                    for col_name, c_idx in col_indices.items():
                        cell = ws.cell(row=r, column=c_idx)
                        val = str(cell.value).strip() if cell.value is not None else ""
                        
                        if val != "" and val != "nan" and val != "None":
                            if col_name == 'Work Time':
                                if val != '08:00': cell.fill = fill_color
                            else:
                                cell.fill = fill_color
                
                wb.save(save_path)
                messagebox.showinfo("Thành công", "Xuất file Excel và Kẻ khung thành công!")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu và định dạng file:\n{str(e)}")

def show_login():
    root = tk.Tk()
    root.withdraw()
    login_win = tk.Toplevel(root)
    login_win.title("iamhisu - Đăng nhập")
    login_win.geometry("300x160")
    login_win.resizable(False, False)
    
    try:
        login_win.iconbitmap('fr.ico')
    except Exception:
        pass
    
    login_win.update_idletasks()
    width = 300
    height = 160
    x = (login_win.winfo_screenwidth() // 2) - (width // 2)
    y = (login_win.winfo_screenheight() // 2) - (height // 2)
    login_win.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    login_win.protocol("WM_DELETE_WINDOW", root.destroy)
    tk.Label(login_win, text="ID đăng nhập:", font=("Arial", 10)).pack(pady=(15, 5))
    entry_id = tk.Entry(login_win, font=("Arial", 10), justify="center")
    entry_id.pack()

    tk.Label(login_win, text="Mật khẩu:", font=("Arial", 10)).pack(pady=(5, 5))
    entry_pass = tk.Entry(login_win, show="*", font=("Arial", 10), justify="center")
    entry_pass.pack()

    def check_login(event=None):
        if entry_id.get() == "admin" and entry_pass.get() == "3042":
            login_win.destroy()
            root.deiconify()
            app = TimekeepApp(root)
        else:
            messagebox.showerror("Lỗi", "Sai ID hoặc mật khẩu!", parent=login_win)

    login_win.bind('<Return>', check_login)
    tk.Button(login_win, text="Đăng nhập", command=check_login, bg="lightblue", font=("Arial", 10, "bold")).pack(pady=10)
    root.mainloop()

if __name__ == "__main__":
    show_login()
