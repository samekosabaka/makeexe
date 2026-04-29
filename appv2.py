import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

class TimekeepApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phần mềm xử lý dữ liệu chấm công")
        self.root.geometry("1000x600")
        
        # Căn giữa cửa sổ chính
        self.root.update_idletasks()
        width = 1000
        height = 600
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        self.file1_path = ""
        self.file2_path = ""
        self.df_result = None

        # --- Tùy chỉnh giao diện Bảng (Treeview Style) ---
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
                        background="white",
                        foreground="black",
                        rowheight=25, 
                        fieldbackground="white")
        style.map('Treeview', background=[('selected', '#347083')])
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'), background="#d3d3d3")

        # --- Giao diện (UI) ---
        frame_top = tk.Frame(self.root, pady=10)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="1. File zkt:").grid(row=0, column=0, padx=10, sticky="e")
        self.lbl_file1 = tk.Label(frame_top, text="Chưa chọn file", width=40, anchor="w", bg="white", relief="sunken")
        self.lbl_file1.grid(row=0, column=1, padx=5)
        tk.Button(frame_top, text="Chọn File", command=self.select_file1).grid(row=0, column=2, padx=5)

        tk.Label(frame_top, text="2. Dữ liệu HR:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.lbl_file2 = tk.Label(frame_top, text="Chưa chọn file", width=40, anchor="w", bg="white", relief="sunken")
        self.lbl_file2.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(frame_top, text="Chọn File", command=self.select_file2).grid(row=1, column=2, padx=5, pady=5)

        tk.Button(frame_top, text="Bắt đầu Xử lý", command=self.process_data, bg="lightblue", font=("Arial", 10, "bold")).grid(row=0, column=3, rowspan=2, padx=20, ipadx=10, ipady=10)

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

        # Cài đặt màu sắc cho bảng
        self.tree.tag_configure('oddrow', background="white")
        self.tree.tag_configure('evenrow', background="#f0f0f0") # Màu xám nhạt
        self.tree.tag_configure('anomaly', background="#ffcccc") # Màu đỏ nhạt cho dòng có vấn đề

        frame_bottom = tk.Frame(self.root, pady=10)
        frame_bottom.pack(fill=tk.X)
        tk.Button(frame_bottom, text="Xuất dữ liệu ra Excel", command=self.export_excel, bg="lightgreen", font=("Arial", 12, "bold")).pack(pady=5)

    def select_file1(self):
        self.file1_path = filedialog.askopenfilename(filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv")])
        if self.file1_path:
            self.lbl_file1.config(text=self.file1_path.split('/')[-1])

    def select_file2(self):
        self.file2_path = filedialog.askopenfilename(filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv")])
        if self.file2_path:
            self.lbl_file2.config(text=self.file2_path.split('/')[-1])

    def read_file(self, path):
        try:
            if path.lower().endswith('.csv'):
                return pd.read_csv(path, dtype=str)
            else:
                return pd.read_excel(path, dtype=str)
        except ValueError as ve:
            if "Excel file format cannot be determined" in str(ve):
                try:
                    return pd.read_csv(path, dtype=str, sep=None, engine='python')
                except:
                    return pd.read_csv(path, dtype=str, sep='\t', engine='python')
            else:
                raise ve

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
                if col not in df1.columns:
                    df1[col] = ''
            
            res_df = df1[cols_to_keep].copy()
            res_df.replace('nan', '', inplace=True)
            res_df.fillna('', inplace=True)

            if 'Mã số nhân viên' not in df2.columns or 'Thời gian quẹt thẻ' not in df2.columns:
                messagebox.showerror("Lỗi dữ liệu", "File 2 không đúng chuẩn (Thiếu cột 'Mã số nhân viên' hoặc 'Thời gian quẹt thẻ')!")
                return
                
            df2['Mã số nhân viên'] = df2['Mã số nhân viên'].astype(str).str.strip()
            df2['MSNV_Match'] = df2['Mã số nhân viên'].apply(lambda x: x if x.startswith('S') else 'S' + x.zfill(6))
            df2['Thời gian quẹt thẻ'] = pd.to_datetime(df2['Thời gian quẹt thẻ'], errors='coerce')
            
            def time_diff(t_in, t_out, timetable):
                try:
                    fmt = "%H:%M"
                    tin_dt = datetime.strptime(t_in[:5], fmt)
                    tout_dt = datetime.strptime(t_out[:5], fmt)
                    if tout_dt < tin_dt:
                        tout_dt += timedelta(days=1)
                    
                    diff = tout_dt - tin_dt
                    if "HC" in str(timetable).upper():
                        diff -= timedelta(hours=1, minutes=30)
                    
                    total_seconds = int(diff.total_seconds())
                    if total_seconds < 0: 
                        return ""
                    hours = total_seconds // 3600
                    minutes = (total_seconds % 3600) // 60
                    return f"{hours:02d}:{minutes:02d}"
                except:
                    return ""

            for idx, row in res_df.iterrows():
                c_in = str(row['Clock In']).strip()
                c_out = str(row['Clock Out']).strip()
                msnv = str(row['MSNV']).strip()
                
                try:
                    date_val = pd.to_datetime(row['Date']).date()
                except:
                    date_val = None

                if date_val and (c_in == '' or c_out == ''):
                    mask = (df2['MSNV_Match'] == msnv) & (df2['Thời gian quẹt thẻ'].dt.date == date_val)
                    swipes = df2.loc[mask, 'Thời gian quẹt thẻ'].dropna().sort_values()
                    
                    if not swipes.empty:
                        min_time = swipes.iloc[0].strftime("%H:%M")
                        max_time = swipes.iloc[-1].strftime("%H:%M")
                        
                        if c_in == '' and c_out == '':
                            if len(swipes) >= 2:
                                res_df.at[idx, 'Clock In'] = min_time
                                res_df.at[idx, 'Clock Out'] = max_time
                            elif len(swipes) == 1:
                                res_df.at[idx, 'Clock In'] = min_time
                        
                        elif c_in == '' and c_out != '':
                            res_df.at[idx, 'Clock In'] = min_time
                        elif c_in != '' and c_out == '':
                            res_df.at[idx, 'Clock Out'] = max_time

                c_in = res_df.at[idx, 'Clock In']
                c_out = res_df.at[idx, 'Clock Out']
                w_time = str(res_df.at[idx, 'Work Time']).strip()
                timetable = str(res_df.at[idx, 'Timetable'])

                if w_time == '' and c_in != '' and c_out != '':
                    w_time = time_diff(c_in, c_out, timetable)
                    res_df.at[idx, 'Work Time'] = w_time

                cham_cong_val = ""
                if w_time != '':
                    try:
                        h_str, m_str = w_time.split(":")
                        h = int(h_str)
                        m = int(m_str)
                        
                        phangio = str(h)
                        phanphut = ".5" if m >= 30 else ""
                        phanduoi = ""
                        tt_upper = timetable.upper()
                        
                        if "HC" in tt_upper:
                            phanduoi = ""
                        elif "CA 1" in tt_upper:
                            phanduoi = "C1"
                        elif "CA 2" in tt_upper:
                            phanduoi = "C2"
                        elif "CA 3" in tt_upper:
                            phanduoi = "C3"
                        
                        cham_cong_val = phangio + phanphut + phanduoi
                    except:
                        pass
                
                res_df.at[idx, 'Kết Quả Công'] = cham_cong_val

            self.df_result = res_df
            self.display_data(self.df_result)
            messagebox.showinfo("Thành công", "Xử lý dữ liệu hoàn tất!")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi trong quá trình xử lý:\n{str(e)}")

    def display_data(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["column"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=90, anchor="center")

        # Lấy chỉ mục các cột để kiểm tra điều kiện tô màu
        cols = list(df.columns)
        try:
            idx_late = cols.index('Late')
            idx_early = cols.index('Early')
            idx_ot = cols.index('OT Time')
            idx_work = cols.index('Work Time')
        except ValueError:
            idx_late = idx_early = idx_ot = idx_work = -1

        df_rows = df.to_numpy().tolist()
        
        for index, row in enumerate(df_rows):
            is_anomaly = False
            
            # Kiểm tra xem dòng này có Đi trễ, Về sớm, Tăng ca, hoặc Giờ làm khác 8 tiếng không
            if idx_late != -1:
                v_late = str(row[idx_late]).strip()
                v_early = str(row[idx_early]).strip()
                v_ot = str(row[idx_ot]).strip()
                v_work = str(row[idx_work]).strip()

                if v_late != '' or v_early != '' or v_ot != '' or (v_work != '' and v_work != '08:00'):
                    is_anomaly = True

            # Nếu có bất thường -> Tô nền đỏ nhạt nguyên dòng
            if is_anomaly:
                self.tree.insert("", "end", values=row, tags=('anomaly',))
            elif index % 2 == 0:
                self.tree.insert("", "end", values=row, tags=('evenrow',))
            else:
                self.tree.insert("", "end", values=row, tags=('oddrow',))

    def export_excel(self):
        if self.df_result is None or self.df_result.empty:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
            return
            
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            try:
                # Lưu file excel trước
                self.df_result.to_excel(save_path, index=False)
                
                # Sau đó mở file ra và tiến hành TÔ MÀU THEO Ô
                wb = load_workbook(save_path)
                ws = wb.active
                
                # Màu vàng nổi bật
                fill_color = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
                # Tìm vị trí (cột) của các tiêu đề cần tô
                headers = [cell.value for cell in ws[1]]
                col_indices = {}
                for col_name in ['Late', 'Early', 'OT Time', 'Work Time']:
                    if col_name in headers:
                        col_indices[col_name] = headers.index(col_name) + 1 # openpyxl đếm từ 1
                
                # Duyệt qua các dòng và tô màu các ô thỏa mãn điều kiện
                for r in range(2, ws.max_row + 1):
                    for col_name, c_idx in col_indices.items():
                        cell = ws.cell(row=r, column=c_idx)
                        val = str(cell.value).strip() if cell.value is not None else ""
                        
                        if val != "":
                            if col_name == 'Work Time':
                                if val != '08:00':
                                    cell.fill = fill_color
                            else:
                                # Các cột Late, Early, OT Time cứ có giá trị là tô màu
                                cell.fill = fill_color
                
                # Lưu lại file đã tô màu
                wb.save(save_path)

                messagebox.showinfo("Thành công", "Xuất file Excel thành công!")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu và tô màu file:\n{str(e)}")

def show_login():
    root = tk.Tk()
    root.withdraw()

    login_win = tk.Toplevel(root)
    login_win.title("Đăng nhập hệ thống")
    login_win.geometry("300x160")
    login_win.resizable(False, False)
    
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
