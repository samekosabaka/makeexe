import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
import os

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

        self.tree.tag_configure('oddrow', background="white")
        self.tree.tag_configure('evenrow', background="#f0f0f0")

        frame_bottom = tk.Frame(self.root, pady=10)
        frame_bottom.pack(fill=tk.X)
        tk.Button(frame_bottom, text="Xuất dữ liệu ra Excel", command=self.export_excel, bg="lightgreen", font=("Arial", 12, "bold")).pack(pady=5)

    def select_file1(self):
        self.file1_path = filedialog.askopenfilename(filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv")])
        if self.file1_path:
            # Tự động phát hiện và chuyển đổi file .xls giả mạo thành .xlsx
            if self.file1_path.lower().endswith('.xls'):
                try:
                    pd.read_excel(self.file1_path)
                except ValueError as ve:
                    if "Excel file format cannot be determined" in str(ve):
                        try:
                            encodings = ['utf-8', 'utf-8-sig', 'cp1258', 'latin1', 'utf-16']
                            df_temp = None
                            for enc in encodings:
                                try:
                                    df_temp = pd.read_csv(self.file1_path, sep=None, engine='python', encoding=enc, dtype=str)
                                    break
                                except:
                                    continue
                            
                            if df_temp is None:
                                for enc in encodings:
                                    try:
                                        df_temp = pd.read_csv(self.file1_path, sep='\t', engine='python', encoding=enc, dtype=str)
                                        break
                                    except:
                                        continue
                            
                            if df_temp is not None:
                                new_path = self.file1_path + "x"
                                df_temp.to_excel(new_path, index=False)
                                self.file1_path = new_path
                                messagebox.showinfo("Tự động sửa lỗi", "Phát hiện File ZKT là định dạng giả mạo (CSV).\nPhần mềm đã tự động chuyển đổi thành file chuẩn (.xlsx) và nạp vào hệ thống!")
                        except Exception:
                            pass 
            
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
                except UnicodeDecodeError:
                    continue
                except Exception:
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
                        except UnicodeDecodeError:
                            continue
                        except Exception:
                            continue
                    
                    for enc in encodings:
                        try:
                            return pd.read_csv(path, dtype=str, sep='\t', engine='python', encoding=enc)
                        except:
                            continue
                            
                    raise ValueError("Không thể đọc được cấu trúc file. Hãy mở file bằng Excel và chọn Save As lại thành định dạng .xlsx chuẩn.")
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
            
            # Khởi tạo các cột theo đúng thứ tự yêu cầu
            cols_to_keep = ['MSNV', 'Name', 'Date', 'Timetable', 'On duty', 'Off duty', 'Clock In', 'Clock Out', 'Late', 'Early', 'OT Time', 'Work Time']
            for col in cols_to_keep:
                if col not in df1.columns:
                    df1[col] = ''
            
            res_df = df1[cols_to_keep].copy()
            res_df['Kết Quả Công'] = ''
            res_df['Nghỉ phép'] = ''
            res_df['Tăng ca'] = ''
            
            res_df.replace('nan', '', inplace=True)
            res_df.fillna('', inplace=True)

            if 'Mã số nhân viên' not in df2.columns or 'Thời gian quẹt thẻ' not in df2.columns:
                messagebox.showerror("Lỗi dữ liệu", "File 2 không đúng chuẩn (Thiếu cột 'Mã số nhân viên' hoặc 'Thời gian quẹt thẻ')!")
                return
                
            df2['Mã số nhân viên'] = df2['Mã số nhân viên'].astype(str).str.strip()
            df2['MSNV_Match'] = df2['Mã số nhân viên'].apply(lambda x: x if x.startswith('S') else 'S' + x.zfill(6))
            df2['Thời gian quẹt thẻ'] = pd.to_datetime(df2['Thời gian quẹt thẻ'], errors='coerce')

            for idx, row in res_df.iterrows():
                c_in = str(row['Clock In']).strip()
                c_out = str(row['Clock Out']).strip()
                msnv = str(row['MSNV']).strip()
                timetable = str(row['Timetable']).strip()
                
                try:
                    date_val = pd.to_datetime(row['Date']).date()
                except:
                    date_val = None

                if date_val and (c_in == '' or c_out == ''):
                    is_ca3 = "CA 3" in timetable.upper()
                    
                    if c_in == '' and c_out == '':
                        # Nếu từ đầu cả 2 trống thì KHÔNG bù dữ liệu từ file 2
                        pass 
                    
                    elif c_in == '' and c_out != '':
                        if is_ca3:
                            date_in = date_val - timedelta(days=1)
                            mask_in = (df2['MSNV_Match'] == msnv) & (df2['Thời gian quẹt thẻ'].dt.date == date_in)
                            swipes_in = df2.loc[mask_in, 'Thời gian quẹt thẻ'].dropna().sort_values()
                            if not swipes_in.empty:
                                res_df.at[idx, 'Clock In'] = swipes_in.iloc[-1].strftime("%H:%M")
                        else:
                            mask = (df2['MSNV_Match'] == msnv) & (df2['Thời gian quẹt thẻ'].dt.date == date_val)
                            swipes = df2.loc[mask, 'Thời gian quẹt thẻ'].dropna().sort_values()
                            if not swipes.empty:
                                res_df.at[idx, 'Clock In'] = swipes.iloc[0].strftime("%H:%M")
                    
                    elif c_in != '' and c_out == '':
                        mask_out = (df2['MSNV_Match'] == msnv) & (df2['Thời gian quẹt thẻ'].dt.date == date_val)
                        swipes_out = df2.loc[mask_out, 'Thời gian quẹt thẻ'].dropna().sort_values()
                        if not swipes_out.empty:
                            res_df.at[idx, 'Clock Out'] = swipes_out.iloc[-1].strftime("%H:%M")

                c_in = res_df.at[idx, 'Clock In']
                c_out = res_df.at[idx, 'Clock Out']
                
                # ĐIỀN BÙ VÂN TAY VÀ PLP
                if c_in == '' and c_out == '':
                    res_df.at[idx, 'Nghỉ phép'] = "PLP"
                elif c_in == '' and c_out != '':
                    res_df.at[idx, 'Clock In'] = "Bù vân tay"
                    res_df.at[idx, 'Work Time'] = "Bù vân tay"
                elif c_in != '' and c_out == '':
                    res_df.at[idx, 'Clock Out'] = "Bù vân tay"
                    res_df.at[idx, 'Work Time'] = "Bù vân tay"

                w_time = str(res_df.at[idx, 'Work Time']).strip()

                # --- CHỈ TÍNH LẠI NẾU WORK TIME TRỐNG HOẶC LÀ CA HC NỬA BUỔI ---
                needs_recalc = False
                if w_time == '' or w_time == 'Bù vân tay':
                    needs_recalc = True
                elif "HC" in timetable.upper() and c_in != '' and c_out != '' and c_in != 'Bù vân tay' and c_out != 'Bù vân tay':
                    try:
                        fmt = "%H:%M"
                        tin_dt = datetime.strptime(c_in[:5], fmt)
                        tout_dt = datetime.strptime(c_out[:5], fmt)
                        if tout_dt < tin_dt: tout_dt += timedelta(days=1)
                        physical_sec = (tout_dt - tin_dt).total_seconds()
                        
                        # Nếu đi làm <= 5.5 tiếng (đi làm 1 buổi), ZKTeco thường tính sai do trừ lố 1.5h
                        if physical_sec <= 5.5 * 3600:
                            needs_recalc = True
                    except:
                        pass

                # --- THỰC HIỆN TÍNH TOÁN THÔNG MINH ---
                if needs_recalc and c_in != '' and c_out != '' and c_in != 'Bù vân tay' and c_out != 'Bù vân tay':
                    try:
                        fmt = "%H:%M"
                        tin_dt = datetime.strptime(c_in[:5], fmt)
                        tout_dt = datetime.strptime(c_out[:5], fmt)
                        if tout_dt < tin_dt:
                            tout_dt += timedelta(days=1)
                        
                        total_sec = int((tout_dt - tin_dt).total_seconds())
                        
                        # TRỪ THỜI GIAN NGHỈ TRƯA CA HC (11:30 - 13:00) 
                        if "HC" in timetable.upper():
                            lunch_start = tin_dt.replace(hour=11, minute=30, second=0)
                            lunch_end = tin_dt.replace(hour=13, minute=0, second=0)
                            
                            overlap_start = max(tin_dt, lunch_start)
                            overlap_end = min(tout_dt, lunch_end)
                            
                            if overlap_end > overlap_start:
                                total_sec -= int((overlap_end - overlap_start).total_seconds())
                        
                        if total_sec >= 0:
                            if total_sec > 8 * 3600:
                                res_df.at[idx, 'Work Time'] = "08:00"
                                ot_sec = total_sec - 8 * 3600
                                res_df.at[idx, 'OT Time'] = f"{ot_sec // 3600:02d}:{(ot_sec % 3600) // 60:02d}"
                                res_df.at[idx, 'Late'] = ""  
                                res_df.at[idx, 'Early'] = ""
                            else:
                                res_df.at[idx, 'Work Time'] = f"{total_sec // 3600:02d}:{(total_sec % 3600) // 60:02d}"
                                res_df.at[idx, 'OT Time'] = "" 
                                
                                on_duty = str(res_df.at[idx, 'On duty']).strip()
                                off_duty = str(res_df.at[idx, 'Off duty']).strip()
                                
                                late_val = ""
                                early_val = ""
                                
                                if on_duty != '' and off_duty != '':
                                    try:
                                        dt_on = datetime.strptime(on_duty[:5], fmt)
                                        dt_off = datetime.strptime(off_duty[:5], fmt)
                                        dt_in = datetime.strptime(c_in[:5], fmt)
                                        dt_out = datetime.strptime(c_out[:5], fmt)
                                        
                                        if dt_off < dt_on:
                                            dt_off += timedelta(days=1)
                                            
                                        if dt_in < dt_on and (dt_on - dt_in).total_seconds() > 12 * 3600:
                                            dt_in += timedelta(days=1)
                                        elif dt_in > dt_on and (dt_in - dt_on).total_seconds() > 12 * 3600:
                                            dt_in -= timedelta(days=1)
                                            
                                        if dt_out < dt_on and (dt_on - dt_out).total_seconds() > 12 * 3600:
                                            dt_out += timedelta(days=1)
                                            
                                        late_sec = int((dt_in - dt_on).total_seconds())
                                        if late_sec > 0:
                                            late_val = f"{late_sec // 3600:02d}:{(late_sec % 3600) // 60:02d}"
                                            
                                        early_sec = int((dt_off - dt_out).total_seconds())
                                        if early_sec > 0:
                                            early_val = f"{early_sec // 3600:02d}:{(early_sec % 3600) // 60:02d}"
                                    except:
                                        pass
                                        
                                res_df.at[idx, 'Late'] = late_val
                                res_df.at[idx, 'Early'] = early_val
                        else:
                            res_df.at[idx, 'Work Time'] = "00:00"
                    except:
                        pass
                
                # TÍNH TOÁN KẾT QUẢ CÔNG
                w_time = str(res_df.at[idx, 'Work Time']).strip()
                cham_cong_val = ""
                if w_time != '' and w_time != 'Bù vân tay':
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

        df_rows = df.to_numpy().tolist()
        
        for index, row in enumerate(df_rows):
            if index % 2 == 0:
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
                                if val != '08:00':
                                    cell.fill = fill_color
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
