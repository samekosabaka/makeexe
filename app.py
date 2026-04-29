import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from datetime import datetime, timedelta
import numpy as np

class TimekeepApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Phần mềm xử lý dữ liệu chấm công")
        self.root.geometry("1000x600")
        
        self.file1_path = ""
        self.file2_path = ""
        self.df_result = None

        # --- Giao diện (UI) ---
        frame_top = tk.Frame(self.root, pady=10)
        frame_top.pack(fill=tk.X)

        # Cụm File 1
        tk.Label(frame_top, text="1. File zkt:").grid(row=0, column=0, padx=10, sticky="e")
        self.lbl_file1 = tk.Label(frame_top, text="Chưa chọn file", width=40, anchor="w", bg="white", relief="sunken")
        self.lbl_file1.grid(row=0, column=1, padx=5)
        tk.Button(frame_top, text="Chọn File", command=self.select_file1).grid(row=0, column=2, padx=5)

        # Cụm File 2
        tk.Label(frame_top, text="2. Dữ liệu HR:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.lbl_file2 = tk.Label(frame_top, text="Chưa chọn file", width=40, anchor="w", bg="white", relief="sunken")
        self.lbl_file2.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(frame_top, text="Chọn File", command=self.select_file2).grid(row=1, column=2, padx=5, pady=5)

        # Nút Xử lý
        tk.Button(frame_top, text="Bắt đầu Xử lý", command=self.process_data, bg="lightblue", font=("Arial", 10, "bold")).grid(row=0, column=3, rowspan=2, padx=20, ipadx=10, ipady=10)

        # Bảng hiển thị dữ liệu
        frame_middle = tk.Frame(self.root)
        frame_middle.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Scrollbar cho bảng
        scroll_y = tk.Scrollbar(frame_middle, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(frame_middle, orient=tk.HORIZONTAL)
        
        self.tree = ttk.Treeview(frame_middle, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.tree.yview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.config(command=self.tree.xview)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Nút Xuất file
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
        if path.endswith('.csv'):
            return pd.read_csv(path, dtype=str)
        else:
            return pd.read_excel(path, dtype=str)

    def process_data(self):
        if not self.file1_path or not self.file2_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn đủ cả 2 file!")
            return

        try:
            df1 = self.read_file(self.file1_path)
            df2 = self.read_file(self.file2_path)

            # 2.1 Lấy dữ liệu file 1
            # 2.1.1 Đổi AC-No thành MSNV
            df1['AC-No.'] = df1['AC-No.'].astype(str).str.strip()
            df1['MSNV'] = 'S' + df1['AC-No.'].str[2:]
            
            # 2.1.2 Chọn các cột kế tiếp
            cols_to_keep = ['MSNV', 'Name', 'Date', 'Timetable', 'On duty', 'Off duty', 'Clock In', 'Clock Out', 'Late', 'Early', 'OT Time', 'Work Time']
            # Đảm bảo các cột tồn tại
            for col in cols_to_keep:
                if col not in df1.columns:
                    df1[col] = ''
            
            res_df = df1[cols_to_keep].copy()
            res_df.replace('nan', '', inplace=True)
            res_df.fillna('', inplace=True)

            # 2.2 Xử lý đối chiếu với file 2
            # Chuẩn hóa dữ liệu file 2
            df2['Mã số nhân viên'] = df2['Mã số nhân viên'].astype(str).str.strip()
            # Xử lý để khớp MSNV (nếu file 2 chưa có S thì thêm S)
            df2['MSNV_Match'] = df2['Mã số nhân viên'].apply(lambda x: x if x.startswith('S') else 'S' + x.zfill(6))
            df2['Thời gian quẹt thẻ'] = pd.to_datetime(df2['Thời gian quẹt thẻ'], errors='coerce')
            
            # Hàm hỗ trợ tính toán thời gian
            def time_diff(t_in, t_out, timetable):
                try:
                    fmt = "%H:%M"
                    tin_dt = datetime.strptime(t_in, fmt)
                    tout_dt = datetime.strptime(t_out, fmt)
                    # Nếu out < in (qua đêm), cộng thêm 1 ngày
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

            # Duyệt qua từng dòng file kết quả
            for idx, row in res_df.iterrows():
                c_in = str(row['Clock In']).strip()
                c_out = str(row['Clock Out']).strip()
                msnv = str(row['MSNV']).strip()
                
                try:
                    date_val = pd.to_datetime(row['Date']).date()
                except:
                    date_val = None

                # Lọc dữ liệu quẹt thẻ của MSNV trong ngày tương ứng ở File 2
                if date_val and (c_in == '' or c_out == ''):
                    mask = (df2['MSNV_Match'] == msnv) & (df2['Thời gian quẹt thẻ'].dt.date == date_val)
                    swipes = df2.loc[mask, 'Thời gian quẹt thẻ'].dropna().sort_values()
                    
                    if not swipes.empty:
                        min_time = swipes.iloc[0].strftime("%H:%M")
                        max_time = swipes.iloc[-1].strftime("%H:%M")
                        
                        # 2.2.1 Nếu cả 2 đều rỗng
                        if c_in == '' and c_out == '':
                            if len(swipes) >= 2:
                                res_df.at[idx, 'Clock In'] = min_time
                                res_df.at[idx, 'Clock Out'] = max_time
                            elif len(swipes) == 1:
                                # Nếu chỉ có 1 lần quẹt, tự động phân bổ gần giờ on duty / off duty nhất (mặc định cho In)
                                res_df.at[idx, 'Clock In'] = min_time
                        
                        # 2.2.2 Nếu 1 trong 2 rỗng
                        elif c_in == '' and c_out != '':
                            res_df.at[idx, 'Clock In'] = min_time
                        elif c_in != '' and c_out == '':
                            # Lấy quẹt thẻ muộn nhất cho clock out
                            res_df.at[idx, 'Clock Out'] = max_time

                # Cập nhật lại biến sau khi điền
                c_in = res_df.at[idx, 'Clock In']
                c_out = res_df.at[idx, 'Clock Out']
                w_time = str(res_df.at[idx, 'Work Time']).strip()
                timetable = str(res_df.at[idx, 'Timetable'])

                # 2.3 Tính Work time nếu trống
                if w_time == '' and c_in != '' and c_out != '':
                    w_time = time_diff(c_in, c_out, timetable)
                    res_df.at[idx, 'Work Time'] = w_time

                # 2.4 Tạo cột mới: Chấm Công (Dựa vào Work Time)
                cham_cong_val = ""
                if w_time != '':
                    try:
                        h_str, m_str = w_time.split(":")
                        h = int(h_str)
                        m = int(m_str)
                        
                        # 2.4.1 Lấy phần giờ
                        phangio = str(h)
                        # 2.4.2 Phần phút
                        phanphut = ".5" if m >= 30 else ""
                        # 2.4.3 Phần đuôi ca làm
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
            self.tree.column(col, width=100)

        df_rows = df.to_numpy().tolist()
        for row in df_rows:
            self.tree.insert("", "end", values=row)

    def export_excel(self):
        if self.df_result is None or self.df_result.empty:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
            return
            
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            try:
                self.df_result.to_excel(save_path, index=False)
                messagebox.showinfo("Thành công", "Xuất file Excel thành công!")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Lỗi khi lưu file:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TimekeepApp(root)
    root.mainloop()