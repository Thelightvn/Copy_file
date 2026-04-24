import os
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl


class FileCopierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ứng dụng Trích xuất & Sao chép File")
        self.root.geometry("650x450")
        self.root.resizable(False, False)

        # Biến lưu trữ đường dẫn
        self.excel_path = tk.StringVar()
        self.src_folder = tk.StringVar()
        self.dest_folder = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        # Frame chính
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Dòng 1: Chọn file Excel ---
        tk.Label(main_frame, text="1. File Excel (Danh sách file):", font=("Arial", 10, "bold")).grid(row=0, column=0,
                                                                                                      sticky="w",
                                                                                                      pady=5)
        tk.Entry(main_frame, textvariable=self.excel_path, width=50, state="readonly").grid(row=0, column=1, padx=10,
                                                                                            pady=5)
        tk.Button(main_frame, text="Duyệt...", width=10, command=self.browse_excel).grid(row=0, column=2, pady=5)

        # --- Dòng 2: Chọn thư mục nguồn ---
        tk.Label(main_frame, text="2. Thư mục Nguồn (Chứa file):", font=("Arial", 10, "bold")).grid(row=1, column=0,
                                                                                                    sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.src_folder, width=50, state="readonly").grid(row=1, column=1, padx=10,
                                                                                            pady=5)
        tk.Button(main_frame, text="Duyệt...", width=10, command=self.browse_src).grid(row=1, column=2, pady=5)

        # --- Dòng 3: Chọn thư mục đích ---
        tk.Label(main_frame, text="3. Thư mục Đích (Lưu file):", font=("Arial", 10, "bold")).grid(row=2, column=0,
                                                                                                  sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.dest_folder, width=50, state="readonly").grid(row=2, column=1, padx=10,
                                                                                             pady=5)
        tk.Button(main_frame, text="Duyệt...", width=10, command=self.browse_dest).grid(row=2, column=2, pady=5)

        # --- Log Text Area ---
        tk.Label(main_frame, text="Trạng thái / Nhật ký:", font=("Arial", 10, "bold")).grid(row=3, column=0, sticky="w",
                                                                                            pady=(15, 5))
        self.log_text = tk.Text(main_frame, height=10, width=72, state="disabled", bg="#f4f4f4")
        self.log_text.grid(row=4, column=0, columnspan=3, pady=5)

        # --- Progress Bar ---
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=500, mode="determinate")
        self.progress.grid(row=5, column=0, columnspan=2, sticky="w", pady=15)

        # --- Nút Thực thi ---
        self.btn_start = tk.Button(main_frame, text="Bắt đầu Copy", font=("Arial", 10, "bold"), bg="#4CAF50",
                                   fg="white", width=15, command=self.start_copy_thread)
        self.btn_start.grid(row=5, column=2, pady=15)

    def browse_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filepath:
            self.excel_path.set(filepath)

    def browse_src(self):
        folderpath = filedialog.askdirectory()
        if folderpath:
            self.src_folder.set(folderpath)

    def browse_dest(self):
        folderpath = filedialog.askdirectory()
        if folderpath:
            self.dest_folder.set(folderpath)

    def log(self, message):
        """Hàm ghi log vào text area"""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def start_copy_thread(self):
        # Kiểm tra đầu vào
        if not self.excel_path.get() or not self.src_folder.get() or not self.dest_folder.get():
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn đầy đủ File Excel, Thư mục Nguồn và Thư mục Đích!")
            return

        # Disable nút Start để tránh bấm nhiều lần
        self.btn_start.config(state="disabled")
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)  # Xóa log cũ
        self.log_text.config(state="disabled")

        # Chạy logic copy trong một luồng (thread) riêng để không làm đơ UI
        thread = threading.Thread(target=self.process_copy)
        thread.daemon = True
        thread.start()

    def process_copy(self):
        excel_file = self.excel_path.get()
        src_dir = self.src_folder.get()
        dest_dir = self.dest_folder.get()

        try:
            # 1. Đọc danh sách từ Excel
            self.log("Đang đọc file Excel...")
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            sheet = wb.active

            filenames = []
            for row in range(2, sheet.max_row + 1):
                cell_value = sheet[f"A{row}"].value
                if cell_value:
                    raw_string = str(cell_value).strip()
                    # Tự động cắt bỏ phần đường dẫn, chỉ lấy tên file cuối cùng
                    # (Thay thế \ thành / để đảm bảo chạy đúng trên mọi định dạng đường dẫn)
                    filename_only = raw_string.replace('\\', '/').split('/')[-1]
                    filenames.append(filename_only)

            total_files = len(filenames)
            if total_files == 0:
                self.log("Không tìm thấy tên file nào ở cột A (từ hàng 2).")
                self.reset_ui()
                return

            # 2. Quét và lập bản đồ toàn bộ file trong thư mục nguồn (Bao gồm Subfolder)
            self.log("Đang quét cấu trúc thư mục nguồn (bao gồm cả thư mục con)...")
            self.root.update_idletasks()

            # Dictionary lưu { "tên_file.ext" : "đường_dẫn_tuyệt_đối" }
            file_map = {}
            for root_dir, dirs, files in os.walk(src_dir):
                for file in files:
                    # Nếu có file trùng tên ở các thư mục khác nhau, file quét được đầu tiên sẽ được giữ lại
                    if file not in file_map:
                        file_map[file] = os.path.join(root_dir, file)

            self.log(f"Đã quét xong. Tìm thấy tổng cộng {len(file_map)} file trong thư mục nguồn.")

            # 3. Tiến hành đối chiếu và Copy
            self.log(f"Bắt đầu xử lý {total_files} file từ danh sách Excel...")
            self.progress["maximum"] = total_files
            self.progress["value"] = 0

            success_count = 0
            missing_count = 0

            for filename in filenames:
                dest_file_path = os.path.join(dest_dir, filename)

                # Kiểm tra xem file có trong bản đồ đã quét hay không
                if filename in file_map:
                    src_file_path = file_map[filename]
                    try:
                        shutil.copy2(src_file_path, dest_file_path)
                        self.log(f"[THÀNH CÔNG] Đã copy: {filename}")
                        success_count += 1
                    except Exception as e:
                        self.log(f"[LỖI] Không thể copy {filename}: {str(e)}")
                else:
                    self.log(f"[KHÔNG TÌM THẤY] File không tồn tại ở bất kỳ thư mục con nào: {filename}")
                    missing_count += 1

                self.progress["value"] += 1
                self.root.update_idletasks()

            self.log("-" * 50)
            self.log(f"HOÀN THÀNH! Thành công: {success_count}/{total_files}. Không tìm thấy/Lỗi: {missing_count}.")
            messagebox.showinfo("Hoàn thành",
                                f"Đã xử lý xong!\nThành công: {success_count}\nLỗi/Không thấy: {missing_count}")

        except Exception as e:
            self.log(f"[LỖI NGHIÊM TRỌNG] {str(e)}")
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{str(e)}")

        finally:
            self.reset_ui()

    def reset_ui(self):
        self.btn_start.config(state="normal")
        self.progress["value"] = 0


if __name__ == "__main__":
    root = tk.Tk()
    app = FileCopierApp(root)
    root.mainloop()