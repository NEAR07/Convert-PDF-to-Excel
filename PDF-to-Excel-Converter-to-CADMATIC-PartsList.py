import pdfplumber
import pandas as pd
from tkinter import Tk, Label, Button, filedialog, messagebox, Toplevel
from tkinter import font as tkfont


class PDFToExcelConverter:
    def __init__(self, master):
        self.master = master
        master.title("Convert PDF to Excel (PARTLIST CADMATIC)")
        master.geometry("500x300")
        master.configure(bg="#2E4053")  # Warna biru navy


        self.font = tkfont.Font(family="Arial", size=12)


        self.label = Label(master, text="Convert PDF to Excel", font=self.font, bg="#2E4053", fg="#FFFFFF")  # Warna teks putih
        self.label.pack(pady=20)


        self.convert_button = Button(master, text="Select PDF File", command=self.select_pdf_file, font=self.font, bg="#3498DB", fg="#FFFFFF")  # Warna biru muda
        self.convert_button.pack(pady=10)


        self.status_label = Label(master, text="", font=self.font, bg="#2E4053", fg="#FFFFFF")  # Warna teks putih
        self.status_label.pack(pady=10)


    def select_pdf_file(self):
        pdf_path = filedialog.askopenfilename(title="Pilih file PDF", filetypes=[("PDF files", "*.pdf")])


        if not pdf_path:
            self.status_label.config(text="Tidak ada file PDF yang dipilih.", fg="#FF0000")  # Warna merah
            return


        try:
            output_path = self.save_excel_file()
            if output_path:
                self.extract_data(pdf_path, output_path)
                self.show_success_message(output_path)
        except Exception as e:
            self.status_label.config(text=str(e), fg="#FF0000")  # Warna merah


    def save_excel_file(self):
        return filedialog.asksaveasfilename(title="Simpan File Excel", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])


    def extract_data(self, pdf_path, output_path):
        data_list = []
        profile_number = None
        block_number = None


        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split("\n")


                for line in lines:
                    if "Block :" in line:
                        block_number = line.split(":")[1].strip().replace("|", "")
                    if "PROFILE :" in line:
                        profile_number = line.split(":")[1].strip().replace("|", "")


                for line in lines:
                    if line.startswith("*"):
                        parts = [p.strip() for p in line.split("|")]
                        if len(parts) == 10:
                            parts.insert(0, "")
                        if len(parts) == 11:
                            parts[1] = parts[1].lstrip("*")
                            row = [block_number, profile_number] + parts[1:11]
                            data_list.append(row)


        if not block_number or not profile_number:
            raise ValueError("Block number or PROFILE number not found in the PDF.")


        columns = ["Block", "Profile", "BCode-P-Sub-Part", "Job", "Mat", "Profile Type", "Thi", "Length", "Qty", "Pos", "Grade", "Description"]
        df = pd.DataFrame(data_list, columns=columns)
        df.to_excel(output_path, index=False, sheet_name="Part List", engine="xlsxwriter")


    def show_success_message(self, output_path):
        success_window = Toplevel(self.master)
        success_window.title("Proses Convert Berhasil")
        success_window.configure(bg="#2E4053")  # Warna biru navy


        success_label = Label(success_window, text="Data berhasil disimpan ke :", font=self.font, bg="#2E4053", fg="#FFFFFF")  # Warna teks putih
        success_label.pack(pady=10)


        success_path_label = Label(success_window, text=output_path, font=self.font, bg="#2E4053", fg="#FFFFFF", wraplength=400)  # Warna teks putih
        success_path_label.pack(pady=10)


        success_button = Button(success_window, text="Tutup", command=success_window.destroy, font=self.font, bg="#3498DB", fg="#FFFFFF")  # Warna biru muda
        success_button.pack(pady=10)


        success_window.update_idletasks()
        success_window.geometry(f"{success_window.winfo_reqwidth()}x{success_window.winfo_reqheight()}")




if __name__ == "__main__":
    root = Tk()
    app = PDFToExcelConverter(root)
    root.mainloop()



