import customtkinter as ctk
import PyPDF2
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, Toplevel
import fitz  # PyMuPDF
import re
import pandas as pd
import smtplib
from email.message import EmailMessage
from tqdm import tqdm
from PIL import Image, ImageTk

def split_pdf(input_pdf_path):
    output_files = []
    current_directory = os.getcwd()
    reader = PyPDF2.PdfReader(input_pdf_path)
    for i, page in enumerate(reader.pages):
        writer = PyPDF2.PdfWriter()
        writer.add_page(page)
        output_file = os.path.join(current_directory, f'{os.path.splitext(os.path.basename(input_pdf_path))[0]}_page_{i + 1}.pdf')
        with open(output_file, 'wb') as f:
            writer.write(f)
        output_files.append(output_file)
        print(f'Saved: {output_file}')
    return output_files

def process_and_rename_pdf(pdf_path, personel):
    pdf_document = fitz.open(pdf_path)
    pdf_text_list = [pdf_document.load_page(page_num).get_text() for page_num in range(len(pdf_document))]
    pdf_document.close()
    numbers_list = [number for text in pdf_text_list for number in re.findall(r'\b\d{4,5}\b', text)]
    if len(numbers_list) >= 2:
        code_number = int(numbers_list[1])
        if code_number in personel['code'].values:
            new_name = personel.loc[personel['code'] == code_number, 'name'].values[0]
            new_pdf_name = f"{new_name}.pdf"
            new_pdf_path = os.path.join(os.path.dirname(pdf_path), new_pdf_name)
            os.rename(pdf_path, new_pdf_path)
            print(f"PDF renamed to {new_pdf_path}")
            return new_pdf_path, code_number
        else:
            return pdf_path, code_number
    print(f"PDF {pdf_path} could not be renamed.")
    return None, None

def send_email_with_attachment(to_email, subject, file_path, person_name, sender_email, sender_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = to_email
    msg.set_content(f"{person_name} عزیز,\n\nفیش حقوقی شما به شرح زیر می‌باشد.\n\nبا احترام")
    with open(file_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(file_path)
    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        print(f"Email sent to {to_email}")

class PDFEmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Splitter and Email Sender")
        self.root.geometry("400x400")

        self.main_frame = ctk.CTkFrame(root)
        self.main_frame.pack(pady=20)

        self.label = ctk.CTkLabel(self.main_frame, text="Select PDF files to process and send:", font=("Helvetica", 14))
        self.label.grid(row=0, column=0, columnspan=2, pady=10)

        self.select_button = ctk.CTkButton(self.main_frame, text="Select PDFs", command=self.select_pdfs, width=200)
        self.select_button.grid(row=1, column=0, columnspan=2, pady=5)

        self.load_excel_button = ctk.CTkButton(self.main_frame, text="Load Excel", command=self.load_excel, width=200)
        self.load_excel_button.grid(row=2, column=0, columnspan=2, pady=5)

        self.email_label = ctk.CTkLabel(self.main_frame, text="Sender Email:", font=("Helvetica", 12))
        self.email_label.grid(row=3, column=0, sticky="e", pady=5)
        self.email_entry = ctk.CTkEntry(self.main_frame, width=200)
        self.email_entry.grid(row=3, column=1, pady=5)

        self.password_label = ctk.CTkLabel(self.main_frame, text="Sender Password:", font=("Helvetica", 12))
        self.password_label.grid(row=4, column=0, sticky="e", pady=5)
        self.password_entry = ctk.CTkEntry(self.main_frame, show="*", width=200)
        self.password_entry.grid(row=4, column=1, pady=5)

        self.send_button = ctk.CTkButton(self.main_frame, text="Process and Send", command=self.process_and_send, width=200)
        self.send_button.grid(row=5, column=0, columnspan=2, pady=10)

        self.status_label = ctk.CTkLabel(self.main_frame, text="", font=("Helvetica", 12), fg_color="green")
        self.status_label.grid(row=6, column=0, columnspan=2, pady=10)

        self.pdf_files = []
        self.personel = pd.DataFrame()

    def select_pdfs(self):
        self.pdf_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_files:
            self.status_label.config(text=f"Selected {len(self.pdf_files)} files.")
        else:
            self.status_label.config(text="No files selected.")

    def load_excel(self):
        excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if excel_path:
            self.personel = pd.read_excel(excel_path)
            self.status_label.config(text=f"Loaded personnel data from {os.path.basename(excel_path)}.")
        else:
            self.status_label.config(text="No Excel file selected.")

    def process_and_send(self):
        if not self.pdf_files:
            messagebox.showwarning("Warning", "Please select PDF files first.")
            return

        if self.personel.empty:
            messagebox.showwarning("Warning", "Please load personnel data from an Excel file.")
            return

        sender_email = self.email_entry.get()
        sender_password = self.password_entry.get()

        if not sender_email or not sender_password:
            messagebox.showwarning("Warning", "لطفا ایمیل و پسورد ایمیل را ارسال کنید.")
            return

        total_files = len(self.pdf_files)
        progress_bar = tqdm(total=total_files, desc="Sending Emails")

        for pdf_path in self.pdf_files:
            split_files = split_pdf(pdf_path)
            for file_path in split_files:
                new_path, code = process_and_rename_pdf(file_path, self.personel)
                if new_path:
                    if code in self.personel['code'].values:
                        person_row = self.personel.loc[self.personel['code'] == code]
                        email_address = person_row['email'].values[0]
                        person_name = person_row['name'].values[0]
                    else:
                        person_name = simpledialog.askstring("Input", f"لطفا نام فرد با کد پرسنلی {code} را وارد کنید:")
                        email_address = simpledialog.askstring("Input", f" ایمیل {person_name} را وارد کنید:")

                    # Show PDF and confirmation before sending email
                    if self.show_pdf_and_confirm(new_path, person_name, email_address):
                        send_email_with_attachment(email_address, f"فیش حقوقی {person_name}", new_path, person_name, sender_email, sender_password)
                        progress_bar.update(1)  # Update progress bar

        progress_bar.close()
        messagebox.showinfo("Success", "All emails sent successfully.")
        self.status_label.config(text="All emails sent successfully.")

        
    def show_pdf_and_confirm(self, pdf_path, person_name, email_address):
        # Create a new window to show the PDF and ask for confirmation
        confirm_window = Toplevel(self.root)
        confirm_window.title("Confirm Email")

        label = ctk.CTkLabel(confirm_window, text=f"آیا می‌خواهید ایمیل را به {person_name} ({email_address}) ارسال کنید؟")
        label.pack(pady=10)

        # Render the first page of the PDF
        pdf_document = fitz.open(pdf_path)
        page = pdf_document.load_page(0)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((400, 400), Image.Resampling.LANCZOS)  # Corrected attribute here
        photo = ImageTk.PhotoImage(img)

        image_label = ctk.CTkLabel(confirm_window, image=photo)
        image_label.image = photo  # Keep a reference to avoid garbage collection
        image_label.pack(pady=10)

        button_frame = ctk.CTkFrame(confirm_window)
        button_frame.pack(pady=10)

        yes_button = ctk.CTkButton(button_frame, text="Yes", command=lambda: self.confirm_action(confirm_window, True))
        yes_button.grid(row=0, column=0, padx=10)

        no_button = ctk.CTkButton(button_frame, text="No", command=lambda: self.confirm_action(confirm_window, False))
        no_button.grid(row=0, column=1, padx=10)

        self.confirm_result = None
        self.root.wait_window(confirm_window)
        return self.confirm_result


    def confirm_action(self, window, result):
        self.confirm_result = result
        window.destroy()

def main():
    root = ctk.CTk()
    app = PDFEmailApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
