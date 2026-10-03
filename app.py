import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox
import os
from docx2pdf import convert
from pdf2docx import Converter

class LojikApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Lojik Okuyucu v2 - Lordum Tugay")
        self.geometry("600x450")
        
        # Arka plan rengini karanlık temaya uygun ayarla
        self.configure(bg="#242424")
        
        # Ana Çerçeve
        self.frame = ctk.CTkFrame(self, fg_color="#242424")
        self.frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.label_title = ctk.CTkLabel(self.frame, text="Lojik Okuyucu ve Dönüştürücü v2", font=ctk.CTkFont(size=22, weight="bold"))
        self.label_title.pack(pady=5)
        
        # Özel Karşılama Mesajı
        self.greeting_label = ctk.CTkLabel(self.frame, text="Hoşgeldin Lordum Tugay", font=ctk.CTkFont(size=18, slant="italic"), text_color="#F39C12")
        self.greeting_label.pack(pady=10)
        
        # Sürükle Bırak Alanı
        self.drop_area = ctk.CTkLabel(
            self.frame, 
            text="📄\nDönüştürülecek Dosyayı Buraya Sürükle\n(Word, PDF veya PowerPoint)", 
            width=450, height=200, 
            fg_color="#333333", corner_radius=15, font=ctk.CTkFont(size=16)
        )
        self.drop_area.pack(pady=15)
        
        # Sürükle bırak olaylarını bağlama
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.handle_drop)
        
        # Durum Çubuğu
        self.status_label = ctk.CTkLabel(self.frame, text="Sistem Hazır Lordum.", text_color="gray", font=ctk.CTkFont(size=14))
        self.status_label.pack(side="bottom", pady=10)
        
    def handle_drop(self, event):
        # Dosya yolunu temizle
        file_path = event.data.strip('{}') 
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext in [".doc", ".docx"]:
            self.word_to_pdf(file_path)
        elif ext == ".pdf":
            # İleride buraya Word/PPTX seçme butonu mantığı eklenecek
            self.pdf_to_word(file_path)
        elif ext == ".pptx":
            self.pptx_to_pdf(file_path)
        else:
            messagebox.showwarning("Desteklenmeyen Format", "Lütfen şimdilik sadece Word, PDF veya PowerPoint (PPTX) dosyası sürükleyin Lordum.")
            
    def word_to_pdf(self, file_path):
        self.status_label.configure(text="Word -> PDF dönüştürülüyor, lütfen bekleyin...")
        self.update()
        try:
            output_path = file_path.rsplit(".", 1)[0] + ".pdf"
            convert(file_path, output_path)
            self.status_label.configure(text="İşlem Başarılı! Dosya aynı klasöre kaydedildi.")
        except Exception as e:
            self.status_label.configure(text="Hata oluştu!")
            messagebox.showerror("Hata", str(e))
            
    def pdf_to_word(self, file_path):
        self.status_label.configure(text="PDF -> Word dönüştürülüyor, lütfen bekleyin...")
        self.update()
        try:
            output_path = file_path.rsplit(".", 1)[0] + ".docx"
            cv = Converter(file_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            self.status_label.configure(text="İşlem Başarılı! Dosya aynı klasöre kaydedildi.")
        except Exception as e:
            self.status_label.configure(text="Hata oluştu!")
            messagebox.showerror("Hata", str(e))
            
    def pptx_to_pdf(self, file_path):
        self.status_label.configure(text="PowerPoint -> PDF dönüştürülüyor, lütfen bekleyin...")
        self.update()
        try:
            import comtypes.client
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            slides = powerpoint.Presentations.Open(file_path)
            output_path = file_path.rsplit(".", 1)[0] + ".pdf"
            slides.SaveAs(output_path, 32) 
            slides.Close()
            powerpoint.Quit()
            self.status_label.configure(text="İşlem Başarılı! Dosya aynı klasöre kaydedildi.")
        except Exception as e:
            self.status_label.configure(text="Hata oluştu!")
            messagebox.showerror("Hata", f"PowerPoint işlemi sırasında bir sorun çıktı:\n{str(e)}")

if __name__ == "__main__":
    app = LojikApp()
    app.mainloop()