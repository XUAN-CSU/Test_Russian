# -*- coding: utf-8 -*-
import os
import sys
import openpyxl
import requests
import pygame
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from urllib.parse import quote
import threading

def resource_path(relative_path):
    """Get path relative to exe or script folder"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class RussianVocabularyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Russian Vocabulary A1 Level")
        self.root.geometry("800x600")
        
        # Initialize variables
        self.use_proxy = tk.BooleanVar(value=False)
        self.show_english = tk.BooleanVar(value=True)
        self.proxy_url = tk.StringVar(value="http://proxyhost:port")
        self.status_var = tk.StringVar(value="Ready")  # 先初始化 status_var
        
        # Initialize pygame mixer
        pygame.mixer.init()
        
        # Load data
        self.setup_directories()
        self.load_excel_data()
        
        # Create UI
        self.create_widgets()
        
    def setup_directories(self):
        """Setup directories"""
        self.base_dir = resource_path("")
        self.audio_dir = resource_path("audio_files")
        os.makedirs(self.audio_dir, exist_ok=True)
        
    def load_excel_data(self):
        """Load Excel data"""
        excel_path = resource_path("words.xlsx")
        
        if not os.path.exists(excel_path):
            self.create_sample_excel(excel_path)
            
        self.wb = openpyxl.load_workbook(excel_path)
        self.sheet = self.wb.active
        
        # Load words into memory
        self.words = []
        for row in self.sheet.iter_rows(min_row=2):
            russian = row[0].value
            english = row[1].value
            if russian and english:
                self.words.append((russian, english))
                
        print(f"Loaded {len(self.words)} words")
        
    def create_sample_excel(self, path):
        """Create sample Excel file with A1 level Russian words"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Words"

        ws.cell(row=1, column=1, value="Russian")
        ws.cell(row=1, column=2, value="English")

        # A1 Level Russian words
        words = [
            ("привет", "Hello"),
            ("пока", "Goodbye"),
            ("спасибо", "Thank you"),
            ("пожалуйста", "Please"),
            ("да", "Yes"),
            ("нет", "No"),
            ("хорошо", "Good"),
            ("плохо", "Bad"),
            ("извините", "Excuse me"),
            ("я", "I"),
            ("ты", "You"),
            ("он", "He"),
            ("она", "She"),
            ("мы", "We"),
            ("они", "They"),
            ("дом", "House"),
            ("стол", "Table"),
            ("стул", "Chair"),
            ("книга", "Book"),
            ("ручка", "Pen"),
            ("вода", "Water"),
            ("чай", "Tea"),
            ("кофе", "Coffee"),
            ("хлеб", "Bread"),
            ("молоко", "Milk"),
            ("день", "Day"),
            ("ночь", "Night"),
            ("утро", "Morning"),
            ("вечер", "Evening"),
            ("понедельник", "Monday"),
            ("вторник", "Tuesday"),
            ("среда", "Wednesday"),
            ("четверг", "Thursday"),
            ("пятница", "Friday"),
            ("суббота", "Saturday"),
            ("воскресенье", "Sunday"),
        ]

        for i, (ru, en) in enumerate(words, start=2):
            ws.cell(row=i, column=1, value=ru)
            ws.cell(row=i, column=2, value=en)

        wb.save(path)
        messagebox.showinfo("Info", f"Sample Excel file created: {path}")
        
    def get_proxy(self):
        """Get proxy configuration if enabled"""
        if self.use_proxy.get() and self.proxy_url.get().strip():
            return self.proxy_url.get().strip()
        return None
        
    def download_audio(self, text, language):
        """Download audio file"""
        safe_text = "".join(c if c.isalnum() else "_" for c in text)
        filename = os.path.join(self.audio_dir, f"{language}_{safe_text}.mp3")
        
        if os.path.exists(filename):
            return filename
            
        try:
            encoded_text = quote(text)
            if language == "ru":
                url = f"https://translate.google.com/translate_tts?ie=UTF-8&tl=ru&client=tw-ob&q={encoded_text}"
            else:
                url = f"https://translate.google.com/translate_tts?ie=UTF-8&tl=en&client=tw-ob&q={encoded_text}"
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Referer': 'https://translate.google.com/'
            }
            
            # Proxy configuration
            proxy = self.get_proxy()
            proxies = None
            if proxy:
                proxies = {
                    'http': proxy,
                    'https': proxy
                }
            
            self.update_status(f"Downloading {language}: {text}")
            response = requests.get(url, headers=headers, timeout=30, proxies=proxies)
            
            if response.status_code == 200:
                with open(filename, 'wb') as f:
                    f.write(response.content)
                self.update_status(f"Downloaded: {text}")
                return filename
            else:
                self.update_status(f"Download failed: {text} (Status: {response.status_code})")
                
        except requests.exceptions.Timeout:
            self.update_status(f"Timeout while downloading: {text}")
        except requests.exceptions.ProxyError as e:
            self.update_status(f"Proxy error: {e}")
        except Exception as e:
            self.update_status(f"Download failed: {e}")
        
        return None
        
    def play_audio(self, filename):
        """Play audio file"""
        try:
            pygame.mixer.music.load(filename)
            pygame.mixer.music.play()
        except Exception as e:
            messagebox.showerror("Error", f"Play audio failed: {e}")
            
    def play_audio_async(self, filename):
        """Play audio asynchronously"""
        thread = threading.Thread(target=self.play_audio, args=(filename,))
        thread.daemon = True
        thread.start()
        
    def update_status(self, message):
        """Update status message"""
        self.status_var.set(message)
        self.root.update_idletasks()
        print(message)
        
    def on_download_all(self):
        """Download all audio files"""
        def download_thread():
            self.download_btn.config(state="disabled")
            success_count = 0
            total_count = len(self.words) * 2  # Russian and English
            
            for i, (russian, english) in enumerate(self.words):
                self.update_status(f"Downloading {i+1}/{len(self.words)}: {russian}")
                
                # Download Russian audio
                if self.download_audio(russian, "ru"):
                    success_count += 1
                    
                # Download English audio if show_english is enabled
                if self.show_english.get():
                    if self.download_audio(english, "en"):
                        success_count += 1
                        
            self.update_status(f"Download completed: {success_count}/{total_count} files")
            self.download_btn.config(state="normal")
            messagebox.showinfo("Download Complete", f"Downloaded {success_count}/{total_count} audio files")
            
        thread = threading.Thread(target=download_thread)
        thread.daemon = True
        thread.start()
        
    def on_play_russian(self, word):
        """Play Russian word audio"""
        filename = self.download_audio(word, "ru")
        if filename:
            self.play_audio_async(filename)
            self.update_status(f"Playing Russian: {word}")
        else:
            messagebox.showerror("Error", f"Cannot play audio for: {word}")
            
    def on_play_english(self, word):
        """Play English word audio"""
        if not self.show_english.get():
            return
            
        filename = self.download_audio(word, "en")
        if filename:
            self.play_audio_async(filename)
            self.update_status(f"Playing English: {word}")
            
    def refresh_word_list(self):
        """Refresh the word list display"""
        # Clear existing widgets
        for widget in self.word_frame.winfo_children():
            widget.destroy()
            
        # Create word list
        for i, (russian, english) in enumerate(self.words):
            # Russian word button
            ru_btn = ttk.Button(
                self.word_frame, 
                text=russian,
                width=20,
                command=lambda w=russian: self.on_play_russian(w)
            )
            ru_btn.grid(row=i, column=0, padx=5, pady=2, sticky="w")
            
            # English translation (button if show_english is enabled, else label)
            if self.show_english.get():
                en_btn = ttk.Button(
                    self.word_frame, 
                    text=english,
                    width=20,
                    command=lambda w=english: self.on_play_english(w)
                )
                en_btn.grid(row=i, column=1, padx=5, pady=2, sticky="w")
            else:
                en_label = ttk.Label(self.word_frame, text=english, width=20)
                en_label.grid(row=i, column=1, padx=5, pady=2, sticky="w")
                
    def on_show_english_changed(self):
        """Callback when show_english checkbox is changed"""
        self.refresh_word_list()
        
    def create_widgets(self):
        """Create UI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Russian Vocabulary A1 Level", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Proxy settings
        proxy_check = ttk.Checkbutton(settings_frame, text="Use Proxy", 
                                     variable=self.use_proxy)
        proxy_check.grid(row=0, column=0, sticky="w", padx=5)
        
        proxy_entry = ttk.Entry(settings_frame, textvariable=self.proxy_url, width=40)
        proxy_entry.grid(row=0, column=1, padx=5, sticky="w")
        
        # Show English translation
        english_check = ttk.Checkbutton(settings_frame, text="Show English Translation", 
                                       variable=self.show_english,
                                       command=self.on_show_english_changed)
        english_check.grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        
        # Download button
        self.download_btn = ttk.Button(settings_frame, text="Download All Audio", 
                                      command=self.on_download_all)
        self.download_btn.grid(row=2, column=0, columnspan=2, pady=10)
        
        # Word list frame
        list_frame = ttk.LabelFrame(main_frame, text="Vocabulary Words", padding="10")
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # Column headers
        russian_label = ttk.Label(list_frame, text="Russian Words", font=("Arial", 12, "bold"))
        russian_label.grid(row=0, column=0, padx=5, pady=5)
        
        english_label = ttk.Label(list_frame, text="English Translation", font=("Arial", 12, "bold"))
        english_label.grid(row=0, column=1, padx=5, pady=5)
        
        # Word list container with scrollbar
        container = ttk.Frame(list_frame)
        container.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create canvas and scrollbar
        canvas = tk.Canvas(container, height=300)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Word frame inside scrollable frame
        self.word_frame = ttk.Frame(scrollable_frame)
        self.word_frame.pack(fill="both", expand=True)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Status bar
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                relief="sunken", anchor="w")
        status_label.pack(fill="x", padx=5, pady=2)
        
        # Instructions
        instructions = """
Instructions:
1. Check 'Use Proxy' and enter proxy URL if needed
2. Check 'Show English Translation' to see English words
3. Click 'Download All Audio' to download all audio files
4. Click on Russian words to hear pronunciation
5. Click on English words to hear translation (if enabled)
        """
        
        instructions_text = scrolledtext.ScrolledText(main_frame, height=6, width=70)
        instructions_text.grid(row=4, column=0, columnspan=3, pady=10)
        instructions_text.insert("1.0", instructions)
        instructions_text.config(state="disabled")
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        
        # Initial word list display
        self.refresh_word_list()

if __name__ == "__main__":
    root = tk.Tk()
    app = RussianVocabularyApp(root)
    root.mainloop()