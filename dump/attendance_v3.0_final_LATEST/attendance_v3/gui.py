"""
근태 자동 입력 v3.0 - GUI
"""
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import datetime


class AttendanceGUI:
    """GUI 클래스"""
    
    def __init__(self, on_execute, on_retry):
        """
        초기화
        
        Args:
            on_execute: 실행 버튼 콜백
            on_retry: 재입력 버튼 콜백
        """
        self.on_execute = on_execute
        self.on_retry = on_retry
        
        self.root = tk.Tk()
        self.root.title("근태 자동 입력 v3.0")
        self.root.geometry("800x600")
        
        # 파일 경로
        self.raw_file = tk.StringVar()
        self.yeoju_file = tk.StringVar()
        self.smc_file = tk.StringVar()
        self.base_date = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        
        self._create_widgets()
    
    def _create_widgets(self):
        """위젯 생성"""
        # 파일 선택 프레임
        file_frame = tk.LabelFrame(self.root, text="파일 선택", padx=10, pady=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 원시 데이터
        tk.Label(file_frame, text="원시 데이터:").grid(row=0, column=0, sticky=tk.W, pady=2)
        tk.Entry(file_frame, textvariable=self.raw_file, width=50).grid(row=0, column=1, padx=5, pady=2)
        tk.Button(file_frame, text="찾아보기", command=lambda: self._browse_file(self.raw_file, "원시 데이터")).grid(row=0, column=2, pady=2)
        
        # 여주 근태표
        tk.Label(file_frame, text="여주 근태표:").grid(row=1, column=0, sticky=tk.W, pady=2)
        tk.Entry(file_frame, textvariable=self.yeoju_file, width=50).grid(row=1, column=1, padx=5, pady=2)
        tk.Button(file_frame, text="찾아보기", command=lambda: self._browse_file(self.yeoju_file, "여주 근태표")).grid(row=1, column=2, pady=2)
        
        # SMC 근태표
        tk.Label(file_frame, text="SMC 근태표:").grid(row=2, column=0, sticky=tk.W, pady=2)
        tk.Entry(file_frame, textvariable=self.smc_file, width=50).grid(row=2, column=1, padx=5, pady=2)
        tk.Button(file_frame, text="찾아보기", command=lambda: self._browse_file(self.smc_file, "SMC 근태표")).grid(row=2, column=2, pady=2)
        
        # 날짜 입력 프레임
        date_frame = tk.LabelFrame(self.root, text="기준 날짜", padx=10, pady=10)
        date_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(date_frame, text="날짜 (YYYY-MM-DD):").grid(row=0, column=0, sticky=tk.W, pady=2)
        tk.Entry(date_frame, textvariable=self.base_date, width=20).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        # 버튼 프레임
        button_frame = tk.Frame(self.root, padx=10, pady=10)
        button_frame.pack(fill=tk.X)
        
        self.execute_btn = tk.Button(button_frame, text="실행", command=self._on_execute_click, bg="green", fg="white", width=15, height=2)
        self.execute_btn.pack(side=tk.LEFT, padx=5)
        
        self.retry_btn = tk.Button(button_frame, text="재입력", command=self._on_retry_click, bg="blue", fg="white", width=15, height=2, state=tk.DISABLED)
        self.retry_btn.pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="종료", command=self.root.quit, width=15, height=2).pack(side=tk.LEFT, padx=5)
        
        # 로그 프레임
        log_frame = tk.LabelFrame(self.root, text="로그", padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.logbox = scrolledtext.ScrolledText(log_frame, height=20, state=tk.NORMAL, wrap=tk.WORD)
        self.logbox.pack(fill=tk.BOTH, expand=True)
    
    def _browse_file(self, var: tk.StringVar, title: str):
        """파일 찾아보기"""
        filetypes = [
            ("Excel 파일", "*.xlsx;*.xls"),
            ("모든 파일", "*.*")
        ]
        filename = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if filename:
            var.set(filename)
    
    def _on_execute_click(self):
        """실행 버튼 클릭"""
        # 파일 검증
        if not self.raw_file.get():
            messagebox.showerror("오류", "원시 데이터 파일을 선택하세요.")
            return
        if not self.yeoju_file.get():
            messagebox.showerror("오류", "여주 근태표 파일을 선택하세요.")
            return
        if not self.smc_file.get():
            messagebox.showerror("오류", "SMC 근태표 파일을 선택하세요.")
            return
        
        # 날짜 검증
        try:
            datetime.strptime(self.base_date.get(), "%Y-%m-%d")
        except:
            messagebox.showerror("오류", "날짜 형식이 올바르지 않습니다. (YYYY-MM-DD)")
            return
        
        # 로그 초기화
        self.logbox.delete(1.0, tk.END)
        
        # 실행
        self.on_execute(
            raw_file=self.raw_file.get(),
            yeoju_file=self.yeoju_file.get(),
            smc_file=self.smc_file.get(),
            base_date=self.base_date.get()
        )
    
    def _on_retry_click(self):
        """재입력 버튼 클릭"""
        self.on_retry()
    
    def enable_retry_button(self):
        """재입력 버튼 활성화"""
        self.retry_btn.config(state=tk.NORMAL)
    
    def disable_retry_button(self):
        """재입력 버튼 비활성화"""
        self.retry_btn.config(state=tk.DISABLED)
    
    def run(self):
        """GUI 실행"""
        self.root.mainloop()
