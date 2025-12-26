"""
근태 자동 입력 v3.0 - 로깅
"""
import tkinter as tk
from datetime import datetime
from enum import Enum


class LogLevel(Enum):
    """로그 레벨"""
    DEBUG = "DEBUG"
    INFO = "INFO"
    SUCCESS = "SUCCESS"
    WARNING = "WARNING"
    ERROR = "ERROR"


class Logger:
    """로거 클래스"""
    
    # 레벨별 색상
    COLORS = {
        LogLevel.DEBUG: "gray",
        LogLevel.INFO: "black",
        LogLevel.SUCCESS: "green",
        LogLevel.WARNING: "orange",
        LogLevel.ERROR: "red",
    }
    
    def __init__(self, logbox: tk.Text = None):
        """
        초기화
        
        Args:
            logbox: GUI 로그 출력 위젯
        """
        self.logbox = logbox
        self.warning_count = 0
        self.error_count = 0
        
        if self.logbox:
            # 태그 설정
            for level, color in self.COLORS.items():
                self.logbox.tag_config(level.value, foreground=color)
    
    def _log(self, level: LogLevel, message: str):
        """
        로그 출력
        
        Args:
            level: 로그 레벨
            message: 메시지
        """
        # 시간
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # 포맷
        log_line = f"{timestamp} [{level.value}] {message}\n"
        
        # 콘솔 출력
        print(log_line.strip())
        
        # GUI 출력
        if self.logbox:
            self.logbox.insert(tk.END, log_line, level.value)
            self.logbox.see(tk.END)
            self.logbox.update()
        
        # 카운트
        if level == LogLevel.WARNING:
            self.warning_count += 1
        elif level == LogLevel.ERROR:
            self.error_count += 1
    
    def debug(self, message: str):
        """디버그 로그"""
        self._log(LogLevel.DEBUG, message)
    
    def info(self, message: str):
        """정보 로그"""
        self._log(LogLevel.INFO, message)
    
    def success(self, message: str):
        """성공 로그"""
        self._log(LogLevel.SUCCESS, message)
    
    def warning(self, message: str):
        """경고 로그"""
        self._log(LogLevel.WARNING, message)
    
    def error(self, message: str):
        """오류 로그"""
        self._log(LogLevel.ERROR, message)
    
    def separator(self, char: str = "-", length: int = 60):
        """구분선"""
        self.info(char * length)
    
    def section(self, title: str):
        """섹션 헤더"""
        self.separator("=")
        self.info(title)
        self.separator("=")
    
    def has_warnings(self) -> bool:
        """경고가 있는지"""
        return self.warning_count > 0
    
    def has_errors(self) -> bool:
        """오류가 있는지"""
        return self.error_count > 0
    
    def get_warning_count(self) -> int:
        """경고 개수"""
        return self.warning_count
    
    def get_error_count(self) -> int:
        """오류 개수"""
        return self.error_count
    
    def reset_counts(self):
        """카운트 초기화"""
        self.warning_count = 0
        self.error_count = 0
