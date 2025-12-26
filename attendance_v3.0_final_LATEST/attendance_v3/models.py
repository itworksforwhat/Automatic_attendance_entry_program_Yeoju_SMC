"""
근태 자동 입력 v3.0 - 데이터 모델
"""
from dataclasses import dataclass
from datetime import datetime, date
from typing import Optional, List


@dataclass
class AttendanceRecord:
    """출퇴근 기록"""
    name: str
    date: date
    check_in: Optional[datetime] = None
    check_out: Optional[datetime] = None
    
    def has_check_in(self) -> bool:
        """출근 시간이 있는지"""
        return self.check_in is not None
    
    def has_check_out(self) -> bool:
        """퇴근 시간이 있는지"""
        return self.check_out is not None
    
    def is_complete(self) -> bool:
        """출퇴근 모두 있는지"""
        return self.has_check_in() and self.has_check_out()
    
    def get_check_in_str(self) -> str:
        """출근 시간 문자열"""
        return self.check_in.strftime("%H:%M") if self.check_in else ""
    
    def get_check_out_str(self) -> str:
        """퇴근 시간 문자열"""
        return self.check_out.strftime("%H:%M") if self.check_out else ""


@dataclass
class WorkPattern:
    """근무 패턴 분석 결과"""
    work_days: List[date]          # 근무일 목록
    holidays: List[date]            # 공휴일 목록
    weekends: List[date]            # 주말 목록
    avg_attendance: float           # 평균 출근 인원
    threshold: float                # 공휴일 판단 임계값
    
    def is_work_day(self, target_date: date) -> bool:
        """근무일인지 확인"""
        return target_date in self.work_days
    
    def is_holiday(self, target_date: date) -> bool:
        """공휴일인지 확인"""
        return target_date in self.holidays
    
    def is_weekend(self, target_date: date) -> bool:
        """주말인지 확인"""
        return target_date in self.weekends


@dataclass
class ProblemData:
    """문제가 있는 데이터"""
    name: str
    date: date
    issue: str                      # 문제 유형
    check_in: Optional[str] = None  # 원본 출근
    check_out: Optional[str] = None # 원본 퇴근
    fixed_check_in: Optional[str] = None   # 수정된 출근
    fixed_check_out: Optional[str] = None  # 수정된 퇴근
    
    def to_dict(self) -> dict:
        """딕셔너리로 변환 (Excel 출력용)"""
        return {
            '이름': self.name,
            '날짜': self.date.strftime('%Y-%m-%d'),
            '문제': self.issue,
            '출근': self.check_in or '',
            '퇴근': self.check_out or '',
            '수정_출근': self.fixed_check_in or '',
            '수정_퇴근': self.fixed_check_out or '',
        }


@dataclass
class ValidationResult:
    """데이터 검증 결과"""
    valid_records: List[AttendanceRecord]  # 정상 데이터
    problems: List[ProblemData]            # 문제 데이터
    
    def has_problems(self) -> bool:
        """문제가 있는지"""
        return len(self.problems) > 0
    
    def get_stats(self) -> dict:
        """통계 정보"""
        return {
            'valid_count': len(self.valid_records),
            'problem_count': len(self.problems),
            'total_count': len(self.valid_records) + len(self.problems),
        }


@dataclass
class ProcessResult:
    """처리 결과"""
    check_in: str              # 결정된 출근 시간
    check_out: str             # 결정된 퇴근 시간
    base_date: Optional[date]  # 기준 날짜
    pattern: str               # 사용된 패턴
    
    def __str__(self):
        date_str = self.base_date.strftime('%Y-%m-%d') if self.base_date else 'N/A'
        return f"출근={self.check_in or '없음'}, 퇴근={self.check_out or '없음'}, 날짜={date_str}, 패턴={self.pattern}"
