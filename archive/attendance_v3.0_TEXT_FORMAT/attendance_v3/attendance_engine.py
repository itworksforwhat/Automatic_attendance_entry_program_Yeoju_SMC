"""
근태 자동 입력 v3.0 - 출퇴근 처리 엔진
"""
from datetime import datetime, date
from typing import Dict, Tuple
from models import AttendanceRecord, ProcessResult, WorkPattern


class AttendanceEngine:
    """출퇴근 처리 엔진"""
    
    def __init__(self, pattern: WorkPattern, logger):
        """
        초기화
        
        Args:
            pattern: 근무 패턴
            logger: 로거
        """
        self.pattern = pattern
        self.logger = logger
    
    def decide_times(
        self,
        name: str,
        today_map: Dict[str, AttendanceRecord],
        yesterday_map: Dict[str, AttendanceRecord]
    ) -> ProcessResult:
        """
        출퇴근 시간 결정
        
        Args:
            name: 이름
            today_map: 오늘 데이터 맵
            yesterday_map: 전일 데이터 맵
            
        Returns:
            ProcessResult: 처리 결과
        """
        # 이름 정규화 (공백 제거, 소문자 변환)
        name_normalized = name.replace(" ", "").lower()
        
        # 오늘/전일 데이터 가져오기 (정규화된 이름으로 매칭)
        today = None
        yesterday = None
        
        for key, value in today_map.items():
            if key.replace(" ", "").lower() == name_normalized:
                today = value
                break
        
        for key, value in yesterday_map.items():
            if key.replace(" ", "").lower() == name_normalized:
                yesterday = value
                break
        
        # 편의상 변수 추출
        cin_today = today.check_in if today else None
        cout_today = today.check_out if today else None
        cin_yest = yesterday.check_in if yesterday else None
        cout_yest = yesterday.check_out if yesterday else None
        
        # 퇴근 날짜는 퇴근 시간의 날짜를 사용 (야간 근무 고려)
        dout_yest = cout_yest.date() if cout_yest else (yesterday.date if yesterday else None)
        
        # 디버깅
        self.logger.debug(f"    출퇴근 시간: cin_today={cin_today}, cout_today={cout_today}")
        self.logger.debug(f"    전일: cin_yest={cin_yest}, cout_yest={cout_yest}")
        self.logger.debug(f"    퇴근 날짜: dout_yest={dout_yest}")
        
        # 케이스별 처리
        
        # 케이스 1: 오늘 출근+퇴근 모두 있음
        if cin_today and cout_today:
            return ProcessResult(
                check_in=cin_today.strftime("%H:%M"),
                check_out=cout_today.strftime("%H:%M"),
                base_date=today.date,
                pattern="today_complete"
            )
        
        # 케이스 2: 오늘 출근만 있음 (퇴근 대기 또는 야간)
        if cin_today and not cout_today:
            # 오늘 출근이 주간 (0~12시)
            if cin_today.hour < 12:
                # 전일 퇴근이 있으면 사용
                if cout_yest:
                    return ProcessResult(
                        check_in=cin_today.strftime("%H:%M"),
                        check_out=cout_yest.strftime("%H:%M"),
                        base_date=dout_yest,
                        pattern="today_checkin_with_prev_checkout"
                    )
                else:
                    # 전일 퇴근 없음 - 출근만
                    return ProcessResult(
                        check_in=cin_today.strftime("%H:%M"),
                        check_out="",
                        base_date=today.date,
                        pattern="today_checkin_only"
                    )
            else:
                # 야간 근무 (12시 이후 출근)
                # 전일 퇴근 사용
                if cout_yest:
                    return ProcessResult(
                        check_in=cin_today.strftime("%H:%M"),
                        check_out=cout_yest.strftime("%H:%M"),
                        base_date=dout_yest,
                        pattern="night_shift"
                    )
                else:
                    return ProcessResult(
                        check_in=cin_today.strftime("%H:%M"),
                        check_out="",
                        base_date=today.date,
                        pattern="night_shift_no_checkout"
                    )
        
        # 케이스 3: 오늘 퇴근만 있음 (전일 야간 근무)
        if not cin_today and cout_today:
            # 전일 출근 사용
            if cin_yest:
                return ProcessResult(
                    check_in=cin_yest.strftime("%H:%M"),
                    check_out=cout_today.strftime("%H:%M"),
                    base_date=yesterday.date,
                    pattern="prev_night_shift"
                )
            else:
                return ProcessResult(
                    check_in="",
                    check_out=cout_today.strftime("%H:%M"),
                    base_date=today.date,
                    pattern="checkout_only"
                )
        
        # 케이스 4: 오늘 데이터 없음 - 전일 확인
        if not cin_today and not cout_today:
            # 전일 출근+퇴근 있음
            if cin_yest and cout_yest:
                # 전일이 야간 근무인지 확인 (출근 12시 이후)
                if cin_yest.hour >= 12:
                    # 야간 근무자 → 출근+퇴근 모두 사용
                    return ProcessResult(
                        check_in=cin_yest.strftime("%H:%M"),
                        check_out=cout_yest.strftime("%H:%M"),
                        base_date=dout_yest,
                        pattern="prev_night_shift_complete"
                    )
                else:
                    # 주간 근무자 → 미출근 (퇴근만 사용)
                    return ProcessResult(
                        check_in="",
                        check_out=cout_yest.strftime("%H:%M"),
                        base_date=dout_yest,
                        pattern="absent_with_prev_checkout"
                    )
            
            # 전일 출근만 있음 → 완전 결근
            if cin_yest and not cout_yest:
                return ProcessResult(
                    check_in="",
                    check_out="",
                    base_date=None,
                    pattern="prev_checkin_only_no_data"
                )
            
            # 전일 데이터 없음 → 완전 결근
            return ProcessResult(
                check_in="",
                check_out="",
                base_date=None,
                pattern="no_data"
            )
        
        # 기타 (도달하지 않아야 함)
        return ProcessResult(
            check_in="",
            check_out="",
            base_date=None,
            pattern="unknown"
        )
