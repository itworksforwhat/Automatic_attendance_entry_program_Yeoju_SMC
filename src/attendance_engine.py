"""
근태 자동 입력 v3.0 - 출퇴근 처리 엔진
"""
from datetime import datetime, date
from typing import Dict, Optional
from models import AttendanceRecord, ProcessResult, WorkPattern
from overtime_calculator import OvertimeCalculator
from employee_manager import EmployeeManager
from constants import SHIFT_BOUNDARY_HOUR, AttendancePattern


class AttendanceEngine:
    """출퇴근 처리 엔진"""

    def __init__(self, pattern: WorkPattern, logger, employee_manager: Optional[EmployeeManager] = None):
        """
        초기화

        Args:
            pattern: 근무 패턴
            logger: 로거
            employee_manager: 직원 관리자 (없으면 기본 생성)
        """
        self.pattern = pattern
        self.logger = logger
        self.overtime_calculator = OvertimeCalculator()
        self.employee_manager = employee_manager if employee_manager else EmployeeManager()

    def _calculate_overtime(
        self,
        name: str,
        target_date: date,
        checkin_time: Optional[datetime],
        checkout_time: Optional[datetime]
    ) -> Optional[OvertimeRecord]:
        """
        잔업 시간 계산 헬퍼

        Args:
            name: 직원 이름
            target_date: 날짜
            checkin_time: 출근 시간
            checkout_time: 퇴근 시간

        Returns:
            OvertimeRecord 또는 None
        """
        # 출퇴근 시간이 모두 있어야 계산 가능
        if not checkin_time or not checkout_time:
            return None

        # 평일만 처리 (공휴일/주말 제외)
        if not self.pattern.is_work_day(target_date):
            return None

        # 직원 유형 조회
        employee_type = self.employee_manager.get_employee_type(name)

        # 근무 유형 판단
        shift_type = self.overtime_calculator.determine_shift_type(checkin_time)

        # 잔업 계산
        overtime = self.overtime_calculator.calculate_overtime(
            name=name,
            target_date=target_date,
            checkout_time=checkout_time,
            employee_type=employee_type,
            shift_type=shift_type
        )

        if overtime:
            self.logger.info(f"    잔업: {overtime.overtime_hours}시간 (석식: {'O' if overtime.meal_provided else 'X'})")

        return overtime

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

        # 디버깅: today가 None이면 상세 로그 출력
        if today is None:
            self.logger.debug(f"    [디버깅] '{name}' (정규화: '{name_normalized}')를 today_map에서 찾지 못함")
            self.logger.debug(f"    [디버깅] today_map 전체 키: {list(today_map.keys())}")

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
            overtime = self._calculate_overtime(name, today.date, cin_today, cout_today)
            return ProcessResult(
                check_in=cin_today.strftime("%H:%M"),
                check_out=cout_today.strftime("%H:%M"),
                base_date=today.date,
                pattern=AttendancePattern.TODAY_COMPLETE,
                overtime=overtime
            )
        
        # 케이스 2: 오늘 출근만 있음 (퇴근 대기 또는 야간)
        if cin_today and not cout_today:
            # 오늘 출근이 주간 (0~12시)
            if cin_today.hour < SHIFT_BOUNDARY_HOUR:
                # 전일 퇴근이 있으면 사용
                if cout_yest:
                    overtime = self._calculate_overtime(name, dout_yest, cin_today, cout_yest)
                    return ProcessResult(
                        check_in=cin_today.strftime("%H:%M"),
                        check_out=cout_yest.strftime("%H:%M"),
                        base_date=dout_yest,
                        pattern=AttendancePattern.TODAY_CHECKIN_WITH_PREV_CHECKOUT,
                        overtime=overtime
                    )
                else:
                    # 전일 퇴근 없음 - 출근만
                    return ProcessResult(
                        check_in=cin_today.strftime("%H:%M"),
                        check_out="",
                        base_date=today.date,
                        pattern=AttendancePattern.TODAY_CHECKIN_ONLY
                    )
            else:
                # 야간 근무 (12시 이후 출근)
                # 전일 퇴근 사용
                if cout_yest:
                    overtime = self._calculate_overtime(name, dout_yest, cin_today, cout_yest)
                    return ProcessResult(
                        check_in=cin_today.strftime("%H:%M"),
                        check_out=cout_yest.strftime("%H:%M"),
                        base_date=dout_yest,
                        pattern=AttendancePattern.NIGHT_SHIFT,
                        overtime=overtime
                    )
                else:
                    return ProcessResult(
                        check_in=cin_today.strftime("%H:%M"),
                        check_out="",
                        base_date=today.date,
                        pattern=AttendancePattern.NIGHT_SHIFT_NO_CHECKOUT
                    )
        
        # 케이스 3: 오늘 퇴근만 있음 (전일 야간 근무)
        if not cin_today and cout_today:
            # 전일 출근 사용
            if cin_yest:
                overtime = self._calculate_overtime(name, yesterday.date, cin_yest, cout_today)
                return ProcessResult(
                    check_in=cin_yest.strftime("%H:%M"),
                    check_out=cout_today.strftime("%H:%M"),
                    base_date=yesterday.date,
                    pattern=AttendancePattern.PREV_NIGHT_SHIFT,
                    overtime=overtime
                )
            else:
                return ProcessResult(
                    check_in="",
                    check_out=cout_today.strftime("%H:%M"),
                    base_date=today.date,
                    pattern=AttendancePattern.CHECKOUT_ONLY
                )
        
        # 케이스 4: 오늘 데이터 없음 - 전일 확인
        if not cin_today and not cout_today:
            # 전일 출근+퇴근 있음
            if cin_yest and cout_yest:
                # 전일이 야간 근무인지 확인 (출근 12시 이후)
                if cin_yest.hour >= SHIFT_BOUNDARY_HOUR:
                    # 야간 근무자 → 출근+퇴근 모두 사용
                    overtime = self._calculate_overtime(name, dout_yest, cin_yest, cout_yest)
                    return ProcessResult(
                        check_in=cin_yest.strftime("%H:%M"),
                        check_out=cout_yest.strftime("%H:%M"),
                        base_date=dout_yest,
                        pattern=AttendancePattern.PREV_NIGHT_SHIFT_COMPLETE,
                        overtime=overtime
                    )
                else:
                    # 주간 근무자 → 미출근 (퇴근만 사용)
                    return ProcessResult(
                        check_in="",
                        check_out=cout_yest.strftime("%H:%M"),
                        base_date=dout_yest,
                        pattern=AttendancePattern.ABSENT_WITH_PREV_CHECKOUT
                    )
            
            # 전일 출근만 있음 → 완전 결근
            if cin_yest and not cout_yest:
                return ProcessResult(
                    check_in="",
                    check_out="",
                    base_date=None,
                    pattern=AttendancePattern.PREV_CHECKIN_ONLY_NO_DATA
                )
            
            # 전일 데이터 없음 → 완전 결근
            return ProcessResult(
                check_in="",
                check_out="",
                base_date=None,
                pattern=AttendancePattern.NO_DATA
            )
        
        # 기타 (도달하지 않아야 함)
        return ProcessResult(
            check_in="",
            check_out="",
            base_date=None,
            pattern=AttendancePattern.UNKNOWN
        )
