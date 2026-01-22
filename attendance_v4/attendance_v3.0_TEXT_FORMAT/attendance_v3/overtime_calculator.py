"""
근태 자동 입력 v3.0 - 잔업 시간 계산
"""
from datetime import datetime, date, time
from typing import Optional, List, Tuple
from models import OvertimeRecord, EmployeeType
from constants import SHIFT_BOUNDARY_HOUR
import config


class OvertimeCalculator:
    """잔업 시간 계산기"""

    def __init__(self):
        """초기화"""
        # 정규 근무 시간 파싱
        self.day_shift_end = self._parse_time(config.DAY_SHIFT_END)
        self.night_shift_end = self._parse_time(config.NIGHT_SHIFT_END)

    def _parse_time(self, time_str: str) -> time:
        """시간 문자열을 time 객체로 변환"""
        return datetime.strptime(time_str, "%H:%M").time()

    def _time_to_minutes(self, t: time) -> int:
        """time 객체를 분 단위로 변환"""
        return t.hour * 60 + t.minute

    def _get_overtime_table(self, employee_type: EmployeeType) -> List[Tuple[str, float, bool]]:
        """직원 유형에 따른 잔업 테이블 반환"""
        if employee_type == EmployeeType.SMT:
            return config.OVERTIME_TABLE_SMT
        else:
            # 일반 직원과 관리자는 동일한 테이블 사용
            return config.OVERTIME_TABLE_NORMAL

    def _find_overtime_hours(self, checkout_time: time, overtime_table: List[Tuple[str, float, bool]]) -> Tuple[float, bool]:
        """
        퇴근 시간에 맞는 잔업 시간 찾기

        Args:
            checkout_time: 실제 퇴근 시간
            overtime_table: 잔업 테이블

        Returns:
            (잔업시간, 석식제공여부) 튜플
        """
        checkout_minutes = self._time_to_minutes(checkout_time)

        # 이전 시간대 적용 방식
        # 퇴근 시간 이하의 가장 큰 시간대를 찾음
        selected_hours = 0.0
        selected_meal = False

        for time_str, hours, meal in overtime_table:
            table_time = self._parse_time(time_str)
            table_minutes = self._time_to_minutes(table_time)

            # 퇴근 시간이 테이블 시간 이상이면 해당 잔업시간 적용
            if checkout_minutes >= table_minutes:
                selected_hours = hours
                selected_meal = meal
            else:
                # 이미 시간을 초과했으므로 중단
                break

        return selected_hours, selected_meal

    def calculate_overtime(
        self,
        name: str,
        target_date: date,
        checkout_time: datetime,
        employee_type: EmployeeType,
        shift_type: str = "주간"
    ) -> Optional[OvertimeRecord]:
        """
        잔업 시간 계산

        Args:
            name: 직원 이름
            target_date: 날짜
            checkout_time: 퇴근 시간
            employee_type: 직원 유형
            shift_type: 근무 유형 (주간/야간)

        Returns:
            OvertimeRecord 또는 None (잔업이 없는 경우)
        """
        # 퇴근 시간 추출
        checkout_time_only = checkout_time.time()

        # 정규 종료 시간 결정
        if shift_type == "주간":
            regular_end = self.day_shift_end
            regular_end_str = config.DAY_SHIFT_END
        else:  # 야간
            regular_end = self.night_shift_end
            regular_end_str = config.NIGHT_SHIFT_END

        # 잔업이 발생했는지 확인
        # 주간: 17:40 이후 퇴근
        # 야간: 05:40 이후 퇴근
        checkout_minutes = self._time_to_minutes(checkout_time_only)
        regular_minutes = self._time_to_minutes(regular_end)

        # 야간 근무의 경우 다음날 새벽까지이므로 특별 처리
        if shift_type == "야간" and checkout_time_only.hour < SHIFT_BOUNDARY_HOUR:
            # 새벽 시간 (00:00~11:59)은 정규 근무 시간과 비교하기 위해 24시간 추가
            checkout_minutes += 24 * 60
            regular_minutes += 24 * 60 if regular_end.hour < SHIFT_BOUNDARY_HOUR else 0

        # 정규 시간 이전 퇴근이면 잔업 없음
        if checkout_minutes <= regular_minutes:
            return None

        # 잔업 테이블에서 잔업 시간 찾기
        overtime_table = self._get_overtime_table(employee_type)

        # 야간 근무의 경우, 정규 퇴근 이후 시간을 주간 기준으로 변환
        if shift_type == "야간":
            # 정규 퇴근 이후 경과 시간 계산 (분 단위)
            overtime_minutes = checkout_minutes - regular_minutes

            # 주간 정규 퇴근 시간(17:40)에 경과 시간을 더해서 가상의 퇴근 시간 계산
            day_end_minutes = self._time_to_minutes(self.day_shift_end)
            virtual_checkout_minutes = day_end_minutes + overtime_minutes

            # 가상 시간을 time 객체로 변환
            virtual_hours = (virtual_checkout_minutes // 60) % 24
            virtual_mins = virtual_checkout_minutes % 60
            virtual_checkout_time = time(virtual_hours, virtual_mins)

            overtime_hours, meal_provided = self._find_overtime_hours(virtual_checkout_time, overtime_table)
        else:
            overtime_hours, meal_provided = self._find_overtime_hours(checkout_time_only, overtime_table)

        # 잔업 시간이 0이면 잔업 기록 없음
        if overtime_hours == 0:
            return None

        # 잔업 기록 생성
        return OvertimeRecord(
            name=name,
            date=target_date,
            employee_type=employee_type,
            shift_type=shift_type,
            regular_end_time=regular_end_str,
            actual_end_time=checkout_time.strftime("%H:%M"),
            overtime_hours=overtime_hours,
            meal_provided=meal_provided
        )

    def determine_shift_type(self, checkin_time: datetime) -> str:
        """
        출근 시간으로 근무 유형 판단

        Args:
            checkin_time: 출근 시간

        Returns:
            "주간" 또는 "야간"
        """
        # 12시 이후 출근이면 야간 근무
        if checkin_time.hour >= SHIFT_BOUNDARY_HOUR:
            return "야간"
        else:
            return "주간"
