"""
근태 자동 입력 v3.0 - 상수 정의
"""

# ==============================
# 시간 관련 상수
# ==============================
SHIFT_BOUNDARY_HOUR = 12  # 주간/야간 근무 판단 기준 시간 (12시)

# ==============================
# 파일명 상수
# ==============================
PROBLEM_DATA_FILE = "문제_데이터_확인.xlsx"
OVERTIME_RECORD_FILE = "잔업_기록.xlsx"
EMPLOYEE_CONFIG_FILE = "employee_config.json"

# ==============================
# Excel 참조 셀 주소
# ==============================
YEOJU_REFERENCE_CELL = "W37"  # 여주 근태표 참조 셀
SMC_REFERENCE_CELL = "T31"    # SMC 근태표 참조 셀

# ==============================
# 패턴 이름 (Enum 대신 상수로 정의)
# ==============================
class AttendancePattern:
    """출퇴근 패턴 상수"""
    TODAY_COMPLETE = "today_complete"
    TODAY_CHECKIN_WITH_PREV_CHECKOUT = "today_checkin_with_prev_checkout"
    TODAY_CHECKIN_ONLY = "today_checkin_only"
    NIGHT_SHIFT = "night_shift"
    NIGHT_SHIFT_NO_CHECKOUT = "night_shift_no_checkout"
    PREV_NIGHT_SHIFT = "prev_night_shift"
    CHECKOUT_ONLY = "checkout_only"
    PREV_NIGHT_SHIFT_COMPLETE = "prev_night_shift_complete"
    ABSENT_WITH_PREV_CHECKOUT = "absent_with_prev_checkout"
    PREV_CHECKIN_ONLY_NO_DATA = "prev_checkin_only_no_data"
    NO_DATA = "no_data"
    UNKNOWN = "unknown"

# ==============================
# 파일 분류 키워드
# ==============================
YEOJU_FILE_KEYWORDS = ("yj", "여주", "yeoju")
SMC_FILE_KEYWORDS = ("smc",)
