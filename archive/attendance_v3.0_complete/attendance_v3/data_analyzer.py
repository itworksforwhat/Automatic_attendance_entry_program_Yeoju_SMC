"""
근태 자동 입력 v3.0 - 데이터 분석기
스마트하게 공휴일 감지, 데이터 검증 등을 수행
"""
import pandas as pd
from datetime import datetime, date, timedelta
from typing import Dict, List, Tuple
from models import WorkPattern, AttendanceRecord, ProblemData, ValidationResult
from config import COL_DATE, COL_NAME, COL_IN_RAW, COL_OUT_RAW, HOLIDAY_THRESHOLD, MIN_ATTENDANCE


class DataAnalyzer:
    """데이터 분석기"""
    
    def __init__(self, logger):
        """
        초기화
        
        Args:
            logger: 로거 인스턴스
        """
        self.logger = logger
    
    def _map_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        컬럼명 자동 감지 및 매핑
        
        Args:
            df: 원본 데이터프레임
            
        Returns:
            컬럼명이 매핑된 데이터프레임
        """
        self.logger.debug("컬럼명 자동 감지 중...")
        
        # 원본 컬럼 출력
        self.logger.debug(f"원본 컬럼: {list(df.columns)}")
        
        # 컬럼명 정리 (작은따옴표 제거)
        df.columns = [str(col).strip().strip("'\"") for col in df.columns]
        self.logger.debug(f"정리된 컬럼: {list(df.columns)}")
        
        # 컬럼명 매핑 규칙 (우선순위 기반)
        column_mapping = {}
        found_date = False
        found_name = False
        found_in = False
        found_out = False
        
        for col in df.columns:
            col_lower = str(col).lower().strip()
            
            # 날짜 컬럼 (중복 방지)
            if not found_date and ('근무일자' in col_lower or 'date' in col_lower and '일자' in col_lower):
                column_mapping[col] = COL_DATE
                found_date = True
                self.logger.debug(f"  날짜 컬럼: '{col}' → '{COL_DATE}'")
            
            # 이름 컬럼 (중복 방지)
            elif not found_name and ('이름' in col_lower or 'name' in col_lower) and '성명' not in col_lower:
                column_mapping[col] = COL_NAME
                found_name = True
                self.logger.debug(f"  이름 컬럼: '{col}' → '{COL_NAME}'")
            
            # 출근 컬럼 (정확한 매칭, 중복 방지)
            elif not found_in and '출근시간' in col_lower:
                column_mapping[col] = COL_IN_RAW
                found_in = True
                self.logger.debug(f"  출근 컬럼: '{col}' → '{COL_IN_RAW}'")
            
            # 퇴근 컬럼 (정확한 매칭, 중복 방지)
            elif not found_out and '퇴근시간' in col_lower:
                column_mapping[col] = COL_OUT_RAW
                found_out = True
                self.logger.debug(f"  퇴근 컬럼: '{col}' → '{COL_OUT_RAW}'")
        
        # 컬럼 매핑 적용
        if column_mapping:
            df = df.rename(columns=column_mapping)
            self.logger.success(f"컬럼 매핑 완료: {len(column_mapping)}개")
        
        # 중복 컬럼 제거
        df = self._remove_duplicate_columns(df)
        
        # 필수 컬럼 확인
        missing = []
        for required in [COL_DATE, COL_NAME, COL_IN_RAW, COL_OUT_RAW]:
            if required not in df.columns:
                missing.append(required)
        
        if missing:
            self.logger.error(f"필수 컬럼 누락: {missing}")
            self.logger.error(f"현재 컬럼: {list(df.columns)}")
            raise ValueError(f"필수 컬럼이 없습니다: {missing}")
        
        return df
    
    def _remove_duplicate_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        중복 컬럼 제거 (첫 번째만 유지)
        
        Args:
            df: 데이터프레임
            
        Returns:
            중복이 제거된 데이터프레임
        """
        # 중복 컬럼 찾기
        seen = set()
        cols_to_keep = []
        
        for col in df.columns:
            if col not in seen:
                cols_to_keep.append(col)
                seen.add(col)
            else:
                self.logger.warning(f"중복 컬럼 제거: '{col}'")
        
        return df[cols_to_keep]
    
    def analyze_work_pattern(self, df: pd.DataFrame) -> WorkPattern:
        """
        근무 패턴 분석
        
        Args:
            df: 원시 데이터
            
        Returns:
            WorkPattern: 분석 결과
        """
        self.logger.info("근무 패턴 분석 중...")
        
        # 컬럼명 자동 감지 및 매핑
        df = self._map_columns(df)
        
        # 날짜별 출근 인원 계산
        daily_stats = self._calculate_daily_stats(df)
        
        # 평균 출근 인원 계산 (0이 아닌 날만)
        non_zero_days = [count for count in daily_stats.values() if count > 0]
        avg_attendance = sum(non_zero_days) / len(non_zero_days) if non_zero_days else 0
        
        # 임계값 설정
        threshold = avg_attendance * HOLIDAY_THRESHOLD
        
        # 분류
        work_days = []
        holidays = []
        weekends = []
        
        for day, count in daily_stats.items():
            if count == 0:
                weekends.append(day)
            elif count < threshold or count < MIN_ATTENDANCE:
                holidays.append(day)
            else:
                work_days.append(day)
        
        # 로그
        self.logger.success(f"패턴 분석 완료:")
        self.logger.info(f"  평균 출근 인원: {avg_attendance:.1f}명")
        self.logger.info(f"  공휴일 임계값: {threshold:.1f}명 ({int(HOLIDAY_THRESHOLD*100)}%)")
        self.logger.info(f"  근무일: {len(work_days)}일")
        self.logger.info(f"  공휴일: {len(holidays)}일")
        self.logger.info(f"  주말: {len(weekends)}일")
        
        if holidays:
            self.logger.info("감지된 공휴일:")
            for holiday in sorted(holidays):
                count = daily_stats[holiday]
                self.logger.info(f"  - {holiday.strftime('%Y-%m-%d')} ({count}명 출근)")
        
        return WorkPattern(
            work_days=work_days,
            holidays=holidays,
            weekends=weekends,
            avg_attendance=avg_attendance,
            threshold=threshold
        )
    
    def _calculate_daily_stats(self, df: pd.DataFrame) -> Dict[date, int]:
        """
        날짜별 출근 인원 계산
        
        Args:
            df: 원시 데이터
            
        Returns:
            Dict[date, int]: {날짜: 출근 인원}
        """
        # 날짜 컬럼을 date로 변환
        df = df.copy()
        df[COL_DATE] = pd.to_datetime(df[COL_DATE]).dt.date
        
        # 날짜별 출근 인원 (출근시간이 NaN이 아닌 경우)
        daily_series = df.groupby(COL_DATE)[COL_IN_RAW].apply(
            lambda x: x.notna().sum()
        )
        
        # Series를 dict로 변환
        daily_count = {}
        for date_val, count_val in daily_series.items():
            # count_val이 dict인 경우 처리
            if isinstance(count_val, dict):
                # dict의 값들 중 최대값 사용
                count_val = max(count_val.values()) if count_val else 0
            daily_count[date_val] = int(count_val)
        
        return daily_count
    
    def find_previous_workday(self, target_date: date, pattern: WorkPattern, max_days: int = 7) -> date:
        """
        이전 근무일 찾기
        
        Args:
            target_date: 기준 날짜
            pattern: 근무 패턴
            max_days: 최대 검색 일수
            
        Returns:
            date: 이전 근무일 (없으면 None)
        """
        for i in range(1, max_days + 1):
            prev_date = target_date - timedelta(days=i)
            if pattern.is_work_day(prev_date):
                return prev_date
        
        return None
    
    def validate_data(self, df: pd.DataFrame, base_date: date) -> ValidationResult:
        """
        데이터 검증
        
        Args:
            df: 원시 데이터
            base_date: 기준 날짜
            
        Returns:
            ValidationResult: 검증 결과
        """
        self.logger.info("데이터 검증 중...")
        
        # 컬럼명 자동 감지 및 매핑
        df = self._map_columns(df)
        
        valid_records = []
        problems = []
        
        # 날짜 컬럼을 date로 변환
        df = df.copy()
        df[COL_DATE] = pd.to_datetime(df[COL_DATE]).dt.date
        
        # 기준 날짜 데이터만 필터
        df_today = df[df[COL_DATE] == base_date].copy()
        
        for idx, row in df_today.iterrows():
            name = str(row[COL_NAME]).strip()
            
            if not name or name == 'nan':
                continue
            
            # 출퇴근 시간 파싱
            cin_parsed, cin_ok = self._parse_time(row[COL_IN_RAW])
            cout_parsed, cout_ok = self._parse_time(row[COL_OUT_RAW])
            
            # 문제 체크
            issue = self._check_issues(cin_parsed, cout_parsed, cin_ok, cout_ok)
            
            if issue:
                # 문제 데이터
                problems.append(ProblemData(
                    name=name,
                    date=base_date,
                    issue=issue,
                    check_in=str(row[COL_IN_RAW]) if pd.notna(row[COL_IN_RAW]) else None,
                    check_out=str(row[COL_OUT_RAW]) if pd.notna(row[COL_OUT_RAW]) else None,
                ))
            else:
                # 정상 데이터
                valid_records.append(AttendanceRecord(
                    name=name,
                    date=base_date,
                    check_in=cin_parsed,
                    check_out=cout_parsed,
                ))
        
        # 로그
        self.logger.success("데이터 검증 완료:")
        self.logger.info(f"  정상: {len(valid_records)}건")
        
        if problems:
            self.logger.warning(f"  문제: {len(problems)}건")
            self.logger.info("문제 데이터 목록:")
            for p in problems[:10]:  # 최대 10개만 표시
                self.logger.warning(f"  - {p.name} ({p.date}): {p.issue}")
            if len(problems) > 10:
                self.logger.info(f"  ... 외 {len(problems) - 10}건")
        else:
            self.logger.info(f"  문제: 0건")
        
        return ValidationResult(
            valid_records=valid_records,
            problems=problems
        )
    
    def _parse_time(self, value) -> Tuple[datetime, bool]:
        """
        시간 파싱
        
        Args:
            value: 시간 값
            
        Returns:
            (파싱된 시간, 성공 여부)
        """
        if pd.isna(value):
            return None, True  # 빈 값은 정상
        
        # 이미 datetime이면 그대로 반환
        if isinstance(value, datetime):
            return value, True
        
        # pandas Timestamp 처리
        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime(), True
        
        # xlrd의 시간 형식 처리 (float: Excel의 시리얼 날짜/시간)
        if isinstance(value, (int, float)):
            try:
                # Excel 시리얼 날짜를 datetime으로 변환
                from datetime import datetime, timedelta
                # Excel의 기준 날짜: 1899-12-30
                base_date = datetime(1899, 12, 30)
                dt = base_date + timedelta(days=value)
                return dt, True
            except:
                pass
        
        # 문자열 파싱 시도
        if isinstance(value, str):
            value = value.strip()
            
            try:
                # "YYYY/MM/DD HH:MM" 형식 (원시 데이터의 실제 형식!)
                if '/' in value and ' ' in value:
                    dt = datetime.strptime(value, "%Y/%m/%d %H:%M")
                    return dt, True
                
                # "YYYY-MM-DD HH:MM" 형식
                if '-' in value and ' ' in value:
                    dt = datetime.strptime(value, "%Y-%m-%d %H:%M")
                    return dt, True
                
                # "HH:MM" 형식 (시간만)
                if ':' in value and '/' not in value and '-' not in value:
                    parts = value.split(':')
                    if len(parts) == 2:
                        hour = int(parts[0])
                        minute = int(parts[1])
                        return datetime(2000, 1, 1, hour, minute), True
                
                # "8시" 형식
                if '시' in value:
                    hour = int(value.replace('시', '').strip())
                    return datetime(2000, 1, 1, hour, 0), False  # 형식 오류
                
                # 숫자만
                try:
                    hour = int(value)
                    if 0 <= hour <= 23:
                        return datetime(2000, 1, 1, hour, 0), False  # 형식 오류
                except:
                    pass
            
            return None, False  # 파싱 실패
            
        except Exception:
            return None, False  # 파싱 실패
    
    def _check_issues(self, cin: datetime, cout: datetime, cin_ok: bool, cout_ok: bool) -> str:
        """
        문제 체크
        
        Returns:
            str: 문제 설명 (없으면 None)
        """
        # 시간 형식 오류
        if cin and not cin_ok:
            return "출근 시간 형식 오류"
        if cout and not cout_ok:
            return "퇴근 시간 형식 오류"
        
        # 출근만 있음
        if cin and not cout:
            return "출근만 있음 (퇴근 누락)"
        
        # 퇴근만 있음
        if not cin and cout:
            return "퇴근만 있음 (출근 누락)"
        
        # 퇴근 < 출근
        if cin and cout:
            # 24시간 이상 차이나면 야간 근무로 간주
            if cout < cin and (cin.hour - cout.hour) < 12:
                return "퇴근이 출근보다 빠름"
        
        return None
    
    def create_maps(self, df: pd.DataFrame, target_date: date) -> Dict[str, AttendanceRecord]:
        """
        날짜별 출퇴근 맵 생성
        
        Args:
            df: 원시 데이터
            target_date: 대상 날짜
            
        Returns:
            Dict[이름, AttendanceRecord]
        """
        # 컬럼명 자동 감지 및 매핑
        df = self._map_columns(df)
        
        # 날짜 컬럼을 date로 변환
        df = df.copy()
        df[COL_DATE] = pd.to_datetime(df[COL_DATE]).dt.date
        
        # 해당 날짜 데이터만 필터
        df_day = df[df[COL_DATE] == target_date].copy()
        
        result = {}
        for idx, row in df_day.iterrows():
            name = str(row[COL_NAME]).strip()
            
            if not name or name == 'nan':
                continue
            
            # 시간 파싱
            cin_parsed, _ = self._parse_time(row[COL_IN_RAW])
            cout_parsed, _ = self._parse_time(row[COL_OUT_RAW])
            
            result[name] = AttendanceRecord(
                name=name,
                date=target_date,
                check_in=cin_parsed,
                check_out=cout_parsed,
            )
        
        return result
