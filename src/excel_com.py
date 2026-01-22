"""
근태 자동 입력 v3.0 - Excel COM 핸들러
메모 서식 완벽 보존
"""

from datetime import date
from typing import List, Dict, Tuple, Optional
from config import RESET_DATE, RERODE_DATA_YEOJU, RERODE_DATA_SMC
from constants import YEOJU_FILE_KEYWORDS, SMC_FILE_KEYWORDS, YEOJU_REFERENCE_CELL, SMC_REFERENCE_CELL
import os
import pandas as pd


class ExcelCOM:
    """Excel COM 핸들러"""

    def __init__(self, file_path: str, logger):
        """
        초기화

        Args:
            file_path: 엑셀 파일 경로
            logger: 로거
        """
        self.file_path = os.path.abspath(file_path)
        self.logger = logger
        self.excel = None
        self.workbook = None
        self.sheet = None

    def __enter__(self):
        """with 문 지원"""
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """with 문 종료 시 자동 close"""
        self.close()

    def open(self):
        """Excel 열기"""
        try:
            import win32com.client
            import pythoncom

            # COM 초기화
            pythoncom.CoInitialize()

            self.logger.debug("Excel COM 초기화 중...")
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False

            self.logger.debug(f"파일 열기: {self.file_path}")
            self.workbook = self.excel.Workbooks.Open(self.file_path)

            return self

        except Exception as e:
            self.logger.error(f"Excel COM 초기화 실패: {str(e)}")
            self.logger.error("pywin32가 설치되어 있는지 확인하세요: pip install pywin32")
            raise

    def prepare_sheet(self, sheet_name: str, clear_ranges: list):
        """
        시트 준비: 복사 + 셀 지우기

        Args:
            sheet_name: 새 시트 이름
            clear_ranges: 지울 범위 리스트

        Returns:
            생성된 시트 이름
        """
        # 1. 시트 복사
        self._copy_last_sheet(sheet_name)

        # 2. 셀 지우기
        self._clear_ranges(clear_ranges)

        # 3. 저장
        self.save()

        return sheet_name

    def _copy_last_sheet(self, new_name: str):
        """마지막 시트 복사"""
        try:
            # 마지막 시트 가져오기
            last_sheet = self.workbook.Worksheets(self.workbook.Worksheets.Count)
            last_name = last_sheet.Name

            self.logger.info(f"시트 복사: '{last_name}' → '{new_name}'")

            # 같은 이름의 시트가 있는지 확인
            try:
                existing = self.workbook.Worksheets(new_name)
                self.logger.info(f"시트 '{new_name}' 이미 존재 - 기존 시트 사용")
                self.sheet = existing
                return
            except Exception:
                # 시트가 존재하지 않으면 새로 생성
                pass

            # 시트 복사
            last_sheet.Copy(Before=None, After=last_sheet)

            # 복사된 시트 (마지막 위치에 생성됨)
            self.sheet = self.workbook.Worksheets(self.workbook.Worksheets.Count)

            # 이름 변경
            try:
                self.sheet.Name = new_name
                self.logger.success(f"시트 복사 완료: '{last_name}' → '{new_name}'")
            except Exception as e:
                self.logger.warning(
                    f"시트 이름 변경 실패: {str(e)}, 기본 이름 사용: {self.sheet.Name}"
                )

        except Exception as e:
            self.logger.error(f"시트 복사 실패: {str(e)}")
            import traceback

            self.logger.error(traceback.format_exc())
            raise

    def _clear_ranges(self, ranges: list):
        """셀 값 지우기"""
        try:
            cleared = 0

            for rng in ranges:
                self.sheet.Range(rng).ClearContents()
                cleared += self.sheet.Range(rng).Cells.Count

            self.logger.info(f"셀 지우기 완료: {cleared}개")

        except Exception as e:
            self.logger.warning(f"셀 지우기 실패: {str(e)}")

    def _determine_file_type(self) -> Tuple[bool, bool]:
        """
        파일명 기반 여주/SMC 구분

        Returns:
            (is_yeoju, is_smc) 튜플
        """
        path_lower = self.file_path.lower()
        is_yeoju = any(keyword in path_lower for keyword in YEOJU_FILE_KEYWORDS)
        is_smc = any(keyword in path_lower for keyword in SMC_FILE_KEYWORDS)
        self.logger.debug(f"is_yeoju={is_yeoju}, is_smc={is_smc}")
        return is_yeoju, is_smc

    def _get_base_date(self, engine) -> date:
        """
        기준일 계산

        Args:
            engine: AttendanceEngine

        Returns:
            기준 날짜
        """
        base_date = getattr(engine, "base_date", None)
        if base_date is None:
            base_date = date.today()
        return base_date

    def _get_previous_sheet_name(self) -> Optional[str]:
        """
        이전 시트 이름 가져오기

        Returns:
            이전 시트 이름 또는 None
        """
        try:
            if self.sheet.Index > 1:
                prev_sheet = self.workbook.Worksheets(self.sheet.Index - 1)
                prev_sheet_name = prev_sheet.Name
                self.logger.debug(f"이전 시트 이름: {prev_sheet_name!r}")
                return prev_sheet_name
            else:
                self.logger.debug("이전 시트 없음 (첫 번째 시트)")
                return None
        except Exception as e:
            self.logger.warning(f"이전 시트 이름 가져오기 실패: {e}")
            return None

    def _write_reset_date(self, reset_date_str: str):
        """
        RESET_DATE 셀에 기준일 쓰기

        Args:
            reset_date_str: 날짜 문자열 (YYYY-MM-DD)
        """
        for addr in RESET_DATE:
            try:
                self.sheet.Range(addr).Value = reset_date_str
                self.logger.debug(f"RESET_DATE: {addr} <- {reset_date_str}")
            except Exception as e:
                self.logger.warning(f"RESET_DATE 입력 실패 ({addr}): {e}")

    def _write_reference_formula(self, is_yeoju: bool, is_smc: bool, prev_sheet_name: Optional[str]):
        """
        RERODE_DATA 수식 설정

        Args:
            is_yeoju: 여주 파일 여부
            is_smc: SMC 파일 여부
            prev_sheet_name: 이전 시트 이름
        """
        if not prev_sheet_name:
            return

        # 여주 전용
        if is_yeoju:
            for addr in RERODE_DATA_YEOJU:
                try:
                    formula = f"='{prev_sheet_name}'!{YEOJU_REFERENCE_CELL}"
                    self.sheet.Range(addr).Formula = formula
                    self.logger.debug(f"RERODE_DATA_YEOJU: {addr} <- {formula}")
                except Exception as e:
                    self.logger.warning(f"RERODE_DATA_YEOJU 수식 입력 실패 ({addr}): {e}")

        # SMC 전용
        if is_smc:
            for addr in RERODE_DATA_SMC:
                try:
                    formula = f"='{prev_sheet_name}'!{SMC_REFERENCE_CELL}"
                    self.sheet.Range(addr).Formula = formula
                    self.logger.debug(f"RERODE_DATA_SMC: {addr} <- {formula}")
                except Exception as e:
                    self.logger.warning(f"RERODE_DATA_SMC 수식 입력 실패 ({addr}): {e}")

    def _process_attendance_blocks(
        self,
        blocks: list,
        today_map: dict,
        yesterday_map: dict,
        engine,
        collect_overtime: bool
    ) -> Tuple[List, int, int]:
        """
        출퇴근 데이터 블록 처리

        Args:
            blocks: 블록 리스트
            today_map: 오늘 맵
            yesterday_map: 전일 맵
            engine: AttendanceEngine
            collect_overtime: 잔업 수집 여부

        Returns:
            (잔업_기록_리스트, 입력_건수, 처리_인원수) 튜플
        """
        overtime_records = []
        filled = 0
        processed = 0

        for block_idx, block_data in enumerate(blocks, 1):
            # 블록 데이터 언패킹
            if len(block_data) == 4:
                name_range, in_range, out_range, overtime_range = block_data
            else:
                name_range, in_range, out_range = block_data
                overtime_range = None

            self.logger.debug(f"블록 {block_idx}/{len(blocks)} 처리: {name_range}")

            # 범위 가져오기
            name_cells = self.sheet.Range(name_range)
            in_cells = self.sheet.Range(in_range)
            out_cells = self.sheet.Range(out_range)
            overtime_cells = self.sheet.Range(overtime_range) if overtime_range else None

            # 각 행 처리
            for i in range(1, name_cells.Rows.Count + 1):
                processed += 1

                name = str(name_cells.Cells(i, 1).Value or "").strip()
                if not name or name == "None":
                    continue

                # 이름 정규화 및 맵에서 찾기
                name_normalized = name.replace(" ", "").lower()
                found_in_today = any(k.replace(" ", "").lower() == name_normalized for k in today_map.keys())
                found_in_yesterday = any(k.replace(" ", "").lower() == name_normalized for k in yesterday_map.keys())

                # 디버깅: 첫 번째 블록의 첫 번째 이름만 상세 출력
                if block_idx == 1 and i == 1:
                    self.logger.debug(f"    [디버깅] Excel에서 읽은 이름: '{name}' (정규화: '{name_normalized}')")
                    self.logger.debug(f"    [디버깅] today_map 키 샘플: {list(today_map.keys())[:3]}")
                    self.logger.debug(f"    [디버깅] found_in_today={found_in_today}, found_in_yesterday={found_in_yesterday}")

                if not found_in_today and not found_in_yesterday:
                    self.logger.warning(f"    '{name}': 원시 데이터에서 찾을 수 없음")
                    continue

                # 출퇴근 시간 결정
                result = engine.decide_times(name, today_map, yesterday_map)

                # 잔업 정보 수집
                if collect_overtime and result.overtime:
                    overtime_records.append(result.overtime)

                # 출퇴근 시간 기록
                if result.check_in:
                    in_cells.Cells(i, 1).Value = "'" + result.check_in
                    filled += 1
                if result.check_out:
                    out_cells.Cells(i, 1).Value = "'" + result.check_out
                    filled += 1

                # 잔업시간 기록
                if overtime_cells and result.overtime:
                    overtime_cells.Cells(i, 1).Value = result.overtime.overtime_hours
                    filled += 1

                # 로그 출력
                if result.check_in or result.check_out:
                    date_str = result.base_date.strftime("%Y-%m-%d") if result.base_date else "N/A"
                    overtime_str = f", 잔업={result.overtime.overtime_hours}시간" if result.overtime else ""
                    self.logger.info(
                        f"  {name}: 출근={result.check_in or '없음'}, "
                        f"퇴근={result.check_out or '없음'}, "
                        f"날짜={date_str}, 패턴={result.pattern}{overtime_str}"
                    )

        return overtime_records, filled, processed

    def write_attendance(self, blocks: list, today_map: dict, yesterday_map: dict, engine, collect_overtime=True):
        """
        출퇴근 데이터 입력 (조율자 메서드)

        Args:
            blocks: [(이름범위, 출근범위, 퇴근범위, 잔업범위), ...]
            today_map: 오늘 맵
            yesterday_map: 전일 맵
            engine: AttendanceEngine
            collect_overtime: 잔업 데이터 수집 여부

        Returns:
            잔업 기록 리스트 (collect_overtime=True인 경우)
        """
        try:
            self.logger.info("출퇴근 데이터 입력 중...")

            # 1) 파일 유형 판단
            is_yeoju, is_smc = self._determine_file_type()

            # 2) 기준일 계산
            base_date = self._get_base_date(engine)
            reset_date_str = base_date.strftime("%Y-%m-%d")
            self.logger.debug(f"기준일: {reset_date_str}")

            # 3) 이전 시트 이름 가져오기
            prev_sheet_name = self._get_previous_sheet_name()

            # 4) RESET_DATE 셀에 날짜 쓰기
            self._write_reset_date(reset_date_str)

            # 5) RERODE_DATA 수식 설정
            self._write_reference_formula(is_yeoju, is_smc, prev_sheet_name)

            # 6) 출퇴근 데이터 블록 처리
            overtime_records, filled, processed = self._process_attendance_blocks(
                blocks, today_map, yesterday_map, engine, collect_overtime
            )

            # 7) 결과 로깅
            self.logger.separator()
            self.logger.success("출퇴근 데이터 입력 완료")
            self.logger.info(f"  처리: {processed}명")
            self.logger.info(f"  입력: {filled}건")
            if collect_overtime and overtime_records:
                self.logger.info(f"  잔업: {len(overtime_records)}건")
            self.logger.separator()

            return overtime_records if collect_overtime else None

        except Exception as e:
            self.logger.error(f"데이터 입력 실패: {str(e)}")
            raise

    def save(self):
        """저장 (메모 서식 완벽 보존!)"""
        try:
            self.logger.debug("파일 저장 중...")
            self.workbook.Save()
            self.logger.success("파일 저장 완료")
        except Exception as e:
            self.logger.error(f"파일 저장 실패: {str(e)}")
            raise

    def close(self):
        """Excel 종료"""
        try:
            if self.workbook:
                self.logger.debug("워크북 닫기")
                self.workbook.Close(SaveChanges=False)
                self.workbook = None

            if self.excel:
                self.logger.debug("Excel 종료")
                self.excel.Quit()
                self.excel = None

            # COM 정리
            import pythoncom

            pythoncom.CoUninitialize()

        except Exception as e:
            self.logger.debug(f"Excel 종료 중 오류 (무시): {str(e)}")

    @staticmethod
    def save_overtime_records(overtime_records: list, output_path: str, logger):
        """
        잔업 기록을 Excel 파일로 저장

        Args:
            overtime_records: OvertimeRecord 리스트
            output_path: 출력 파일 경로
            logger: 로거
        """
        try:
            if not overtime_records:
                logger.info("저장할 잔업 기록이 없습니다.")
                return

            logger.info(f"잔업 기록 저장 중: {len(overtime_records)}건")

            # OvertimeRecord를 딕셔너리로 변환
            data = [record.to_dict() for record in overtime_records]

            # DataFrame 생성
            df = pd.DataFrame(data)

            # Excel 파일로 저장
            df.to_excel(output_path, index=False, engine='openpyxl')

            logger.success(f"잔업 기록 저장 완료: {output_path}")
            logger.info(f"  총 {len(overtime_records)}건의 잔업 기록")

        except Exception as e:
            logger.error(f"잔업 기록 저장 실패: {str(e)}")
            raise
