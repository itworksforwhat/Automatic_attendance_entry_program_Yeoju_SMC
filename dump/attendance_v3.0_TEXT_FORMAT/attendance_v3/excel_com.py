"""
근태 자동 입력 v3.0 - Excel COM 핸들러
메모 서식 완벽 보존
"""

from datetime import date
from config import RESET_DATE, RERODE_DATA_YEOJU, RERODE_DATA_SMC
import os


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
            except:
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

    def write_attendance(self, blocks: list, today_map: dict, yesterday_map: dict, engine):
        """
        출퇴근 데이터 입력

        Args:
            blocks: [(이름범위, 출근범위, 퇴근범위), ...]
            today_map: 오늘 맵
            yesterday_map: 전일 맵
            engine: AttendanceEngine
        """
        try:
            self.logger.info("출퇴근 데이터 입력 중...")

            # === 0) 파일명 기준 여주 / SMC 구분 ===
            path_lower = self.file_path.lower()
            is_yeoju = ("yj" in path_lower) or ("여주" in path_lower) or ("yeoju" in path_lower)
            is_smc   = ("smc" in path_lower)

            self.logger.debug(f"is_yeoju={is_yeoju}, is_smc={is_smc}")

            # === 1) 기준일 계산 ===
            base_date = getattr(engine, "base_date", None)
            if base_date is None:
                base_date = date.today()

            # RESET_DATE용: YYYY-MM-DD 형식
            reset_date_str = base_date.strftime("%Y-%m-%d")

            # === 1-1) 현재 시트/이전 시트 이름 가져오기 ===
            cur_sheet = self.sheet
            wb = self.workbook
            prev_sheet_name = None

            try:
                if cur_sheet.Index > 1:
                    prev_sheet = wb.Worksheets(cur_sheet.Index - 1)
                    prev_sheet_name = prev_sheet.Name
                    self.logger.debug(f"이전 시트 이름: {prev_sheet_name!r}")
                else:
                    self.logger.debug("이전 시트 없음 (첫 번째 시트)")
            except Exception as e:
                self.logger.warning(f"이전 시트 이름 가져오기 실패: {e}")

            self.logger.debug(
                f"기준일: {reset_date_str}, 이전 시트: {prev_sheet_name}"
            )

            # === 2) RESET_DATE 셀들에 기준일 쓰기 (여주/SMC 공통) ===
            for addr in RESET_DATE:
                try:
                    self.sheet.Range(addr).Value = reset_date_str
                    self.logger.debug(f"RESET_DATE: {addr} <- {reset_date_str}")
                except Exception as e:
                    self.logger.warning(f"RESET_DATE 입력 실패 ({addr}): {e}")

            # === 3) 여주 전용: 이전 시트의 W37 참조 ===
            if is_yeoju and prev_sheet_name:
                for addr in RERODE_DATA_YEOJU:
                    try:
                        ref_addr = "W37"  # 여주에서 참조할 셀 주소
                        formula = f"='{prev_sheet_name}'!{ref_addr}"

                        self.logger.debug(f"prev_sheet_name raw: {repr(prev_sheet_name)}")
                        self.logger.debug(f"formula: {repr(formula)}")

                        cell = self.sheet.Range(addr)
                        cell.Formula = formula
                        self.logger.debug(f"RERODE_DATA_YEOJU: {addr} <- {cell.Formula}")

                    except Exception as e:
                        self.logger.warning(
                            f"RERODE_DATA_YEOJU 수식 입력 실패 ({addr}): {e}"
                        )

            # === 4) SMC 전용: 이전 시트의 T31 참조 ===
            if is_smc and prev_sheet_name:
                for addr in RERODE_DATA_SMC:
                    try:
                        ref_addr = "T31"  # SMC에서 참조할 셀 주소
                        formula = f"='{prev_sheet_name}'!{ref_addr}"

                        self.logger.debug(f"prev_sheet_name raw: {repr(prev_sheet_name)}")
                        self.logger.debug(f"formula: {repr(formula)}")

                        cell = self.sheet.Range(addr)
                        cell.Formula = formula
                        self.logger.debug(f"RERODE_DATA_SMC: {addr} <- {cell.Formula}")

                    except Exception as e:
                        self.logger.warning(
                            f"RERODE_DATA_SMC 수식 입력 실패 ({addr}): {e}"
                        )

            # === 5) 기존 출퇴근 입력 로직 ===
            filled = 0
            processed = 0

            for block_idx, (name_range, in_range, out_range) in enumerate(blocks, 1):
                self.logger.debug(f"블록 {block_idx}/{len(blocks)} 처리: {name_range}")

                # 범위 가져오기
                name_cells = self.sheet.Range(name_range)
                in_cells = self.sheet.Range(in_range)
                out_cells = self.sheet.Range(out_range)

                # 각 행 처리
                for i in range(1, name_cells.Rows.Count + 1):
                    processed += 1

                    name = str(name_cells.Cells(i, 1).Value or "").strip()
                    if not name or name == "None":
                        continue

                    # 디버깅: 이름 출력
                    self.logger.debug(f"  처리 중: '{name}'")

                    # 맵에서 이름 찾기 (대소문자 무시, 공백 제거)
                    name_normalized = name.replace(" ", "").lower()

                    found_in_today = False
                    found_in_yesterday = False

                    for map_name in today_map.keys():
                        if map_name.replace(" ", "").lower() == name_normalized:
                            found_in_today = True
                            self.logger.debug(f"    오늘 맵에서 발견: '{map_name}'")
                            break

                    for map_name in yesterday_map.keys():
                        if map_name.replace(" ", "").lower() == name_normalized:
                            found_in_yesterday = True
                            self.logger.debug(f"    전일 맵에서 발견: '{map_name}'")
                            break

                    if not found_in_today and not found_in_yesterday:
                        self.logger.warning(
                            f"    '{name}': 원시 데이터에서 찾을 수 없음"
                        )
                        continue

                    # 출퇴근 시간 결정
                    result = engine.decide_times(name, today_map, yesterday_map)

                    # 셀에 쓰기 (텍스트 형식으로 강제)
                    if result.check_in:
                        in_cells.Cells(i, 1).Value = "'" + result.check_in
                        filled += 1
                    if result.check_out:
                        out_cells.Cells(i, 1).Value = "'" + result.check_out
                        filled += 1

                    # 로그 (데이터 있을 때만)
                    if result.check_in or result.check_out:
                        date_str = (
                            result.base_date.strftime("%Y-%m-%d")
                            if result.base_date
                            else "N/A"
                        )
                        self.logger.info(
                            f"  {name}: 출근={result.check_in or '없음'}, "
                            f"퇴근={result.check_out or '없음'}, "
                            f"날짜={date_str}, 패턴={result.pattern}"
                        )

            self.logger.separator()
            self.logger.success("출퇴근 데이터 입력 완료")
            self.logger.info(f"  처리: {processed}명")
            self.logger.info(f"  입력: {filled}건")
            self.logger.separator()

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
