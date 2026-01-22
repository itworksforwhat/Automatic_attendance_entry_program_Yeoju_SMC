"""
근태 자동 입력 v3.0 - 메인
"""
import pandas as pd
from datetime import datetime
from tkinter import messagebox
import os

from config import *
from logger import Logger
from gui import AttendanceGUI
from data_analyzer import DataAnalyzer
from attendance_engine import AttendanceEngine
from excel_com import ExcelCOM
from models import ProblemData


class AttendanceProcessor:
    """근태 처리 메인 클래스"""
    
    def __init__(self):
        """초기화"""
        self.logger = None
        self.gui = None
        self.problem_file = "문제_데이터_확인.xlsx"
        self.current_files = {}  # 현재 처리 중인 파일 정보
    
    def run(self):
        """실행"""
        # GUI 생성
        self.gui = AttendanceGUI(
            on_execute=self._execute,
            on_retry=self._retry
        )
        
        # GUI 실행
        self.gui.run()
    
    def _execute(self, raw_file: str, yeoju_file: str, smc_file: str, base_date: str):
        """
        메인 실행 로직
        
        Args:
            raw_file: 원시 데이터 파일
            yeoju_file: 여주 근태표 파일
            smc_file: SMC 근태표 파일
            base_date: 기준 날짜 (YYYY-MM-DD)
        """
        try:
            # 로거 초기화
            self.logger = Logger(self.gui.logbox)
            self.logger.section("근태 자동 입력 v3.0 시작")
            
            # 파일 정보 저장
            self.current_files = {
                'yeoju': yeoju_file,
                'smc': smc_file,
                'base_date': base_date
            }
            
            # 날짜 파싱
            base_date_obj = datetime.strptime(base_date, "%Y-%m-%d").date()
            self.logger.info(f"기준 날짜: {base_date}")
            
            # ========== 1단계: 원시 데이터 로드 ==========
            self.logger.separator()
            self.logger.info("1단계: 원시 데이터 로드")
            
            df = self._load_raw_data(raw_file)
            
            # ========== 2단계: 데이터 분석 ==========
            self.logger.separator()
            self.logger.info("2단계: 데이터 분석")
            
            analyzer = DataAnalyzer(self.logger)
            pattern = analyzer.analyze_work_pattern(df)
            
            # 이전 근무일 찾기
            prev_workday = analyzer.find_previous_workday(base_date_obj, pattern)
            if prev_workday:
                self.logger.info(f"이전 근무일: {prev_workday.strftime('%Y-%m-%d')}")
            else:
                self.logger.warning("이전 근무일을 찾을 수 없습니다")
                prev_workday = base_date_obj  # fallback
            
            # ========== 3단계: 데이터 검증 ==========
            self.logger.separator()
            self.logger.info("3단계: 데이터 검증")
            
            validation = analyzer.validate_data(df, base_date_obj)
            
            # ========== 4단계: 맵 생성 ==========
            self.logger.separator()
            self.logger.info("4단계: 출퇴근 맵 생성")
            
            today_map = analyzer.create_maps(df, base_date_obj)
            yesterday_map = analyzer.create_maps(df, prev_workday)
            
            self.logger.info(f"오늘 맵: {len(today_map)}명")
            self.logger.info(f"전일 맵: {len(yesterday_map)}명")
            
            # ========== 5단계: 정상 데이터 입력 ==========
            self.logger.separator()
            self.logger.info("5단계: 정상 데이터 입력")
            
            engine = AttendanceEngine(pattern, self.logger)
            
            # 여주 근태표
            self._process_file(
                "여주",
                yeoju_file,
                YEOJU_BLOCKS,
                CLEAR_RANGES_YEOJU,
                today_map,
                yesterday_map,
                engine,
                base_date_obj
            )
            
            # SMC 근태표
            self._process_file(
                "SMC",
                smc_file,
                SMC_BLOCKS,
                CLEAR_RANGES_SMC,
                today_map,
                yesterday_map,
                engine,
                base_date_obj
            )
            
            # ========== 6단계: 문제 데이터 처리 ==========
            self.logger.separator()
            self.logger.info("6단계: 문제 데이터 처리")
            
            if validation.has_problems():
                self._save_problem_data(validation.problems)
                
                self.logger.warning(f"문제 데이터 {len(validation.problems)}건 발견")
                self.logger.info(f"파일 생성: {self.problem_file}")
                self.logger.info("파일을 확인하고 수정한 후 '재입력' 버튼을 누르세요")
                
                # 재입력 버튼 활성화
                self.gui.enable_retry_button()
                
                messagebox.showwarning(
                    "문제 데이터 발견",
                    f"문제 데이터 {len(validation.problems)}건이 발견되었습니다.\n\n"
                    f"{self.problem_file} 파일을 확인하고 수정한 후\n"
                    "'재입력' 버튼을 눌러주세요."
                )
            else:
                self.logger.success("문제 데이터 없음")
            
            # ========== 완료 ==========
            self.logger.separator("=")
            self.logger.success("✓ 작업 완료")
            self.logger.separator("=")
            
            if not validation.has_problems():
                messagebox.showinfo("완료", "근태 업데이트가 완료되었습니다.")
            
        except Exception as e:
            self.logger.error(f"오류 발생: {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{str(e)}")
    
    def _load_raw_data(self, file_path: str) -> pd.DataFrame:
        """
        원시 데이터 로드 (.xls 직접 지원)
        
        Args:
            file_path: 파일 경로
            
        Returns:
            DataFrame
        """
        self.logger.info(f"파일 로드: {file_path}")
        
        try:
            # .xls 파일은 xlrd 사용
            if file_path.lower().endswith('.xls'):
                self.logger.info(".xls 파일 감지 - xlrd 사용")
                df = pd.read_excel(file_path, engine='xlrd')
            else:
                # .xlsx는 openpyxl 사용
                df = pd.read_excel(file_path, engine='openpyxl')
            
            self.logger.success(f"파일 로드 완료: {len(df)}행")
            return df
            
        except ImportError as e:
            if 'xlrd' in str(e):
                self.logger.error("xlrd 패키지가 설치되지 않았습니다")
                self.logger.error("설치 방법: pip install xlrd")
                raise Exception("xlrd 패키지가 필요합니다")
            raise
        except Exception as e:
            self.logger.error(f"파일 로드 실패: {str(e)}")
            raise
    
    def _process_file(
        self,
        name: str,
        file_path: str,
        blocks: list,
        clear_ranges: list,
        today_map: dict,
        yesterday_map: dict,
        engine: AttendanceEngine,
        base_date
    ):
        """
        근태표 파일 처리
        
        Args:
            name: 파일 이름 (로그용)
            file_path: 파일 경로
            blocks: 블록 리스트
            clear_ranges: 지울 범위
            today_map: 오늘 맵
            yesterday_map: 전일 맵
            engine: 엔진
            base_date: 기준 날짜
        """
        self.logger.separator()
        self.logger.info(f"[{name} 근태표 처리]")
        self.logger.info(f"파일: {file_path}")
        
        # 시트 이름 생성
        sheet_name = base_date.strftime(SHEET_NAME_FORMAT)
        
        try:
            with ExcelCOM(file_path, self.logger) as excel:
                # 시트 준비
                excel.prepare_sheet(sheet_name, clear_ranges)
                
                # 데이터 입력
                excel.write_attendance(blocks, today_map, yesterday_map, engine)
                
                # 저장 (이미 prepare_sheet에서 저장되었지만 한 번 더)
                excel.save()
            
            self.logger.success(f"{name} 근태표 처리 완료")
            
        except Exception as e:
            self.logger.error(f"{name} 근태표 처리 실패: {str(e)}")
            raise
    
    def _save_problem_data(self, problems: list):
        """
        문제 데이터를 Excel 파일로 저장
        
        Args:
            problems: ProblemData 리스트
        """
        try:
            # DataFrame 생성
            data = [p.to_dict() for p in problems]
            df = pd.DataFrame(data)
            
            # Excel 저장
            df.to_excel(self.problem_file, index=False, engine='openpyxl')
            
            self.logger.info(f"문제 데이터 파일 생성: {self.problem_file}")
            
        except Exception as e:
            self.logger.error(f"문제 데이터 파일 생성 실패: {str(e)}")
    
    def _retry(self):
        """재입력 (사용자가 문제 데이터 수정 후)"""
        try:
            self.logger.separator("=")
            self.logger.info("재입력 시작")
            self.logger.separator("=")
            
            # 문제 데이터 파일 확인
            if not os.path.exists(self.problem_file):
                messagebox.showerror("오류", f"{self.problem_file} 파일이 없습니다.")
                return
            
            # 수정된 데이터 로드
            self.logger.info(f"수정된 데이터 로드: {self.problem_file}")
            df_fixed = pd.read_excel(self.problem_file, engine='openpyxl')
            
            # 검증
            if '수정_출근' not in df_fixed.columns or '수정_퇴근' not in df_fixed.columns:
                messagebox.showerror("오류", "파일 형식이 올바르지 않습니다.")
                return
            
            # 시트 이름
            base_date = datetime.strptime(self.current_files['base_date'], "%Y-%m-%d").date()
            sheet_name = base_date.strftime(SHEET_NAME_FORMAT)
            
            # 여주 근태표 재입력
            self._retry_file("여주", self.current_files['yeoju'], sheet_name, YEOJU_BLOCKS, df_fixed)
            
            # SMC 근태표 재입력
            self._retry_file("SMC", self.current_files['smc'], sheet_name, SMC_BLOCKS, df_fixed)
            
            self.logger.separator("=")
            self.logger.success("✓ 재입력 완료")
            self.logger.separator("=")
            
            # 재입력 버튼 비활성화
            self.gui.disable_retry_button()
            
            messagebox.showinfo("완료", "문제 데이터 재입력이 완료되었습니다.")
            
        except Exception as e:
            self.logger.error(f"재입력 실패: {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
            messagebox.showerror("오류", f"재입력 중 오류가 발생했습니다:\n{str(e)}")
    
    def _retry_file(self, name: str, file_path: str, sheet_name: str, blocks: list, df_fixed: pd.DataFrame):
        """
        파일 재입력
        
        Args:
            name: 파일 이름
            file_path: 파일 경로
            sheet_name: 시트 이름
            blocks: 블록 리스트
            df_fixed: 수정된 데이터
        """
        self.logger.info(f"[{name} 근태표 재입력]")
        
        try:
            with ExcelCOM(file_path, self.logger) as excel:
                # 시트 선택
                excel.sheet = excel.workbook.Worksheets(sheet_name)
                
                filled = 0
                
                # 각 행 처리
                for idx, row in df_fixed.iterrows():
                    name_val = str(row['이름']).strip()
                    cin = str(row['수정_출근']).strip() if pd.notna(row['수정_출근']) else ''
                    cout = str(row['수정_퇴근']).strip() if pd.notna(row['수정_퇴근']) else ''
                    
                    if not cin and not cout:
                        continue
                    
                    # 이름 찾아서 입력
                    found = False
                    for name_range, in_range, out_range in blocks:
                        name_cells = excel.sheet.Range(name_range)
                        in_cells = excel.sheet.Range(in_range)
                        out_cells = excel.sheet.Range(out_range)
                        
                        for i in range(1, name_cells.Rows.Count + 1):
                            cell_name = str(name_cells.Cells(i, 1).Value or "").strip()
                            
                            if cell_name == name_val:
                                if cin:
                                    in_cells.Cells(i, 1).Value = cin
                                    filled += 1
                                if cout:
                                    out_cells.Cells(i, 1).Value = cout
                                    filled += 1
                                
                                self.logger.info(f"  {name_val}: 출근={cin or '없음'}, 퇴근={cout or '없음'}")
                                found = True
                                break
                        
                        if found:
                            break
                    
                    if not found:
                        self.logger.warning(f"  {name_val}: 이름을 찾을 수 없음")
                
                # 저장
                excel.save()
                
                self.logger.success(f"{name} 재입력 완료: {filled}건")
                
        except Exception as e:
            self.logger.error(f"{name} 재입력 실패: {str(e)}")
            raise


if __name__ == "__main__":
    processor = AttendanceProcessor()
    processor.run()
