"""
근태 자동 입력 v3.0 - 직원 유형 관리
"""
import json
import os
from typing import Dict, Optional
from models import EmployeeType


class EmployeeManager:
    """직원 유형 관리자"""

    def __init__(self, config_file: str = "employee_config.json"):
        """
        초기화

        Args:
            config_file: 직원 설정 파일 경로
        """
        self.config_file = config_file
        self.employee_types: Dict[str, EmployeeType] = {}
        self.load_config()

    def load_config(self):
        """설정 파일 로드"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # JSON에서 로드한 문자열을 EmployeeType으로 변환
                    for name, type_str in data.items():
                        try:
                            self.employee_types[name] = EmployeeType[type_str]
                        except KeyError:
                            # 잘못된 타입은 기본값으로
                            self.employee_types[name] = EmployeeType.NORMAL
            except Exception as e:
                print(f"설정 파일 로드 실패: {e}")
                self.employee_types = {}

    def save_config(self):
        """설정 파일 저장"""
        try:
            # EmployeeType을 문자열로 변환하여 저장
            data = {name: emp_type.name for name, emp_type in self.employee_types.items()}
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"설정 파일 저장 실패: {e}")

    def get_employee_type(self, name: str) -> EmployeeType:
        """
        직원의 유형 조회

        Args:
            name: 직원 이름

        Returns:
            EmployeeType (없으면 기본값: NORMAL)
        """
        # 이름 정규화 (공백 제거, 소문자 변환)
        normalized_name = name.strip().lower()

        # 먼저 정확한 이름으로 찾기
        if name in self.employee_types:
            return self.employee_types[name]

        # 정규화된 이름으로 찾기
        for stored_name, emp_type in self.employee_types.items():
            if stored_name.strip().lower() == normalized_name:
                return emp_type

        # 기본값 반환
        return EmployeeType.NORMAL

    def set_employee_type(self, name: str, employee_type: EmployeeType):
        """
        직원의 유형 설정

        Args:
            name: 직원 이름
            employee_type: 직원 유형
        """
        self.employee_types[name] = employee_type
        self.save_config()

    def set_multiple_employees(self, employee_dict: Dict[str, EmployeeType]):
        """
        여러 직원의 유형을 한번에 설정

        Args:
            employee_dict: {이름: 유형} 딕셔너리
        """
        self.employee_types.update(employee_dict)
        self.save_config()

    def remove_employee(self, name: str):
        """
        직원 삭제

        Args:
            name: 직원 이름
        """
        if name in self.employee_types:
            del self.employee_types[name]
            self.save_config()

    def get_all_employees(self) -> Dict[str, EmployeeType]:
        """
        모든 직원 정보 조회

        Returns:
            {이름: 유형} 딕셔너리
        """
        return self.employee_types.copy()

    def clear_all(self):
        """모든 직원 정보 삭제"""
        self.employee_types = {}
        self.save_config()

    def get_employees_by_type(self, employee_type: EmployeeType) -> list:
        """
        특정 유형의 직원 목록 조회

        Args:
            employee_type: 조회할 직원 유형

        Returns:
            해당 유형의 직원 이름 리스트
        """
        return [name for name, emp_type in self.employee_types.items() if emp_type == employee_type]
