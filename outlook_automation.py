"""
Outlook 파일 첨부 자동화 스크립트 (Python 3.8+)

기능:
  1. 특정 파일 선택
  2. 출력 폴더 선택
  3. 파일을 기본 앱으로 열기
  4. Outlook에 파일 첨부
  5. 첨부 파일을 output 폴더에 저장

필수 설치:
  pip install pywin32

사용 방법:
  python outlook_automation.py
"""

import os
import sys
import subprocess
from pathlib import Path
from typing import Optional
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback

# COM 라이브러리 임포트
try:
    import win32com.client as win32
    from win32com.client import GetObject
except ImportError:
    print("❌ pywin32가 설치되지 않았습니다.")
    print("다음 명령으로 설치하세요: pip install pywin32")
    sys.exit(1)


class OutlookAutomation:
    """Outlook 자동화 클래스"""

    def __init__(self):
        """초기화"""
        self.outlook = None
        self.mail_item = None
        self.source_file: Optional[str] = None
        self.output_folder: Optional[str] = None

    def select_file(self) -> bool:
        """
        파일 선택 대화상자 표시
        
        Returns:
            bool: 파일이 선택되면 True
        """
        print("\n[Step 1] 파일 선택")
        print("=" * 50)

        # Tkinter 루트 윈도우 (숨김)
        root = tk.Tk()
        root.withdraw()

        file_path = filedialog.askopenfilename(
            title="첨부할 파일을 선택하세요",
            filetypes=[
                ("모든 파일", "*.*"),
                ("Excel", "*.xlsx *.xls"),
                ("Word", "*.docx *.doc"),
                ("PDF", "*.pdf"),
            ],
        )

        root.destroy()

        if not file_path:
            print("❌ 파일이 선택되지 않았습니다.")
            return False

        self.source_file = file_path
        print(f"✅ 선택된 파일: {self.source_file}")
        print(f"   파일 크기: {os.path.getsize(file_path) / 1024:.2f} KB")
        return True

    def select_output_folder(self) -> bool:
        """
        출력 폴더 선택 대화상자 표시

        Returns:
            bool: 폴더가 선택되면 True
        """
        print("\n[Step 2] 출력 폴더 선택")
        print("=" * 50)

        root = tk.Tk()
        root.withdraw()

        folder_path = filedialog.askdirectory(
            title="저장할 폴더를 선택하세요"
        )

        root.destroy()

        if not folder_path:
            print("❌ 폴더가 선택되지 않았습니다.")
            return False

        self.output_folder = folder_path
        print(f"✅ 선택된 폴더: {self.output_folder}")
        return True

    def open_file_with_default_app(self) -> bool:
        """
        파일을 기본 앱으로 열기

        Returns:
            bool: 성공하면 True
        """
        print("\n[Step 3] 파일을 기본 앱으로 열기")
        print("=" * 50)

        try:
            # Windows에서 기본 앱으로 파일 열기
            os.startfile(self.source_file)
            
            file_name = os.path.basename(self.source_file)
            ext = os.path.splitext(file_name)[1].upper()
            
            app_name = self._get_app_name(ext)
            print(f"✅ {file_name}을 {app_name}(로)로 열었습니다.")
            
            return True

        except Exception as e:
            print(f"❌ 파일 열기 실패: {e}")
            traceback.print_exc()
            return False

    def _get_app_name(self, extension: str) -> str:
        """
        확장자에 따른 앱 이름 반환

        Args:
            extension: 파일 확장자 (예: .XLSX)

        Returns:
            str: 앱 이름
        """
        app_map = {
            ".XLSX": "Excel",
            ".XLS": "Excel",
            ".DOCX": "Word",
            ".DOC": "Word",
            ".PDF": "PDF Reader",
            ".PPT": "PowerPoint",
            ".PPTX": "PowerPoint",
        }
        return app_map.get(extension, "기본 앱")

    def initialize_outlook(self) -> bool:
        """
        Outlook 초기화

        Returns:
            bool: 성공하면 True
        """
        print("\n[Step 4] Outlook 초기화")
        print("=" * 50)

        try:
            # 실행 중인 Outlook 인스턴스 사용
            try:
                self.outlook = GetObject(Class="Outlook.Application")
                print("✅ 기존 Outlook 인스턴스 사용")
            except:
                # Outlook이 없으면 새로 실행
                self.outlook = win32.Dispatch("Outlook.Application")
                print("✅ 새로운 Outlook 인스턴스 생성")

            return True

        except Exception as e:
            print(f"❌ Outlook 초기화 실패: {e}")
            print("   Outlook이 설치되어 있는지 확인하세요.")
            traceback.print_exc()
            return False

    def create_mail_with_attachment(self) -> bool:
        """
        파일이 첨부된 이메일 메시지 생성

        Returns:
            bool: 성공하면 True
        """
        print("\n[Step 5] 이메일 메시지 생성 및 파일 첨부")
        print("=" * 50)

        try:
            # 새 메일 항목 생성
            # 0 = olMailItem
            self.mail_item = self.outlook.CreateItem(0)

            # 메일 속성 설정
            self.mail_item.Subject = f"첨부: {os.path.basename(self.source_file)}"
            self.mail_item.Body = f"파일이 첨부되었습니다: {self.source_file}"

            # 파일 첨부
            self.mail_item.Attachments.Add(
                FileName=self.source_file,
                Position=0,
                DisplayName=os.path.basename(self.source_file),
            )

            print(f"✅ 메일에 파일 첨부: {os.path.basename(self.source_file)}")
            print(f"   첨부 파일 수: {self.mail_item.Attachments.Count}")

            # 메일 표시 (사용자에게 보여줌)
            self.mail_item.Display()
            print("✅ Outlook에서 메시지 표시")

            return True

        except Exception as e:
            print(f"❌ 메일 생성/첨부 실패: {e}")
            traceback.print_exc()
            return False

    def save_attachment(self) -> bool:
        """
        첨부 파일을 output 폴더에 저장

        Returns:
            bool: 성공하면 True
        """
        print("\n[Step 6] 첨부 파일 저장")
        print("=" * 50)

        try:
            if not self.mail_item or self.mail_item.Attachments.Count == 0:
                print("❌ 첨부 파일이 없습니다.")
                return False

            # 첫 번째 첨부 파일 저장
            attachment = self.mail_item.Attachments[1]  # COM은 1부터 시작

            # 저장 경로 생성
            output_path = os.path.join(
                self.output_folder, attachment.FileName
            )

            # 파일이 이미 존재하면 다른 이름으로 저장
            if os.path.exists(output_path):
                base_name = os.path.splitext(attachment.FileName)[0]
                extension = os.path.splitext(attachment.FileName)[1]
                output_path = os.path.join(
                    self.output_folder, f"{base_name}_saved{extension}"
                )
                print(f"⚠️  파일이 이미 존재합니다. 다른 이름으로 저장합니다.")

            # 파일 저장
            attachment.SaveAsFile(output_path)

            print(f"✅ 첨부 파일이 저장되었습니다.")
            print(f"   저장 경로: {output_path}")
            print(f"   파일 크기: {os.path.getsize(output_path) / 1024:.2f} KB")

            return True

        except Exception as e:
            print(f"❌ 첨부 파일 저장 실패: {e}")
            traceback.print_exc()
            return False

    def cleanup(self):
        """정리 작업"""
        if self.mail_item:
            try:
                # 메일 항목 닫기 (저장하지 않음)
                self.mail_item.Close(0)
            except:
                pass

    def run(self) -> bool:
        """
        전체 자동화 프로세스 실행

        Returns:
            bool: 성공하면 True
        """
        print("\n" + "=" * 50)
        print("🔗 Outlook 파일 첨부 자동화 시작")
        print("=" * 50)

        try:
            # Step 1: 파일 선택
            if not self.select_file():
                return False

            # Step 2: 출력 폴더 선택
            if not self.select_output_folder():
                return False

            # Step 3: 파일을 앱으로 열기
            if not self.open_file_with_default_app():
                return False

            # Step 4: Outlook 초기화
            if not self.initialize_outlook():
                return False

            # Step 5: 메일 생성 및 첨부
            if not self.create_mail_with_attachment():
                return False

            # Step 6: 첨부 파일 저장
            print("\n" + "=" * 50)
            print("⚠️  사용자 확인 필요")
            print("=" * 50)
            print("Outlook 창이 열렸습니다.")
            print("메일의 첨부 파일을 확인하시고,")
            print("콘솔에서 엔터를 눌러 저장을 진행하세요.")
            print("또는 Outlook에서 직접 '다른 이름으로 저장'을 사용할 수 있습니다.")
            input("\n엔터를 눌러 계속하세요...")

            if not self.save_attachment():
                return False

            print("\n" + "=" * 50)
            print("✅ 모든 작업이 완료되었습니다!")
            print("=" * 50)
            print(f"저장 위치: {self.output_folder}")

            return True

        except Exception as e:
            print(f"\n❌ 예상 밖의 오류 발생: {e}")
            traceback.print_exc()
            return False

        finally:
            self.cleanup()


def main():
    """메인 함수"""
    app = OutlookAutomation()
    success = app.run()

    if success:
        print("\n🎉 자동화 완료!")
        sys.exit(0)
    else:
        print("\n❌ 자동화 실패")
        sys.exit(1)


if __name__ == "__main__":
    main()
