import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
from typing import Optional

class ExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 파일 프로세서 v2.0 (디버그 모드)")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 데이터 저장 변수
        self.df: Optional[pd.DataFrame] = None
        self.processed_df: Optional[pd.DataFrame] = None
        self.selected_file_path: str = ""
        
        # GUI 초기화
        self.setup_gui()
        
    def setup_gui(self):
        """GUI 레이아웃 설정"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 1. 파일 선택 영역
        self.create_file_selection_area(main_frame)
        
        # 2. 미리보기 영역
        self.create_preview_area(main_frame)
        
        # 3. 파일 저장 영역
        self.create_save_area(main_frame)
        
        # 4. 상태 표시 영역
        self.create_status_area(main_frame)
        
        # 5. 디버그 정보 영역
        self.create_debug_area(main_frame)
        
    def create_file_selection_area(self, parent):
        """파일 선택 영역 생성"""
        # 파일 선택 프레임
        file_frame = ttk.LabelFrame(parent, text="파일 선택", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # 파일 선택 버튼
        self.select_file_btn = ttk.Button(file_frame, text="파일 선택", command=self.select_file)
        self.select_file_btn.grid(row=0, column=0, padx=(0, 10))
        
        # 선택된 파일명 표시
        self.file_label = ttk.Label(file_frame, text="선택된 파일: 없음", foreground="gray")
        self.file_label.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
    def create_preview_area(self, parent):
        """미리보기 영역 생성"""
        # 미리보기 프레임
        preview_frame = ttk.LabelFrame(parent, text="파일 미리보기", padding="10")
        preview_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Treeview (테이블 형태)
        self.tree = ttk.Treeview(preview_frame, show="headings", height=8)
        
        # 스크롤바
        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 그리드 배치
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 초기 메시지
        self.tree.insert("", "end", values=["파일을 선택해주세요."])
        
    def create_save_area(self, parent):
        """파일 저장 영역 생성"""
        # 저장 프레임
        save_frame = ttk.Frame(parent)
        save_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 파일 저장 버튼
        self.save_file_btn = ttk.Button(save_frame, text="파일 저장", command=self.save_file, state="disabled")
        self.save_file_btn.pack(side=tk.LEFT)
        
        # 테스트 저장 버튼 (디버그용)
        self.test_save_btn = ttk.Button(save_frame, text="테스트 저장", command=self.test_save)
        self.test_save_btn.pack(side=tk.LEFT, padx=(10, 0))
        
    def create_status_area(self, parent):
        """상태 표시 영역 생성"""
        # 상태 프레임
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        status_frame.columnconfigure(0, weight=1)
        
        # 상태 라벨
        self.status_label = ttk.Label(status_frame, text="상태: 파일을 선택해주세요.", foreground="blue")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
    def create_debug_area(self, parent):
        """디버그 정보 영역 생성"""
        # 디버그 프레임
        debug_frame = ttk.LabelFrame(parent, text="디버그 정보", padding="10")
        debug_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        debug_frame.columnconfigure(0, weight=1)
        debug_frame.rowconfigure(0, weight=1)
        
        # 디버그 텍스트 영역
        self.debug_text = tk.Text(debug_frame, height=6, wrap=tk.WORD)
        debug_scrollbar = ttk.Scrollbar(debug_frame, orient="vertical", command=self.debug_text.yview)
        self.debug_text.configure(yscrollcommand=debug_scrollbar.set)
        
        self.debug_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        debug_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 초기 디버그 메시지
        self.debug_log("프로그램이 시작되었습니다.")
        
    def debug_log(self, message):
        """디버그 메시지 로그"""
        self.debug_text.insert(tk.END, f"[DEBUG] {message}\n")
        self.debug_text.see(tk.END)
        self.root.update()
        
    def select_file(self):
        """파일 선택 다이얼로그"""
        self.debug_log("파일 선택 다이얼로그를 열었습니다.")
        
        file_path = filedialog.askopenfilename(
            title="처리할 파일을 선택하세요",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.debug_log(f"선택된 파일: {file_path}")
            self.selected_file_path = file_path
            self.file_label.config(text=f"선택된 파일: {os.path.basename(file_path)}")
            self.process_file()
        else:
            self.debug_log("파일이 선택되지 않았습니다.")
            
    def process_file(self):
        """파일 처리"""
        try:
            self.debug_log("파일 처리 시작")
            self.update_status("파일 처리 중...")
            self.root.update()
            
            # 파일 로드
            self.debug_log("파일 로드 중...")
            self.df = self.load_file(self.selected_file_path)
            self.debug_log(f"파일 로드 완료: {self.df.shape[0]}행 x {self.df.shape[1]}열")
            
            # 데이터 처리
            self.debug_log("데이터 처리 중...")
            self.processed_df = self.process_data(self.df)
            self.debug_log(f"데이터 처리 완료: {self.processed_df.shape[0]}행 x {self.processed_df.shape[1]}열")
            
            # 미리보기 업데이트
            self.debug_log("미리보기 업데이트 중...")
            self.update_preview()
            
            # 저장 버튼 활성화
            self.save_file_btn.config(state="normal")
            self.debug_log("저장 버튼이 활성화되었습니다.")
            
            self.update_status("파일 처리가 완료되었습니다.")
            self.debug_log("파일 처리 완료")
            
        except Exception as e:
            error_msg = f"파일 처리 중 오류가 발생했습니다:\n{str(e)}"
            self.debug_log(f"오류 발생: {error_msg}")
            messagebox.showerror("오류", error_msg)
            self.update_status("오류가 발생했습니다.")
            
    def load_file(self, file_path):
        """파일 로드"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
        
        file_extension = os.path.splitext(file_path)[1].lower()
        self.debug_log(f"파일 확장자: {file_extension}")
        
        if file_extension == '.xlsx':
            df = pd.read_excel(file_path)
        elif file_extension == '.csv':
            df = pd.read_csv(file_path)
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {file_extension}")
        
        return df
        
    def process_data(self, df):
        """데이터 처리"""
        # 1. 헤더 정리
        self.debug_log("헤더 정리 중...")
        df = self.clean_headers(df)
        
        # 2. Sample 열 처리
        self.debug_log("Sample 열 처리 중...")
        df, sample_column_name = self.process_sample_columns(df)
        self.debug_log(f"남겨진 sample 열: {sample_column_name}")
        
        # 3. Time 열 추가
        self.debug_log("Time 열 추가 중...")
        df = self.add_time_column(df, sample_column_name)
        
        return df
        
    def clean_headers(self, df):
        """헤더 정리"""
        df.columns = df.columns.str.strip()
        return df
        
    def process_sample_columns(self, df):
        """Sample 열 처리"""
        sample_columns = [col for col in df.columns if 'sample' in col.lower()]
        self.debug_log(f"발견된 sample 열들: {sample_columns}")
        
        if not sample_columns:
            raise ValueError("'sample'을 포함하는 열을 찾을 수 없습니다.")
        
        # 첫 번째 sample 열만 남기고 나머지 삭제
        first_sample_col = sample_columns[0]
        columns_to_drop = sample_columns[1:]
        
        if columns_to_drop:
            df = df.drop(columns=columns_to_drop)
            self.debug_log(f"삭제된 열들: {columns_to_drop}")
        
        return df, first_sample_col
        
    def add_time_column(self, df, sample_column_name):
        """Time 열 추가"""
        if sample_column_name is None:
            raise ValueError("sample 열이 없어 Time 열을 추가할 수 없습니다.")
        
        # Time 열 계산 (sample 값 * 0.01)
        time_values = df[sample_column_name] * 0.01
        
        # 첫 번째 위치에 Time 열 삽입
        df.insert(0, 'Time', time_values)
        
        return df
        
    def update_preview(self):
        """미리보기 업데이트"""
        # 기존 데이터 삭제
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if self.processed_df is None:
            return
            
        # 열 설정
        columns = list(self.processed_df.columns)
        self.tree["columns"] = columns
        
        # 열 헤더 설정
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, minwidth=50)
        
        # 데이터 추가 (처음 5행)
        for i, row in self.processed_df.head().iterrows():
            self.tree.insert("", "end", values=list(row))
            
    def test_save(self):
        """테스트 저장 (디버그용)"""
        self.debug_log("테스트 저장 시작")
        
        if self.processed_df is None:
            self.debug_log("저장할 데이터가 없습니다.")
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
            
        try:
            # 현재 디렉토리에 테스트 파일 저장
            test_file = "test_output.xlsx"
            self.debug_log(f"테스트 파일 저장: {test_file}")
            
            self.processed_df.to_excel(test_file, index=False)
            
            self.debug_log("테스트 저장 성공")
            messagebox.showinfo("테스트 완료", f"테스트 파일이 저장되었습니다:\n{os.path.abspath(test_file)}")
            
        except Exception as e:
            error_msg = f"테스트 저장 실패: {str(e)}"
            self.debug_log(error_msg)
            messagebox.showerror("테스트 실패", error_msg)
            
    def save_file(self):
        """파일 저장"""
        self.debug_log("파일 저장 시작")
        
        if self.processed_df is None:
            self.debug_log("저장할 데이터가 없습니다.")
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
            
        # 기본 파일명 생성
        base_name = os.path.splitext(os.path.basename(self.selected_file_path))[0]
        default_filename = f"{base_name}_processed.xlsx"
        self.debug_log(f"기본 파일명: {default_filename}")
        
        file_path = filedialog.asksaveasfilename(
            title="파일 저장",
            defaultextension=".xlsx",
            initialvalue=default_filename,
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv")
            ]
        )
        
        self.debug_log(f"사용자가 선택한 경로: {file_path}")
        
        if file_path:
            try:
                self.update_status("파일 저장 중...")
                self.root.update()
                
                file_extension = os.path.splitext(file_path)[1].lower()
                self.debug_log(f"저장할 파일 확장자: {file_extension}")
                
                if file_extension == '.xlsx':
                    self.debug_log("Excel 파일로 저장 중...")
                    self.processed_df.to_excel(file_path, index=False)
                elif file_extension == '.csv':
                    self.debug_log("CSV 파일로 저장 중...")
                    self.processed_df.to_csv(file_path, index=False)
                
                self.debug_log("파일 저장 성공")
                self.update_status("파일이 성공적으로 저장되었습니다.")
                messagebox.showinfo("완료", f"파일이 저장되었습니다:\n{file_path}")
                
            except Exception as e:
                error_msg = f"파일 저장 중 오류가 발생했습니다:\n{str(e)}"
                self.debug_log(f"저장 오류: {error_msg}")
                messagebox.showerror("오류", error_msg)
                self.update_status("파일 저장 중 오류가 발생했습니다.")
        else:
            self.debug_log("사용자가 저장을 취소했습니다.")
                
    def update_status(self, message):
        """상태 메시지 업데이트"""
        self.status_label.config(text=f"상태: {message}")
        self.root.update()

def main():
    """메인 함수"""
    root = tk.Tk()
    app = ExcelProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 