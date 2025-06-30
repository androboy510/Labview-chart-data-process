import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
from typing import Optional

class ExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 파일 프로세서 v2.0 (수정된 버전)")
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
        self.tree = ttk.Treeview(preview_frame, show="headings", height=10)
        
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
        
        # 빠른 저장 버튼 (현재 폴더에 저장)
        self.quick_save_btn = ttk.Button(save_frame, text="빠른 저장", command=self.quick_save, state="disabled")
        self.quick_save_btn.pack(side=tk.LEFT, padx=(10, 0))
        
    def create_status_area(self, parent):
        """상태 표시 영역 생성"""
        # 상태 프레임
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=1)
        
        # 상태 라벨
        self.status_label = ttk.Label(status_frame, text="상태: 파일을 선택해주세요.", foreground="blue")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
    def select_file(self):
        """파일 선택 다이얼로그"""
        file_path = filedialog.askopenfilename(
            title="처리할 파일을 선택하세요",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.selected_file_path = file_path
            self.file_label.config(text=f"선택된 파일: {os.path.basename(file_path)}")
            self.process_file()
            
    def process_file(self):
        """파일 처리"""
        try:
            self.update_status("파일 처리 중...")
            self.root.update()
            
            # 파일 로드
            self.df = self.load_file(self.selected_file_path)
            
            # 데이터 처리
            self.processed_df = self.process_data(self.df)
            
            # 미리보기 업데이트
            self.update_preview()
            
            # 저장 버튼 활성화
            self.save_file_btn.config(state="normal")
            self.quick_save_btn.config(state="normal")
            
            self.update_status("파일 처리가 완료되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 처리 중 오류가 발생했습니다:\n{str(e)}")
            self.update_status("오류가 발생했습니다.")
            
    def load_file(self, file_path):
        """파일 로드"""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
        
        file_extension = os.path.splitext(file_path)[1].lower()
        
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
        df = self.clean_headers(df)
        
        # 2. Sample 열 처리
        df, sample_column_name = self.process_sample_columns(df)
        
        # 3. Time 열 추가
        df = self.add_time_column(df, sample_column_name)
        
        return df
        
    def clean_headers(self, df):
        """헤더 정리"""
        df.columns = df.columns.str.strip()
        return df
        
    def process_sample_columns(self, df):
        """Sample 열 처리"""
        sample_columns = [col for col in df.columns if 'sample' in col.lower()]
        
        if not sample_columns:
            raise ValueError("'sample'을 포함하는 열을 찾을 수 없습니다.")
        
        # 첫 번째 sample 열만 남기고 나머지 삭제
        first_sample_col = sample_columns[0]
        columns_to_drop = sample_columns[1:]
        
        if columns_to_drop:
            df = df.drop(columns=columns_to_drop)
        
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
            
    def quick_save(self):
        """빠른 저장 (현재 폴더에 저장)"""
        if self.processed_df is None:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
            
        try:
            # 현재 폴더에 자동으로 저장
            base_name = os.path.splitext(os.path.basename(self.selected_file_path))[0]
            output_file = f"{base_name}_processed.xlsx"
            
            self.update_status("빠른 저장 중...")
            self.root.update()
            
            # 파일 저장
            self.processed_df.to_excel(output_file, index=False)
            
            # 성공 메시지
            full_path = os.path.abspath(output_file)
            self.update_status("빠른 저장이 완료되었습니다.")
            messagebox.showinfo("빠른 저장 완료", 
                              f"파일이 현재 폴더에 저장되었습니다:\n\n{full_path}")
            
        except Exception as e:
            error_msg = f"빠른 저장 중 오류가 발생했습니다:\n{str(e)}"
            messagebox.showerror("오류", error_msg)
            self.update_status("빠른 저장 중 오류가 발생했습니다.")
            
    def save_file(self):
        """파일 저장 (사용자가 위치 선택)"""
        if self.processed_df is None:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
            
        try:
            # 기본 파일명 생성
            base_name = os.path.splitext(os.path.basename(self.selected_file_path))[0]
            default_filename = f"{base_name}_processed.xlsx"
            
            # 파일 저장 다이얼로그
            file_path = filedialog.asksaveasfilename(
                title="파일 저장",
                defaultextension=".xlsx",
                initialvalue=default_filename,
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("CSV files", "*.csv")
                ]
            )
            
            if file_path:  # 사용자가 파일명을 입력하고 저장을 클릭한 경우
                self.update_status("파일 저장 중...")
                self.root.update()
                
                file_extension = os.path.splitext(file_path)[1].lower()
                
                if file_extension == '.xlsx':
                    self.processed_df.to_excel(file_path, index=False)
                elif file_extension == '.csv':
                    self.processed_df.to_csv(file_path, index=False)
                
                self.update_status("파일이 성공적으로 저장되었습니다.")
                messagebox.showinfo("완료", f"파일이 저장되었습니다:\n{file_path}")
            else:
                # 사용자가 취소한 경우
                self.update_status("파일 저장이 취소되었습니다.")
                
        except Exception as e:
            error_msg = f"파일 저장 중 오류가 발생했습니다:\n{str(e)}"
            messagebox.showerror("오류", error_msg)
            self.update_status("파일 저장 중 오류가 발생했습니다.")
                
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