import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
import json
import threading
import time
from typing import Optional, List

class ExcelProcessorGUIv3:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 파일 프로세서 v3.0 - 고급 기능")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 데이터 저장 변수
        self.df: Optional[pd.DataFrame] = None
        self.processed_df: Optional[pd.DataFrame] = None
        self.selected_file_path: str = ""
        
        # 최근 사용 파일 목록
        self.recent_files: List[str] = []
        self.max_recent_files = 10
        self.config_file = "gui_config.json"
        self.load_recent_files()
        
        # GUI 초기화
        self.setup_gui()
        
        # 드래그 앤 드롭 설정
        self.setup_drag_drop()
        
    def setup_gui(self):
        """GUI 레이아웃 설정"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # 1. 파일 선택 영역
        self.create_file_selection_area(main_frame)
        
        # 2. 최근 사용 파일 영역
        self.create_recent_files_area(main_frame)
        
        # 3. 진행 상황 영역
        self.create_progress_area(main_frame)
        
        # 4. 미리보기 영역
        self.create_preview_area(main_frame)
        
        # 5. 파일 저장 영역
        self.create_save_area(main_frame)
        
        # 6. 상태 표시 영역
        self.create_status_area(main_frame)
        
    def create_file_selection_area(self, parent):
        """파일 선택 영역 생성"""
        # 파일 선택 프레임
        file_frame = ttk.LabelFrame(parent, text="파일 선택", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # 파일 선택 버튼
        self.select_file_btn = ttk.Button(file_frame, text="📁 파일 선택", command=self.select_file)
        self.select_file_btn.grid(row=0, column=0, padx=(0, 10))
        
        # 선택된 파일명 표시
        self.file_label = ttk.Label(file_frame, text="선택된 파일: 없음", foreground="gray")
        self.file_label.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 드래그 앤 드롭 안내
        drag_label = ttk.Label(file_frame, text="💡 파일을 여기에 드래그 앤 드롭하세요", 
                              foreground="blue", font=("Arial", 9))
        drag_label.grid(row=1, column=0, columnspan=2, pady=(5, 0))
        
    def create_recent_files_area(self, parent):
        """최근 사용 파일 영역 생성"""
        # 최근 파일 프레임
        recent_frame = ttk.LabelFrame(parent, text="최근 사용 파일", padding="10")
        recent_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        recent_frame.columnconfigure(0, weight=1)
        
        # 최근 파일 리스트박스
        self.recent_listbox = tk.Listbox(recent_frame, height=3, selectmode=tk.SINGLE)
        recent_scrollbar = ttk.Scrollbar(recent_frame, orient="vertical", command=self.recent_listbox.yview)
        self.recent_listbox.configure(yscrollcommand=recent_scrollbar.set)
        
        self.recent_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E))
        recent_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 최근 파일 더블클릭 이벤트
        self.recent_listbox.bind('<Double-Button-1>', self.on_recent_file_select)
        
        # 최근 파일 목록 업데이트
        self.update_recent_files_display()
        
    def create_progress_area(self, parent):
        """진행 상황 영역 생성"""
        # 진행 상황 프레임
        progress_frame = ttk.LabelFrame(parent, text="진행 상황", padding="10")
        progress_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        # 진행 바
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=300)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 진행 상태 라벨
        self.progress_label = ttk.Label(progress_frame, text="대기 중...", foreground="gray")
        self.progress_label.grid(row=1, column=0, sticky=tk.W)
        
    def create_preview_area(self, parent):
        """미리보기 영역 생성"""
        # 미리보기 프레임
        preview_frame = ttk.LabelFrame(parent, text="파일 미리보기", padding="10")
        preview_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
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
        self.tree.insert("", "end", values=["파일을 선택하거나 드래그 앤 드롭해주세요."])
        
    def create_save_area(self, parent):
        """파일 저장 영역 생성"""
        # 저장 프레임
        save_frame = ttk.Frame(parent)
        save_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 파일 저장 버튼
        self.save_file_btn = ttk.Button(save_frame, text="💾 파일 저장", command=self.save_file, state="disabled")
        self.save_file_btn.pack(side=tk.LEFT)
        
        # 빠른 저장 버튼
        self.quick_save_btn = ttk.Button(save_frame, text="⚡ 빠른 저장", command=self.quick_save, state="disabled")
        self.quick_save_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # 최근 파일 목록 지우기 버튼
        self.clear_recent_btn = ttk.Button(save_frame, text="🗑️ 최근 파일 지우기", command=self.clear_recent_files)
        self.clear_recent_btn.pack(side=tk.RIGHT)
        
    def create_status_area(self, parent):
        """상태 표시 영역 생성"""
        # 상태 프레임
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=1)
        
        # 상태 라벨
        self.status_label = ttk.Label(status_frame, text="상태: 파일을 선택해주세요.", foreground="blue")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
    def setup_drag_drop(self):
        """드래그 앤 드롭 설정 (간단한 구현)"""
        # 드래그 앤 드롭을 위한 이벤트 바인딩
        self.root.bind('<B1-Motion>', self.on_drag)
        self.root.bind('<ButtonRelease-1>', self.on_drop)
        
    def on_drag(self, event):
        """드래그 이벤트 (간단한 구현)"""
        pass
        
    def on_drop(self, event):
        """드롭 이벤트 (간단한 구현)"""
        # 실제 드래그 앤 드롭은 tkinterdnd2 라이브러리가 필요하므로
        # 여기서는 파일 선택 다이얼로그를 열도록 함
        self.select_file()
        
    def load_recent_files(self):
        """최근 사용 파일 목록 로드"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.recent_files = config.get('recent_files', [])
                    # 존재하지 않는 파일 제거
                    self.recent_files = [f for f in self.recent_files if os.path.exists(f)]
        except Exception as e:
            print(f"최근 파일 목록 로드 실패: {e}")
            self.recent_files = []
            
    def save_recent_files(self):
        """최근 사용 파일 목록 저장"""
        try:
            config = {'recent_files': self.recent_files}
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"최근 파일 목록 저장 실패: {e}")
            
    def add_recent_file(self, file_path):
        """최근 사용 파일 목록에 추가"""
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
        self.recent_files.insert(0, file_path)
        
        # 최대 개수 제한
        if len(self.recent_files) > self.max_recent_files:
            self.recent_files = self.recent_files[:self.max_recent_files]
            
        self.save_recent_files()
        self.update_recent_files_display()
        
    def update_recent_files_display(self):
        """최근 파일 목록 표시 업데이트"""
        self.recent_listbox.delete(0, tk.END)
        for file_path in self.recent_files:
            filename = os.path.basename(file_path)
            self.recent_listbox.insert(tk.END, filename)
            
    def on_recent_file_select(self, event):
        """최근 파일 선택 이벤트"""
        selection = self.recent_listbox.curselection()
        if selection:
            index = selection[0]
            if index < len(self.recent_files):
                file_path = self.recent_files[index]
                if os.path.exists(file_path):
                    self.selected_file_path = file_path
                    self.file_label.config(text=f"선택된 파일: {os.path.basename(file_path)}")
                    self.process_file()
                else:
                    messagebox.showerror("오류", f"파일을 찾을 수 없습니다: {file_path}")
                    # 존재하지 않는 파일 제거
                    self.recent_files.pop(index)
                    self.save_recent_files()
                    self.update_recent_files_display()
                    
    def clear_recent_files(self):
        """최근 파일 목록 지우기"""
        if messagebox.askyesno("확인", "최근 사용 파일 목록을 모두 지우시겠습니까?"):
            self.recent_files.clear()
            self.save_recent_files()
            self.update_recent_files_display()
            
    def update_progress(self, value, message):
        """진행 상황 업데이트"""
        self.progress_bar['value'] = value
        self.progress_label.config(text=message)
        self.root.update()
        
    def select_file(self):
        """여러 파일 선택 다이얼로그 및 대기열 추가"""
        file_paths = filedialog.askopenfilenames(
            title="처리할 파일을 선택하세요",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_paths:
            self.file_queue = list(file_paths)
            self.process_file_queue()

    def process_file_queue(self):
        """대기열에 있는 파일을 순차적으로 처리"""
        if not hasattr(self, 'file_queue') or not self.file_queue:
            self.update_status("대기열이 비어 있습니다.")
            return
        self.current_file_index = 0
        self.total_files = len(self.file_queue)
        self.process_next_file()

    def process_next_file(self):
        if self.current_file_index >= self.total_files:
            self.update_status("모든 파일 처리가 완료되었습니다.")
            self.progress_bar['value'] = 100
            return
        file_path = self.file_queue[self.current_file_index]
        self.selected_file_path = file_path
        self.file_label.config(text=f"처리 중: {os.path.basename(file_path)} ({self.current_file_index+1}/{self.total_files})")
        self.progress_bar['value'] = (self.current_file_index / self.total_files) * 100
        thread = threading.Thread(target=self._process_and_save_current_file)
        thread.daemon = True
        thread.start()

    def _process_and_save_current_file(self):
        try:
            self.update_progress((self.current_file_index / self.total_files) * 100, f"{os.path.basename(self.selected_file_path)} 처리 중...")
            self.df = self.load_file(self.selected_file_path)
            self.processed_df = self.process_data(self.df)
            # 빠른 저장(원본 폴더)
            base_name = os.path.splitext(os.path.basename(self.selected_file_path))[0]
            dir_name = os.path.dirname(self.selected_file_path)
            output_file = os.path.join(dir_name, f"{base_name}_processed.xlsx")
            self.processed_df.to_excel(output_file, index=False)
            self.update_progress(((self.current_file_index+1) / self.total_files) * 100, f"{os.path.basename(self.selected_file_path)} 저장 완료!")
        except Exception as e:
            self.update_status(f"오류: {os.path.basename(self.selected_file_path)} - {str(e)}")
        finally:
            self.current_file_index += 1
            self.root.after(500, self.process_next_file)
        
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
        """빠른 저장 (원본 파일 폴더에 저장)"""
        if self.processed_df is None:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
            
        try:
            # 원본 파일 폴더에 저장
            base_name = os.path.splitext(os.path.basename(self.selected_file_path))[0]
            dir_name = os.path.dirname(self.selected_file_path)
            output_file = os.path.join(dir_name, f"{base_name}_processed.xlsx")
            
            self.update_status("빠른 저장 중...")
            self.root.update()
            
            # 파일 저장
            self.processed_df.to_excel(output_file, index=False)
            
            # 성공 메시지
            full_path = os.path.abspath(output_file)
            self.update_status("빠른 저장이 완료되었습니다.")
            messagebox.showinfo("빠른 저장 완료", 
                              f"파일이 원본 폴더에 저장되었습니다:\n\n{full_path}")
            
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
                initialfile=default_filename,
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
    app = ExcelProcessorGUIv3(root)
    root.mainloop()

if __name__ == "__main__":
    main() 