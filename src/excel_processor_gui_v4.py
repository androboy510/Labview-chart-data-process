import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
import json
import threading
import time
from typing import Optional, List
from tkinterdnd2 import DND_FILES, TkinterDnD

# 스타일 및 폰트 개선
MODERN_FONT = ('Inter', 13)
TITLE_FONT = ('Inter', 18, 'bold')
SUB_FONT = ('Inter', 11)

class ModernExcelProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Labview waveform chart data processor")
        # 화면 해상도 기준 최대화
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        win_w, win_h = int(screen_w * 0.6), int(screen_h * 0.9)
        self.root.geometry(f"{win_w}x{win_h}")
        self.root.minsize(900, 600)
        self.root.configure(bg="#F7F8FA")
        self.setup_modern_theme()
        self.df = None
        self.processed_df = None
        self.selected_file_path = ""
        self.setup_gui()

    def setup_modern_theme(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Card.TFrame', background="#FFFFFF", relief='flat', borderwidth=0)
        style.configure('Title.TLabel', font=TITLE_FONT, foreground="#222", background="#FFFFFF")
        style.configure('Subtitle.TLabel', font=SUB_FONT, foreground="#888", background="#FFFFFF")
        style.configure('Status.TLabel', font=SUB_FONT, foreground="#007AFF", background="#F7F8FA")
        style.configure('Primary.TButton', font=MODERN_FONT, foreground="#fff", background="#007AFF", borderwidth=0, padding=8)
        style.map('Primary.TButton', background=[('active', '#005BBB')])
        style.configure('Modern.Treeview', font=MODERN_FONT, background="#FAFAFA", fieldbackground="#FAFAFA", borderwidth=0)
        style.configure('Modern.Horizontal.TProgressbar', troughcolor="#E5E5EA", background="#007AFF", borderwidth=0)

    def setup_gui(self):
        # 메인 컨테이너
        main_frame = ttk.Frame(self.root, style='Card.TFrame', padding=30)
        main_frame.pack(fill='both', expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        # 헤더
        self.create_header(main_frame)
        # 파일 선택 카드
        self.create_file_selection_card(main_frame)
        # 진행 상황 카드
        self.create_progress_card(main_frame)
        # 미리보기 카드(여기만 스크롤)
        self.create_preview_card(main_frame)
        # 상태 바
        self.create_status_bar(main_frame)

    def create_header(self, parent):
        header = ttk.Frame(parent, style='Card.TFrame')
        header.grid(row=0, column=0, sticky='ew', pady=(0, 20))
        header.columnconfigure(1, weight=1)
        icon = ttk.Label(header, text="🧊", font=("Arial", 28), background="#FFFFFF")
        icon.grid(row=0, column=0, padx=(0, 15))
        title = ttk.Label(header, text="Labview waveform chart data processor", style='Title.TLabel')
        title.grid(row=0, column=1, sticky='w')
        # subtitle = ttk.Label(header, text="애플/스타트업 스타일의 미니멀 데이터 툴", style='Subtitle.TLabel')
        # subtitle.grid(row=1, column=1, sticky='w', pady=(5, 0))

    def create_file_selection_card(self, parent):
        card = ttk.Frame(parent, style='Card.TFrame', padding=24)
        card.grid(row=1, column=0, sticky='ew', pady=(0, 20))
        card.columnconfigure(1, weight=1)
        title = ttk.Label(card, text="📁 파일 선택", style='Title.TLabel')
        title.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        self.select_file_btn = ttk.Button(card, text="파일 선택", style='Primary.TButton', command=self.select_file)
        self.select_file_btn.grid(row=1, column=0, padx=(0, 15))
        self.file_label = ttk.Label(card, text="선택된 파일이 없습니다", style='Subtitle.TLabel')
        self.file_label.grid(row=1, column=1, sticky='w')
        drag_label = ttk.Label(card, text="💡 파일을 여기에 드래그 앤 드롭하거나 버튼을 클릭하세요", style='Subtitle.TLabel')
        drag_label.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        self.file_select_card = card

    def create_progress_card(self, parent):
        card = ttk.Frame(parent, style='Card.TFrame', padding=24)
        card.grid(row=2, column=0, sticky='ew', pady=(0, 20))
        card.columnconfigure(0, weight=1)
        title = ttk.Label(card, text="📊 진행 상황", style='Title.TLabel')
        title.grid(row=0, column=0, sticky='w', pady=(0, 10))
        self.progress_bar = ttk.Progressbar(card, style='Modern.Horizontal.TProgressbar', mode='determinate', length=400)
        self.progress_bar.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        self.progress_label = ttk.Label(card, text="대기 중...", style='Status.TLabel')
        self.progress_label.grid(row=2, column=0, sticky='w')

    def create_preview_card(self, parent):
        card = ttk.Frame(parent, style='Card.TFrame', padding=24)
        card.grid(row=3, column=0, sticky='nsew', pady=(0, 20))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(0, weight=1)
        title = ttk.Label(card, text="👁️ 데이터 미리보기", style='Title.TLabel')
        title.grid(row=0, column=0, sticky='w', pady=(0, 10))
        # 미리보기 테이블에만 스크롤 적용
        tree_frame = ttk.Frame(card, style='Card.TFrame')
        tree_frame.grid(row=1, column=0, sticky='nsew')
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_frame, style='Modern.Treeview', show="headings", height=12)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self.tree.insert("", "end", values=["파일을 선택하거나 드래그 앤 드롭해주세요."])

    def create_status_bar(self, parent):
        status = ttk.Frame(parent, style='Card.TFrame', padding=(0, 8))
        status.grid(row=4, column=0, sticky='ew')
        status.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(status, text="상태: 준비됨", style='Status.TLabel')
        self.status_label.grid(row=0, column=0, sticky='w')
        time_label = ttk.Label(status, text="v4.1", style='Status.TLabel')
        time_label.grid(row=0, column=1, sticky='e')

    def update_progress(self, value, message):
        """진행 상황 업데이트"""
        self.progress_bar['value'] = value
        self.progress_label.config(text=message)
        self.root.update()
        
    def refresh_interface(self):
        """인터페이스 새로고침"""
        self.update_status("인터페이스가 새로고침되었습니다.")
        
    def select_file(self):
        file_paths = filedialog.askopenfilenames(
            title="처리할 파일을 선택하세요",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if file_paths:
            if len(file_paths) == 1:
                self.selected_file_path = file_paths[0]
                self.show_preview_and_save_options()
            else:
                self.file_queue = list(file_paths)
                self.process_file_queue()

    def process_file_queue(self):
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
        thread = threading.Thread(target=self._process_and_quick_save_current_file)
        thread.daemon = True
        thread.start()

    def _process_and_quick_save_current_file(self):
        try:
            self.update_progress((self.current_file_index / self.total_files) * 100, f"{os.path.basename(self.selected_file_path)} 처리 중...")
            self.df = self.load_file(self.selected_file_path)
            self.processed_df = self.process_data(self.df)
            base_name = os.path.splitext(os.path.basename(self.selected_file_path))[0]
            dir_name = os.path.dirname(self.selected_file_path)
            output_file = os.path.join(dir_name, f"{base_name}_processed.xlsx")
            self.processed_df.to_excel(output_file, index=False)
            full_path = os.path.abspath(output_file)
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
        if self.processed_df is None:
            self.update_status("저장할 데이터가 없습니다.")
            return
        try:
            base_name = os.path.splitext(os.path.basename(self.selected_file_path))[0]
            dir_name = os.path.dirname(self.selected_file_path)
            output_file = os.path.join(dir_name, f"{base_name}_processed.xlsx")
            self.processed_df.to_excel(output_file, index=False)
            full_path = os.path.abspath(output_file)
            self.update_status(f"빠른 저장 완료: {full_path}")
        except Exception as e:
            self.update_status(f"빠른 저장 중 오류: {str(e)}")
            
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
        self.status_label.config(text=message)
        self.root.update()

    def show_preview_and_save_options(self):
        self.df = self.load_file(self.selected_file_path)
        self.processed_df = self.process_data(self.df)
        self.update_preview()
        # 저장 방식 선택 다이얼로그 없이 바로 빠른 저장
        self.quick_save()

    def on_drop_files(self, event):
        try:
            files = self.root.tk.splitlist(event.data)
            file_paths = [f for f in files if f.lower().endswith(('.xlsx', '.csv'))]
            if not file_paths:
                messagebox.showwarning("드래그 앤 드롭", "엑셀/CSV 파일만 지원합니다.")
                return
            if len(file_paths) == 1:
                self.selected_file_path = file_paths[0]
                self.show_preview_and_save_options()
            else:
                self.file_queue = list(file_paths)
                self.process_file_queue()
        except Exception as e:
            messagebox.showerror("드래그 앤 드롭 오류", f"드롭 이벤트 처리 중 오류 발생: {e}")

def main():
    """메인 함수"""
    root = TkinterDnD.Tk()
    app = ModernExcelProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 