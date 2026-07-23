import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re

class TextToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("계층형 텍스트 엑셀 분리 및 변환기")
        self.root.geometry("800x650")
        
        self.parsed_data = []
        self.setup_ui()

    def setup_ui(self):
        # 상단: 텍스트 입력 영역
        top_frame = tk.LabelFrame(self.root, text="1. 원본 텍스트 붙여넣기 (Ctrl+V)", padx=10, pady=10)
        top_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.text_input = tk.Text(top_frame, height=10)
        self.text_input.pack(fill=tk.BOTH, expand=True)

        # 중앙: 컨트롤 버튼 영역
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        convert_btn = tk.Button(btn_frame, text="⬇️ 데이터 분석 및 표 변환", command=self.process_text, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        convert_btn.pack(side=tk.LEFT, padx=5)
        
        copy_all_btn = tk.Button(btn_frame, text="📋 전체 클립보드 복사", command=self.copy_all_to_clipboard)
        copy_all_btn.pack(side=tk.LEFT, padx=5)
        
        export_btn = tk.Button(btn_frame, text="💾 엑셀 파일로 저장", command=self.export_to_excel)
        export_btn.pack(side=tk.LEFT, padx=5)

        # 하단: 결과 확인 및 선택 복사 영역 (Treeview)
        bottom_frame = tk.LabelFrame(self.root, text="2. 변환 결과", padx=10, pady=10)
        bottom_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # === 자동 복사 체크박스 추가 ===
        self.auto_copy_var = tk.BooleanVar(value=True) # 기본값: 체크됨(True)
        auto_copy_chk = tk.Checkbutton(
            bottom_frame, 
            text="마우스를 놓으면(드래그 해제) 자동으로 선택 영역 복사", 
            variable=self.auto_copy_var,
            fg="blue"
        )
        auto_copy_chk.pack(anchor=tk.W, pady=(0, 5))
        
        # 표(Treeview) 생성
        columns = ("A_col", "B_col")
        self.tree = ttk.Treeview(bottom_frame, columns=columns, show="headings")
        self.tree.heading("A_col", text="A열 (원문장/상위)")
        self.tree.heading("B_col", text="B열 (하부설명)")
        
        self.tree.column("A_col", width=350)
        self.tree.column("B_col", width=350)
        
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(bottom_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 이벤트 바인딩
        self.tree.bind("<Control-c>", self.copy_selected_to_clipboard) # 수동 복사 (Ctrl+C)
        self.tree.bind("<ButtonRelease-1>", self.on_mouse_release)     # 마우스 왼쪽 버튼을 놓았을 때

    def process_text(self):
        raw_text = self.text_input.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showwarning("경고", "분석할 텍스트를 입력해주세요.")
            return
            
        lines = raw_text.split('\n')
        self.parsed_data = []
        
        current_main = ""
        sub_items = []
        
        for line in lines:
            if not line.strip():
                continue
                
            is_sub = re.match(r'^[\s\t]+|^\s*[-*•>]\s*', line)
            clean_line = re.sub(r'^\s*[-*•>]*\s*', '', line).strip()
            
            if is_sub:
                if not current_main:
                    current_main = clean_line
                else:
                    sub_items.append(clean_line)
            else:
                if current_main:
                    if not sub_items:
                        self.parsed_data.append([current_main, current_main])
                    else:
                        for sub in sub_items:
                            self.parsed_data.append([current_main, sub])
                
                current_main = clean_line
                sub_items = []
        
        if current_main:
            if not sub_items:
                self.parsed_data.append([current_main, current_main])
            else:
                for sub in sub_items:
                    self.parsed_data.append([current_main, sub])
                    
        self.update_table()

    def update_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for row in self.parsed_data:
            self.tree.insert("", tk.END, values=row)

    def on_mouse_release(self, event):
        """마우스 왼쪽 버튼 클릭/드래그를 놓았을 때 실행되는 함수"""
        if self.auto_copy_var.get(): # 체크박스가 켜져있다면
            self.copy_selected_to_clipboard()

    def copy_selected_to_clipboard(self, event=None):
        """선택된 영역을 엑셀 형식(TSV)으로 클립보드에 복사"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
            
        copied_data = []
        for item in selected_items:
            values = self.tree.item(item, 'values')
            copied_data.append("\t".join(values))
            
        tsv_data = "\n".join(copied_data)
        self.root.clipboard_clear()
        self.root.clipboard_append(tsv_data)
        print("✅ 선택 영역 클립보드 자동 복사 완료") # 콘솔에 로그 출력

    def copy_all_to_clipboard(self):
        if not self.parsed_data:
            return
        tsv_data = "\n".join(["\t".join(row) for row in self.parsed_data])
        self.root.clipboard_clear()
        self.root.clipboard_append(tsv_data)
        messagebox.showinfo("알림", "전체 데이터가 클립보드에 복사되었습니다.\n엑셀에 바로 붙여넣기(Ctrl+V) 하세요.")

    def export_to_excel(self):
        if not self.parsed_data:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")],
            title="엑셀 파일로 저장"
        )
        
        if file_path:
            try:
                df = pd.DataFrame(self.parsed_data, columns=["A열 (상위)", "B열 (하위)"])
                df.to_excel(file_path, index=False)
                messagebox.showinfo("완료", f"파일이 성공적으로 저장되었습니다.\n{file_path}")
            except Exception as e:
                messagebox.showerror("오류", f"저장 중 문제가 발생했습니다:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TextToExcelApp(root)
    root.mainloop()