import os
from tkinter import Tk, filedialog, messagebox, Label, Button, Entry
from PIL import Image

# 이미지 리사이징 함수
def resize_image(input_path, output_path, new_width):
    try:
        img = Image.open(input_path)
        
        # 원본 이미지 크기
        original_width, original_height = img.size
        
        # 비율에 맞는 새로운 높이 계산
        new_height = int((original_height / original_width) * new_width)
        
        # 리사이징
        resized_img = img.resize((new_width, new_height))
        
        # 파일명 추출
        file_name = os.path.basename(input_path)
        
        # 리사이징한 이미지 저장
        output_file_path = os.path.join(output_path, file_name)
        resized_img.save(output_file_path)
        
        return True, f"이미지가 성공적으로 저장되었습니다: {output_file_path}"
    except Exception as e:
        return False, f"이미지 리사이징 중 오류가 발생했습니다: {e}"

# 폴더 선택 함수
def select_folder():
    folder_path = filedialog.askdirectory(title="이미지 파일이 있는 폴더를 선택하세요")
    return folder_path

# 저장할 폴더 선택 함수
def select_output_folder():
    folder_path = filedialog.askdirectory(title="변경된 이미지를 저장할 폴더를 선택하세요")
    return folder_path

# GUI 창 생성
class ResizeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("이미지 리사이징 프로그램")
        self.root.geometry("570x200")
        
        self.input_folder = ""
        self.output_folder = ""
        
        self.create_widgets()

    def create_widgets(self):
        # 입력 폴더 선택 라벨 및 버튼
        self.input_label = Label(self.root, text="입력 디렉토리:")
        self.input_label.grid(row=0, column=0, padx=10, pady=10)

        self.input_folder_entry = Entry(self.root, width=50)
        self.input_folder_entry.grid(row=0, column=1, padx=10, pady=10)

        self.input_button = Button(self.root, text="폴더 선택", command=self.select_input_folder)
        self.input_button.grid(row=0, column=2, padx=10, pady=10)

        # 출력 폴더 선택 라벨 및 버튼
        self.output_label = Label(self.root, text="출력 디렉토리:")
        self.output_label.grid(row=1, column=0, padx=10, pady=10)

        self.output_folder_entry = Entry(self.root, width=50)
        self.output_folder_entry.grid(row=1, column=1, padx=10, pady=10)

        self.output_button = Button(self.root, text="폴더 선택", command=self.select_output_folder)
        self.output_button.grid(row=1, column=2, padx=10, pady=10)

        # 너비 입력 필드
        self.width_label = Label(self.root, text="원하는 너비:")
        self.width_label.grid(row=2, column=0, padx=10, pady=10)

        self.width_entry = Entry(self.root)
        self.width_entry.grid(row=2, column=1, padx=10, pady=10)

        # 리사이징 실행 버튼
        self.resize_button = Button(self.root, text="리사이징 실행", command=self.resize_images)
        self.resize_button.grid(row=3, column=0, columnspan=3, pady=20)

    def select_input_folder(self):
        self.input_folder = select_folder()
        if self.input_folder:
            self.input_folder_entry.delete(0, 'end')  # 기존 내용 삭제
            self.input_folder_entry.insert(0, self.input_folder)  # 경로 입력란에 경로 표시

    def select_output_folder(self):
        self.output_folder = select_output_folder()
        if self.output_folder:
            self.output_folder_entry.delete(0, 'end')  # 기존 내용 삭제
            self.output_folder_entry.insert(0, self.output_folder)  # 경로 입력란에 경로 표시

    def resize_images(self):
        # 사이즈 입력값 받아오기
        try:
            new_width = int(self.width_entry.get())
        except ValueError:
            messagebox.showerror("입력 오류", "너비는 숫자여야 합니다.")
            return

        # 입력 폴더가 선택되었는지 확인
        if not self.input_folder:
            messagebox.showerror("폴더 선택 오류", "이미지 폴더를 선택하세요.")
            return
        
        # 출력 폴더가 선택되었는지 확인
        if not self.output_folder:
            messagebox.showerror("폴더 선택 오류", "저장 폴더를 선택하세요.")
            return

        # 폴더 내 모든 이미지 리사이징
        success_count = 0
        error_count = 0

        for file_name in os.listdir(self.input_folder):
            input_file_path = os.path.join(self.input_folder, file_name)
            if file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                success, message = resize_image(input_file_path, self.output_folder, new_width)
                if success:
                    success_count += 1
                else:
                    error_count += 1

        # 결과 메시지
        result_message = f"리사이징 완료: {success_count}개\n실패: {error_count}개"
        messagebox.showinfo("작업 완료", result_message)


if __name__ == "__main__":
    root = Tk()
    app = ResizeApp(root)
    root.mainloop()
