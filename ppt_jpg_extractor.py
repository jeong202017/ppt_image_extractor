import os
import shutil
from pptx import Presentation
from tkinter import Tk, Label, Button, filedialog, messagebox
import threading
import re

def extract_images_from_all_pptx(input_dir, output_dir_images, output_dir_no_images):
    os.makedirs(output_dir_images, exist_ok=True)
    os.makedirs(output_dir_no_images, exist_ok=True)

    pptx_files = [f for f in os.listdir(input_dir) if f.endswith('.pptx')]
    if not pptx_files:
        messagebox.showwarning("경고", "선택한 폴더에 pptx 파일이 없습니다.")
        return

    for filename in pptx_files:
        pptx_path = os.path.join(input_dir, filename)
        prs = Presentation(pptx_path)
        image_count = 0

        # ppt_name = os.path.splitext(filename)[0]

        def sanitize_filename(name):
            # Windows에서 불가능한 문자 제거 또는 대체
            return re.sub(r'[\\/:"*?<>|]', '_', name)

        ppt_name = sanitize_filename(os.path.splitext(filename)[0])



        ppt_image_folder = os.path.join(output_dir_images, ppt_name)
        os.makedirs(ppt_image_folder, exist_ok=True)

        for slide_index, slide in enumerate(prs.slides):
            for shape_index, shape in enumerate(slide.shapes):
                if hasattr(shape, "image"):
                    image = shape.image
                    image_bytes = image.blob
                    content_type = image.content_type
                    ext = content_type.split('/')[-1]
                    image_filename = f"slide{slide_index+1}_img{shape_index+1}.{ext}"
                    image_path = os.path.join(ppt_image_folder, image_filename)

                    with open(image_path, 'wb') as f:
                        f.write(image_bytes)
                    image_count += 1

        if image_count == 0:
            shutil.copy(pptx_path, os.path.join(output_dir_no_images, filename))
            shutil.rmtree(ppt_image_folder)
            print(f"[❌] 이미지 없음: {filename}")
        else:
            print(f"[✅] {filename} → 이미지 {image_count}개 추출")

    messagebox.showinfo("완료", "이미지 추출이 완료되었습니다.")

def start_extraction():
    input_dir = filedialog.askdirectory(title="PPTX 폴더 선택")
    if not input_dir:
        return

    output_dir_images = os.path.join(input_dir, 'extracted_images')
    output_dir_no_images = os.path.join(input_dir, 'no_image_pptx')

    # 작업이 오래 걸릴 수 있으니 스레드로 처리
    threading.Thread(
        target=extract_images_from_all_pptx,
        args=(input_dir, output_dir_images, output_dir_no_images)
    ).start()

# GUI 구성
def run_gui():
    root = Tk()
    root.title("PPTX 이미지 추출기")
    root.geometry("400x200")

    Label(root, text="폴더 내 PPTX 파일에서 이미지를 추출합니다.", pady=20).pack()
    Button(root, text="PPTX 폴더 선택 및 이미지 추출", command=start_extraction, padx=20, pady=10).pack()
    Label(root, text="© 2025 PPT Extractor", pady=20).pack()

    root.mainloop()

if __name__ == "__main__":
    run_gui()