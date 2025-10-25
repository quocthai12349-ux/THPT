import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Alignment
import os
import re
from PIL import Image

# === CẤU HÌNH ===
GITHUB_REPO = "quocthai12349-ux/THPT"  # repo GitHub chứa ảnh
folder_path = "."                       # thư mục chứa file PDF
excel_path = os.path.join(folder_path, "tracnghiem_tonghop.xlsx")
image_dir = os.path.join(folder_path, "images")
os.makedirs(image_dir, exist_ok=True)


# === HÀM HỖ TRỢ ===
def clean_title(s, fallback):
    """Lọc và lấy tên bài từ dòng đầu PDF, nếu không có thì lấy từ tên file."""
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    match = re.search(r"(Bài|Dạng|Bổ trợ)\s*\d*[:\-–]?\s*(.*)", s, re.IGNORECASE)
    if match:
        s = match.group(0)
    s = re.sub(r"(?i)\b(group|vật lý|physics|đề|trắc nghiệm|đúng sai)\b", "", s)
    s = s.strip(" -:").title()
    if not s or len(s) < 3:
        s = fallback.replace("(đề)", "").replace(".pdf", "").strip()
    return s


def extract_images(page, pdf_name):
    """Trích xuất toàn bộ ảnh từ 1 trang PDF."""
    images = []
    for i, img in enumerate(page.get_images(full=True)):
        xref = img[0]
        base_image = page.parent.extract_image(xref)
        img_bytes = base_image["image"]
        img_ext = base_image.get("ext", "png")
        img_name = f"{os.path.splitext(pdf_name)[0]}_p{page.number + 1}_img{i + 1}.{img_ext}"
        img_path = os.path.join(image_dir, img_name)
        with open(img_path, "wb") as f:
            f.write(img_bytes)
        images.append(img_name)
    return images


def format_question(text):
    """Chuẩn hóa câu hỏi: xuống dòng cho các đáp án a,b,c,d."""
    text = re.sub(r"(?<=\S)-\s+(?=\S)", "", text)  # nối các từ bị ngắt
    text = re.sub(r"\n+", " ", text)

    # xuống dòng cho a), A), a., A. hoặc có dấu gạch ngang trước
    text = re.sub(r"\s*([–-]?\s*[a-dA-D][\)\.])", r"\n\1", text)

    # thay các khoảng trắng dư thành 1 space
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


# === XỬ LÝ PDF ===
def process_pdfs():
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("⚠️ Không tìm thấy file PDF trong thư mục.")
        return []

    rows = []
    for pdf_name in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_name)
        doc = fitz.open(pdf_path)

        # Lấy tên bài từ dòng đầu tiên hợp lý
        first_page_text = doc.load_page(0).get_text("text")
        first_line = ""
        for ln in first_page_text.splitlines():
            if any(k in ln.lower() for k in ["bài", "dạng", "bổ trợ"]):
                first_line = ln
                break
        title = clean_title(first_line or first_page_text[:100], pdf_name)

        print(f"📘 Đang xử lý: {pdf_name}  →  {title}")

        for p in range(doc.page_count):
            page = doc.load_page(p)
            text = page.get_text("text")
            parts = re.split(r"(?i)(?=Câu\s*\d+[:.)])", text)
            imgs = extract_images(page, pdf_name)

            for part in parts:
                part = part.strip()
                if len(part) < 10:
                    continue
                q_text = format_question(part)

                # Ảnh ở trên đầu câu hỏi
                img_tags = ""
                if imgs:
                    for im in imgs:
                        link = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/images/{im}"
                        img_tags += f'<img src="{link}">\n'

                full_text = f"{img_tags.strip()}\n{q_text}" if img_tags else q_text
                rows.append((full_text.strip(), title))
    return rows


# === XUẤT EXCEL ===
def export_excel(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "tracnghiem"
    ws["A1"] = "Câu hỏi"
    ws["C1"] = "Tên bài"

    for i, (q, t) in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=q)
        ws.cell(row=i, column=3, value=t)
        ws.cell(row=i, column=1).alignment = Alignment(wrapText=True, vertical="top")
        ws.cell(row=i, column=3).alignment = Alignment(wrapText=True, vertical="top")

    wb.save(excel_path)
    print(f"\n✅ Đã tạo file Excel: {excel_path}")
    print(f"🖼️ Ảnh được lưu trong thư mục: {image_dir}")


# === MAIN ===
if __name__ == "__main__":
    rows = process_pdfs()
    if rows:
        export_excel(rows)
    else:
        print("Không có dữ liệu để ghi Excel.")
