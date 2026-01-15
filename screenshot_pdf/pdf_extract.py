import cv2
import numpy as np
from pdf2image import convert_from_path

# ---- 使用者設定 ----
pdf_path = r"I:\KM\manual.pdf"      # 輸入PDF檔案
page_start = 1              # 頁數: 開始頁 (1起算)
page_end = 174                # 頁數: 結束頁（含）。例如1~5表示第1到5頁
poppler_path = None         # 若系統找不到poppler，需寫如 r"C:\poppler-xx\Library\bin"
# ---------------------

# 1. 讀取 PDF、轉換成圖片（每頁一張）
pages = convert_from_path(
    pdf_path, 
    fmt='png', 
    first_page=page_start, 
    last_page=page_end, 
    poppler_path=poppler_path
)

for idx, page in enumerate(pages, start=page_start):
    img = np.array(page)
    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)  # PIL 轉 OpenCV 格式

    # 影像處理流程 (基本等同於 crop.py)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blurred, 50, 150)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    dilated = cv2.dilate(edges, kernel, iterations=3)
    closed = cv2.morphologyEx(dilated, cv2.MORPH_CLOSE, kernel, iterations=3)
    contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)

    # 輸出（每頁）
    # cv2.imwrite(f"page_{idx}_gray.jpg", gray)
    # cv2.imwrite(f"page_{idx}_blur.jpg", blurred)
    # cv2.imwrite(f"page_{idx}_edges.jpg", edges)
    # cv2.imwrite(f"page_{idx}_dilated.jpg", dilated)
    # cv2.imwrite(f"page_{idx}_closed.jpg", closed)
    cv2.imwrite(f"page_{idx}_output.jpg", img)

print(f"已處理第 {page_start} 到 {page_end} 頁 PDF。")

# 如需裁剪效果可針對最大輪廓額外加上
# if contours:
#     best = max(contours, key=cv2.contourArea)
#     x, y, w, h = cv2.boundingRect(best)
#     cropped = img[y:y + h, x:x + w]
#     cv2.imwrite(f"page_{idx}_cropped.jpg", cropped)