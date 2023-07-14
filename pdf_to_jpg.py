from pdf2image import convert_from_path

def convert_pdf_to_jpg(pdf_path, output_path):
    images = convert_from_path(pdf_path)
    for i, image in enumerate(images):
        image.save(f"{output_path}/page_{i+1}.jpg", "JPEG")

# 指定要转换的 PDF 文件路径和输出路径
pdf_path = "path/to/your/pdf_file.pdf"
output_path = "path/to/save/the/images"

# 调用函数进行转换
convert_pdf_to_jpg(pdf_path, output_path)
