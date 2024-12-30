from django.http import JsonResponse, FileResponse
from django.shortcuts import render
from django.conf import settings
from PIL import Image, ImageFilter, ImageEnhance
from openpyxl import Workbook
import pytesseract
import os

def preprocess_image(image):
    image = image.convert("L")
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    image = image.filter(ImageFilter.MedianFilter(size=3))
    return image

def convert_image(request):
    if request.method == 'POST' and request.FILES['image']:
        try:
            image_file = request.FILES['image']
            image = Image.open(image_file)

            # Preprocess image and extract text
            image = preprocess_image(image)
            extracted_text = pytesseract.image_to_string(image, lang="eng", config="--psm 6")

            # Save extracted text for debugging
            debug_path = os.path.join('output', 'debug_text.txt')
            with open(debug_path, 'w', encoding='utf-8') as debug_file:
                debug_file.write(extracted_text)

            # Create Excel file from extracted text
            workbook = Workbook()
            sheet = workbook.active
            rows = extracted_text.split("\n")
            for i, row in enumerate(rows):
                columns = row.split()
                for j, cell in enumerate(columns):
                    sheet.cell(row=i+1, column=j+1).value = cell

            excel_path = os.path.join('output', 'output.xlsx')
            workbook.save(excel_path)

            return FileResponse(open(excel_path, 'rb'), as_attachment=True, filename="output.xlsx")
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)
    return JsonResponse({'error': 'No image uploaded'}, status=400)

def upload_page(request):
    return render(request, 'upload.html')
