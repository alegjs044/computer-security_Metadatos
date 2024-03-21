
from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS, IFD
import os
from docx import Document
from openpyxl import load_workbook
import PyPDF2

def get_image_metadata(image_path):
    image_metadata = {}
    image_object = Image.open(image_path)
    exif_data = image_object.getexif()
    if exif_data:
        for tag_id, value in exif_data.items():
            tag = TAGS.get(tag_id)
            image_metadata[tag] = value
            if isinstance(value, bytes):
                value = value.decode()
            print(f"{tag}: {value}")
    return image_metadata

def get_geo(exif):
    gps_info = exif.get(IFD.GPSInfo)
    gps_data = {}
    if gps_info:
        for key, value in gps_info.items():
            tag = GPSTAGS.get(key)
            gps_data[tag] = value
    return gps_data

def get_coordinates(gps_data):
    lat_ref = 1 if gps_data['GPSLatitudeRef'] == 'N' else -1
    long_ref = 1 if gps_data['GPSLongitudeRef'] == 'E' else -1
    latitude = gps_data['GPSLatitude']
    latitude = float(latitude[0] + latitude[1] / 60 + latitude[2] / 3600)
    longitude = gps_data['GPSLongitude']
    longitude = float(longitude[0] + longitude[1] / 60 + longitude[2] / 3600)
    return latitude * lat_ref, longitude * long_ref

def extract_docx_metadata(docx_path):
    docx_metadata = {}
    doc = Document(docx_path)
    for prop in doc.core_properties:
        docx_metadata[prop] = str(getattr(doc.core_properties, prop))
        print(f"{prop}: {docx_metadata[prop]}")
    return docx_metadata

def extract_xlsx_metadata(xlsx_path):
    xlsx_metadata = {}
    wb = load_workbook(filename=xlsx_path)
    for prop in wb.properties:
        xlsx_metadata[prop] = str(getattr(wb.properties, prop))
        print(f"{prop}: {xlsx_metadata[prop]}")
    return xlsx_metadata

def extract_pdf_metadata(pdf_path):
    pdf_metadata = {}
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfFileReader(f)
        doc_info = reader.getDocumentInfo()
        for key, value in doc_info.items():
            pdf_metadata[key] = value
            print(f"{key}: {value}")
    return pdf_metadata

if __name__ == "__main__":
    directory_path = input("Ingrese la ruta del directorio: ")
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
            print(f"\n∞∞∞ Metadatos de la Imagen: {filename} ∞∞∞")
            get_image_metadata(file_path)
            gps_results = get_geo(Image.open(file_path).getexif())
            if gps_results:
                print("\n∞∞∞ Datos GPS ∞∞∞")
                print(gps_results)
                print("\n∞∞∞ Coordenadas (Latitud-Longitud) de la imagen ∞∞∞")
                print(get_coordinates(gps_results))
        elif filename.lower().endswith('.docx'):
            print(f"\n∞∞∞ Metadatos del Documento Word: {filename} ∞∞∞")
            extract_docx_metadata(file_path)
        elif filename.lower().endswith('.xlsx'):
            print(f"\n∞∞∞ Metadatos de la Hoja de Cálculo Excel: {filename} ∞∞∞")
            extract_xlsx_metadata(file_path)
        elif filename.lower().endswith('.pdf'):
            print(f"\n∞∞∞ Metadatos del Documento PDF: {filename} ∞∞∞")
            extract_pdf_metadata(file_path)
        else:
            print(f"No se pudo extraer metadatos de {filename}: formato no compatible.")

