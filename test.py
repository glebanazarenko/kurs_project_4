def get_metadata(self, book_path: str) -> Tuple:
        # Возвращает кортеж с метаданными книги (атрибуты, количество страниц и т.д.)
        with open(book_path, 'rb') as file:
            pdf = PdfFileReader(file)
            info = pdf.getDocumentInfo()
            return (info.title, info.author, info.created.year, pdf.getNumPages())

def get_preview(self, book_path: str) -> Image.Image:
        # Возвращает изображение превью (скриншот 1-й страницы) книги в виде байтов
        images = convert_from_path(book_path, dpi=200, first_page=1, last_page=1)
        image = images[0]

        # Конвертируем PIL Image в bytes
        byte_arr = io.BytesIO()
        image.save(byte_arr, format='PNG')
        return byte_arr.getvalue()