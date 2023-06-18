import os
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog, Menu, ttk
from bookAnalyzer import BookAnalyzer
import matplotlib.pyplot as plt
from PIL import Image, ImageTk
from io import BytesIO


class App:
    def __init__(self, root, analyzer):
        self.root = root
        self.analyzer = analyzer

        self.tree = ttk.Treeview(self.root, columns=('Title', 'Author', 'Num Pages'), show='headings')
        self.tree.heading('Title', text='Title')
        self.tree.heading('Author', text='Author')
        self.tree.heading('Num Pages', text='Num Pages')
        self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

        self.text_widget = tk.Text(self.root)
        self.text_widget.grid(row=1, column=0, columnspan=4, sticky="nsew")
        self.text_widget.configure(state='disabled')  # делаем текстовое поле read-only

        self.display_button = tk.Button(self.root, text="Показать книги", command=self.display_all_books)
        self.display_button.grid(row=0, column=0)

        self.process_directory_button = tk.Button(self.root, text="Обработать директорию", command=self.process_directory)
        self.process_directory_button.grid(row=0, column=1)

        self.reset_database_button = tk.Button(self.root, text="Перезапустить базу данных", command=self.reset_database)
        self.reset_database_button.grid(row=0, column=2)

        # Создание выпадающего меню
        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Основные функции", menu=file_menu)
        file_menu.add_command(label="Поиск книг по названию", command=self.search_books_by_title)
        file_menu.add_command(label="Поиск книг по автору", command=self.search_books_by_author)
        file_menu.add_command(label="Поиск книг по расширению", command=self.search_books_by_extension)
        file_menu.add_command(label="Самые большие книги", command=self.display_largest_books)
        file_menu.add_command(label="Книги с наибольшим количеством страниц", command=self.display_books_with_most_pages)
        file_menu.add_command(label="Недавно добавленные книги", command=self.display_recently_added_books)
        file_menu.add_command(label="Книги без автора", command=self.display_books_without_author)
        file_menu.add_command(label="Книги без метаданных", command=self.display_books_without_metadata)
        file_menu.add_command(label="Показать статистику расширений файлов", command=self.display_file_extension_statistics)
        file_menu.add_command(label="Показать график количества страниц", command=self.display_books_pages_chart)


        # Задаём растягиваемость строк и столбцов
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)  # Изменим номер строки для растягиваемости, т.к. текстовое поле теперь во второй строке (нумерация с 0)

    def display_all_books(self):
        try:
            # Создаем новую таблицу с нужными столбцами
            self.tree = ttk.Treeview(self.root, columns=('ID', 'Title', 'Author', 'File Ext', 'File Path', 'File Size', 'Num Pages', 'Metadata'), show='headings')
            self.tree.heading('ID', text='ID')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('File Ext', text='Расширение файла')
            self.tree.heading('File Path', text='Путь к файлу')
            self.tree.heading('File Size', text='Размер файла')
            self.tree.heading('Num Pages', text='Кол-во страниц')
            self.tree.heading('Metadata', text='Метаданные')
            self.tree.grid(row=1, column=0, columnspan=8, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[4]  # предполагается, что путь к файлу хранится в 5-м столбце
                os.startfile(file_path)

            def show_preview(event):
                item = self.tree.identify('item', event.x, event.y)
                preview_bytes = self.tree.item(item, "values")[8]  # Предполагается, что превью книги хранится в 9-м столбце
                print(preview_bytes)
                #image = Image.open(BytesIO(preview_bytes))

                # Создаем новое окно и добавляем в него изображение
                #plt.figure()
                #plt.imshow(image)
                #plt.axis('off')
                #plt.show()

            self.tree.bind('<Double-1>', open_file)
            self.tree.bind('<Double-3>', show_preview)  # Правый клик для предварительного просмотра

            books = self.analyzer.get_all_books()

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def process_directory(self):
        try:
            directory = filedialog.askdirectory()
            file_types = simpledialog.askstring("Ввод", "Введите типы файлов (через запятую):")
            file_types = [file_type.strip() for file_type in file_types.split(',')]
            exclude = simpledialog.askstring("Ввод", "Введите директории для исключения (через запятую):")
            exclude = [dir_.strip() for dir_ in exclude.split(',')]
            max_depth = simpledialog.askinteger("Ввод", "Введите максимальную глубину:")
            convert_odt_to_pdf = simpledialog.askstring("Ввод", "Конвертировать ODT в PDF? (д/н):").lower() == 'д'
            convert_docx_to_pdf = simpledialog.askstring("Ввод", "Конвертировать DOCX в PDF? (д/н):").lower() == 'д'

            self.analyzer.process_directory(directory, file_types=file_types, exclude=exclude,
                                            max_depth=max_depth, convert_odt_to_pdf=convert_odt_to_pdf,
                                            convert_docx_to_pdf=convert_docx_to_pdf)
            messagebox.showinfo("Успех", "Директория обработана успешно")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            
    def reset_database(self):
        try:
            reset = simpledialog.askstring("Ввод", "Перезапустить базу данных? (д/н):").lower() == 'д'
            convert_odt_to_pdf = simpledialog.askstring("Ввод", "Конвертировать ODT в PDF? (д/н):").lower() == 'д'
            convert_docx_to_pdf = simpledialog.askstring("Ввод", "Конвертировать DOCX в PDF? (д/н):").lower() == 'д'

            self.analyzer = BookAnalyzer('books.db', reset=reset, convert_odt_to_pdf=convert_odt_to_pdf, convert_docx_to_pdf=convert_docx_to_pdf)
            messagebox.showinfo("Успех", "База данных перезапущена успешно")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def search_books_by_title(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Title', 'Author', 'Num Pages', 'Path'), show='headings')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('Num Pages', text='Кол-во страниц')
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[3]  # предполагается, что путь к файлу хранится в 4-м столбце
                os.startfile(file_path)

            self.tree.bind('<Double-1>', open_file)


            title = simpledialog.askstring("Ввод", "Введите название книги:")
            books = self.analyzer.search_books_by_title(title)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def search_books_by_author(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Title', 'Author', 'Num Pages', 'Path'), show='headings')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('Num Pages', text='Кол-во страниц')
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[3]  # предполагается, что путь к файлу хранится в 4-м столбце
                os.startfile(file_path)

            self.tree.bind('<Double-1>', open_file)


            author = simpledialog.askstring("Ввод", "Введите имя автора:")
            books = self.analyzer.search_books_by_author(author)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def search_books_by_extension(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Title', 'Author', 'Extension', 'Path'), show='headings')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('Extension', text='Расширение')
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[3]  # предполагается, что путь к файлу хранится в 4-м столбце
                os.startfile(file_path)

            self.tree.bind('<Double-1>', open_file)

            extension = simpledialog.askstring("Ввод", "Введите расширение файла:")
            books = self.analyzer.search_books_by_extension(extension)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
    
    def display_largest_books(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Rank', 'Title', 'Author', 'Size', 'Path'), show='headings')
            self.tree.column('Rank', width=50)  # Здесь мы задаём ширину столбца 'Rank'
            self.tree.heading('Rank', text='Ранг')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('Size', text='Размер')
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=5, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[4]  # предполагается, что путь к файлу хранится в 5-м столбце
                os.startfile(file_path)

            self.tree.bind('<Double-1>', open_file)

            limit = simpledialog.askstring("Ввод", "Введите сколько файлов хотите выбрать")
            offset = simpledialog.askstring("Ввод", "Начать с ? (начинается с 1)")
            offset = int(offset) - 1
            books = self.analyzer.get_largest_books(int(limit), offset)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for i, book in enumerate(books, start=1 + int(offset)):
                self.tree.insert('', 'end', values=(i, *book))

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    
    def display_books_with_most_pages(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Rank', 'Title', 'Author', 'Num Pages', 'Path'), show='headings')
            self.tree.column('Rank', width=50)  # Задаем ширину столбца 'Rank'
            self.tree.heading('Rank', text='Ранг')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('Num Pages', text='Кол-во страниц')
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=5, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[4]  # Теперь путь к файлу хранится в 5-м столбце
                os.startfile(file_path)

            self.tree.bind('<Double-1>', open_file)

            limit = simpledialog.askstring("Ввод", "Введите сколько файлов хотите выбрать")
            offset = simpledialog.askstring("Ввод", "Начать с ? (начинается с 1)")
            offset = int(offset) - 1
            books = self.analyzer.get_books_with_most_pages(int(limit), offset)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for i, book in enumerate(books, start=1 + int(offset)):
                self.tree.insert('', 'end', values=(i, *book))

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


    def display_recently_added_books(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Rank', 'Title', 'Author', 'Path'), show='headings')
            self.tree.column('Rank', width=50)  # Задаем ширину столбца 'Rank'
            self.tree.heading('Rank', text='Ранг')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Author', text='Автор')
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[3]  # Теперь путь к файлу хранится в 4-м столбце
                os.startfile(file_path)

            self.tree.bind('<Double-1>', open_file)

            limit = simpledialog.askstring("Ввод", "Введите сколько файлов хотите выбрать")
            offset = simpledialog.askstring("Ввод", "Начать с ? (начинается с 1)")
            offset = int(offset) - 1
            books = self.analyzer.get_recently_added_books(int(limit), offset)

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for i, book in enumerate(books, start=1 + int(offset)):
                self.tree.insert('', 'end', values=(i, *book))

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Создаем новую функцию для отображения книг без автора
    def display_books_without_author(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Title', 'Num Pages', 'Path'), show='headings')
            self.tree.heading('Title', text='Название')
            self.tree.heading('Num Pages', text='Кол-во страниц')
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=3, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[2]
                os.startfile(file_path)

            self.tree.bind('<Double-1>', open_file)

            books = self.analyzer.get_books_without_author()

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Создаем новую функцию для отображения книг без метаданных
    def display_books_without_metadata(self):
        try:
            self.tree = ttk.Treeview(self.root, columns=('Title', 'File Ext', 'File Size', 'Path'), show='headings')
            self.tree.heading('Title', text='Название')
            self.tree.heading('File Ext', text='Расширение файла')
            self.tree.heading('File Size', text='Размер файла')
            self.tree.heading('Path', text='Путь к файлу')
            self.tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

            def open_file(event):
                item = self.tree.identify('item', event.x, event.y)
                file_path = self.tree.item(item, "values")[3]
                os.startfile(file_path)

            self.tree.bind('<Double-1>', open_file)

            books = self.analyzer.get_books_without_metadata()

            # Очищаем таблицу
            for i in self.tree.get_children():
                self.tree.delete(i)

            # Вставляем новые данные
            for book in books:
                self.tree.insert('', 'end', values=book)

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def display_file_extension_statistics(self):
        try:
            # Получить статистику
            stats = self.analyzer.get_file_extension_statistics()

            # Если статистика пуста, выводим сообщение об ошибке
            if not stats:
                messagebox.showerror("Ошибка", "Не удалось получить статистику расширения файлов.")
                return

            # Создаем список для названий и значений
            labels = []
            values = []

            # Заполняем списки
            for ext, count in stats:
                labels.append(ext)
                values.append(count)

            # Создаем фигуру и оси
            fig, ax = plt.subplots()

            # Создаем круговую диаграмму
            ax.pie(values, labels=labels, autopct='%1.1f%%')

            # Добавляем название
            ax.set_title("Статистика расширений файлов")

            # Показываем диаграмму
            plt.show()

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
    
    def display_books_pages_chart(self):
        try:
            self.analyzer.plot_books_pages()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


root = tk.Tk()
analyzer = BookAnalyzer('books.db')  # создайте свой экземпляр анализатора здесь
app = App(root, analyzer)
root.mainloop()
