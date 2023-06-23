import os
import tkinter as tk

from presentation_converter import PresentationConverter
from presentation_converter2 import PresentationConverter2

class StartMenu:
    def __init__(self, root):
        self.root = root
        self.background_image = None
        self.image1 = None
        self.image2 = None
        self.create_menu()

    def create_menu(self):

        # Отримати шлях до поточної директорії
        current_directory = os.path.dirname(os.path.abspath(__file__))

        # Фонове зображення
        background_image_path = os.path.join(current_directory, "background_image.png")
        self.background_image = tk.PhotoImage(file=background_image_path)
        background_label = tk.Label(self.root, image=self.background_image)
        background_label.place(x=0, y=0, relwidth=1, relheight=1)

        # Завантаження та розміщення першої картинки
        image1_path = os.path.join(current_directory, "image1.png")
        self.image1 = tk.PhotoImage(file=image1_path)
        image_label1 = tk.Label(self.root, image=self.image1)
        image_label1.place(x=50, y=50)

        # Завантаження та розміщення другої картинки
        image2_path = os.path.join(current_directory, "image2.png")
        self.image2 = tk.PhotoImage(file=image2_path)
        image_label2 = tk.Label(self.root, image=self.image2)
        image_label2.place(x=450, y=50)

        # Кнопка 1 для генерування свідоцтва
        button1 = tk.Button(self.root, text="Свідоцтво", font=("Arial", 14, "bold"), bg="#007f7c",
                            fg="white", command=self.open_certificate)
        button1.place(x=100, y=300)

        # Кнопка 2 для генерування додатку до свідоцтва
        button2 = tk.Button(self.root, text="Додаток до свідоцтва", font=("Arial", 14, "bold"), bg="#007f7c",
                            fg="white", command=self.open_appendix_to_the_certificate)
        button2.place(x=450, y=300)

    def open_certificate(self):
        self.root.destroy()
        converter = PresentationConverter()
        converter.create_gui()

    def open_appendix_to_the_certificate(self):
        self.root.destroy()
        converter2 = PresentationConverter2()
        converter2.create_gui()


def main():
    root = tk.Tk()
    root.geometry("800x450")
    root.resizable(width=False, height=False)
    StartMenu(root)  # Створюємо об'єкт класу StartMenu для головного меню
    root.mainloop()


if __name__ == "__main__":
    main()
