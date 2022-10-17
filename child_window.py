from tkinter import *

# класс для дочернего окна
class ChildWindow():
    def __init__(self,parent, width, height, title='Vesta', resizable=(False, False),icon=None):
        # Создаем дочернее окно указывая что родительским явлется окно переданное в аргументе parent
        self.root = Toplevel(parent)
        # Прописываем атрибуты
        self.root.title(title)  # Название окна
        self.root.geometry(f'{width}x{height}+600+200') # Размер окна и сдвиг относительно левого верхнего угла
        self.root.resizable(resizable[0],resizable[1]) # Возможность изменения размера окна
        if icon:
            self.root.iconbitmap(icon) # иконка

        # Сразу же как только будет создаваться дочернее окно, оно будет забирать фокусировку  на себя
        self.grab_focus()

    def grab_focus(self):
        """
        Функция для перехвата событий в дочернее окно, т.е. чтобы дочернее окно было в фокусе а главным окном нельзя было
        пользоваться пока не произойдет определенное событие
        """
        # Перехватываем все события
        self.root.grab_set()
        # Захватываем фокус
        self.root.focus_set()
        # Ждем пока дочернее окно не будет закрыто
        self.root.wait_window()



