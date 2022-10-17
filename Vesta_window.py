from tkinter import *
from child_window import ChildWindow

class Window:
    def __init__(self, width, height, title='Vesta', resizable=(False, False),icon=None):
        # Создаем экземпляр класса
        self.root = Tk()
        # Прописываем атрибуты
        self.root.title(title)  # Название окна
        self.root.geometry(f'{width}x{height}+600+200') # Размер окна и сдвиг относительно левого верхнего угла
        self.root.resizable(resizable[0],resizable[1]) # Возможность изменения размера окна
        if icon:
            self.root.iconbitmap(icon) # иконка

    def run(self):
        self.root.mainloop()

    def create_child(self,width, height, title='Vesta', resizable=(False, False),icon=None):
        """
        :param width: Ширина окна
        :param height: Высота окна
        :param title: Заголовк окна
        :param resizable: Можно ли изменять окно
        :param icon: Иконка окна
        :return: экземпляр дочернего окна
        """
        ChildWindow(self.root,width,height,title,resizable,icon)




if __name__ == "__main__":
    window = Window(400,500,icon='favicon.ico')
    window.create_child(200,100,title='Дочернее окно')
    window.run()