from tkinter import *
from child_window import ChildWindow


class Window:
    def __init__(self, width, height, title='Vesta', resizable=(False, False), icon=None):
        # Создаем экземпляр класса
        self.root = Tk()
        # Прописываем атрибуты
        self.root.title(title)  # Название окна
        self.root.geometry(f'{width}x{height}+600+200')  # Размер окна и сдвиг относительно левого верхнего угла
        self.root.resizable(resizable[0], resizable[1])  # Возможность изменения размера окна
        if icon:
            self.root.iconbitmap(icon)  # иконка

        self.label = Label(self.root,text="I am label",bg="green",relief=RIDGE,font='Consolas 15')
        # self.logo_image = PhotoImage(file='logo.png')
        # self.label = Label(self.root,image=self.logo_image)
        # self.label.image = self.logo_image

    def run(self):
        self.draw_widgets() # Сначала прорисовываем виджеты
        self.root.mainloop() # Звпускаем окно

    def draw_widgets(self):
        """
        Метод для отрисовки виджетов в окне с помощью pack
        """
        self.label.pack(anchor=E,padx=100,pady=150) #anchor выбираем где будет отрисован label
        """
        padx -отступ от границы по оси х
        pady- отступ от границы по оси y
        """


    def create_child(self, width, height, title='Vesta', resizable=(False, False), icon=None):
        """
        :param width: Ширина окна
        :param height: Высота окна
        :param title: Заголовк окна
        :param resizable: Можно ли изменять окно
        :param icon: Иконка окна
        :return: экземпляр дочернего окна
        """
        ChildWindow(self.root, width, height, title, resizable, icon)


if __name__ == "__main__":
    window = Window(400, 500)
    # window.create_child(200, 100, title='Дочернее окно')
    window.run()
