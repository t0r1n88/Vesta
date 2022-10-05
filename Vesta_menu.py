from tkinter import *

# Создаем объект окна
window = Tk()
# задаем название окна
window.title('Веста')
# Определяем размеры окна и его расположение считая от левого верхнего угла
window.geometry("400x500+600+200")
# Запрешаем пользователю менять размеры окна программы
window.resizable(False,False)



window.mainloop()
