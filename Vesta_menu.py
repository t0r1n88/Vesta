

from tkinter import *

# Основано на серии роликов https://www.youtube.com/watch?v=grnchooO2wU&list=PLjRuaCofWO0MVr1xkRXiHR3OAZ4M0C2nW&index=4
# Создаем объект окна
window = Tk()
# задаем название окна
window.title('Веста')
# Определяем размеры окна и его расположение считая от левого верхнего угла
window.geometry("400x500+600+200")
# Запрешаем пользователю менять размеры окна программы
window.resizable(False,False)



window.mainloop()
