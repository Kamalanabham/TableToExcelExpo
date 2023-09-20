import filecmp
import fileinput
import os
from tkinter import *
from turtle import bgcolor, color
from PIL import Image,ImageTk
from tkinter import filedialog
import PIL
from multiprocessing import Process,Queue,Pipe


class dataholder:
    def __init__(self):
        self.nameoffile=None
        self.img=None
    def setname(self,name):
        self.nameoffile=name
    def getname(self):
        return self.nameoffile

    def clickfun(self):
    
        self.nameoffile = filedialog.askopenfilename(initialdir='Photos',title='Select an image',filetypes=(('png files','*.png'),('all files','*.*')))
        self.img = Image.open(self.nameoffile)
        h=self.img.height
        w=self.img.width
        hval=h//200
        wval=w//200
        hei=h//hval
        wid=w//wval
        resized_image= self.img.resize((hei,wid))
        resized_image.save('resized.png')
        new_image= ImageTk.PhotoImage(resized_image)
        l = Label(root,image=new_image)
        l.grid()
        l.place(relx=0.5, rely=0.35, anchor=CENTER)
        convbutton.configure(state='normal')


    def convertfun(self):
        #com='python convertcla.py '+ self.nameoffile
        #com=r'python "E:\\Kamalanabham\\Google Drive\\B tech Notes\\Projects\\TableToExcelExpo\\convert.py" {}'.format(self.nameoffile)
        com=r'python "E:\\Kamalanabham\\Google Drive\\B tech Notes\\Projects\\TableToExcelExpo\\convert.py {}"'.format(self.nameoffile)
        fname=self.nameoffile.substring()
        os.system("python convert.py "+self.nameoffile)
        #os.system('python E:\\Kamalanabham\\Google Drive\\B tech Notes\\Projects\\TableToExcelExpo\\convert.py {}'.format(self.nameoffile))



if __name__ == '__main__':

    root = Tk()
    root.configure(bg='#537cf5')
    root.attributes('-fullscreen',True)
    ob = dataholder()
    mylabel= Label(root,text="IMPORT THE IMAGE YOU WOULD LIKE TO CONVERT",bg='#537cf5', fg='#000000',font=("Arial", 40))
    mylabel.grid()
    mylabel.place(relx=0.5, rely=0.1,anchor=CENTER)
    importbutton = Button(root,text="IMPORT",command=ob.clickfun,highlightbackground='#ffdd00',bg='#000000',fg='#ffffff',height=3,width=10,font=("Arial", 15),borderwidth=0)
    importbutton.grid()
    importbutton.place(relx=0.3, rely=0.6, anchor=CENTER)
    convbutton = Button(root,text="CONVERT",command=ob.convertfun,highlightbackground='#ffdd00',bg='#000000',fg='#ffffff',height=3,width=10,font=("Arial", 15),state='disabled',borderwidth=0)
    convbutton.grid()
    convbutton.place(relx=0.7, rely=0.6, anchor=CENTER)
    quitbutton = Button(root,text="QUIT",command=root.quit,highlightbackground='#ffdd00',bg='#b01010',fg='#ffffff',height=3,width=10,font=("Arial", 15),borderwidth=0)
    quitbutton.grid()
    quitbutton.place(relx=0.5, rely=0.85, anchor=CENTER)
    root.mainloop()