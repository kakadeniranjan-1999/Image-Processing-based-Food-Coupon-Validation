import datetime
from PIL import Image, ImageTk
import tkinter as tk
import pyzbar.pyzbar as pyzbar
import datetime
import cv2
import os
import ctypes
import pandas as pd
class Application:
    def __init__(self):
        """ Initialize application which uses OpenCV + Tkinter. It displays
            a video stream in a Tkinter window and stores current snapshot on disk """
        self.vs = cv2.VideoCapture(0) # capture video frames, 0 is your default video camera
        self.vs.set(3,500)
        self.vs.set(4,500)
        self.current_image = None  # current image from the camera

        self.root = tk.Tk()  # initialize root window
        self.root.title("MODEL UNITED NATIONS 2020")  # set window title
        self.root.geometry('700x600')
        # self.destructor function gets fired when the window is closed
        self.root.protocol('WM_DELETE_WINDOW', self.destructor)
        self.panel = tk.Label(self.root)  # initialize image panel
        self.panel.place(x=25,y=100)
        self.mesa = ImageTk.PhotoImage(Image.open(r"MESA logo.png"))
        self.mesa_label=tk.Label(self.root,image=self.mesa)
        self.mesa_label.place(x=590,y=0)
        self.vi=ImageTk.PhotoImage(Image.open(r"VI logo.png"))
        self.vi_label=tk.Label(self.root,image=self.vi)
        self.vi_label.place(x=10,y=0)
        self.mun=ImageTk.PhotoImage(Image.open(r"MUN.jpg"))
        self.mun_label=tk.Label(self.root,image=self.mun)
        self.mun_label.place(x=300,y=0)
        # start a self.video_loop that constantly pools the video sensor
        # for the most recently read frame
        self.video_loop()

    def video_loop(self):
        """ Get frame from the video stream and show it in Tkinter """
        ok, frame = self.vs.read()  # read frame from video stream
        if ok:  # frame captured without any errors
            cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)  # convert colors from BGR to RGBA
            self.current_image = Image.fromarray(cv2image)  # convert image for PIL
            imgtk = ImageTk.PhotoImage(image=self.current_image)  # convert image for tkinter
            self.panel.imgtk = imgtk  # anchor imgtk so it does not be deleted by garbage-collector
            self.panel.config(image=imgtk)  # show the image
            self.decodedObjects = pyzbar.decode(frame)
            if bool(self.decodedObjects)!=False:
                self.scanned=self.decodedObjects[0].data.decode('utf-8')
                self.add_to_excel(self.scanned)
        self.root.after(30, self.video_loop)  # call the same function after 30 milliseconds
    def add_to_excel(self,text):
        df=pd.read_excel(r'E:\MESA\MUN2020\Food Coupon System\FOOD.xlsx')
        did=df['Delegate_ID'].to_list()
        try:
            index=did.index(text[12:24])
            if df['Dinner Status'].iloc[index]=='Pending':
                df.loc[(df.Delegate_ID==text[12:24]),'Dinner Status'] = 'Done'
                df.loc[(df.Delegate_ID==text[12:24]),'Time Stamp'] = datetime.datetime.now().time()
                ctypes.windll.user32.MessageBoxW(0,text,'Welcome', 64)
                print(df['Dinner Status'].iloc[index])
                df.to_excel(r'E:\MESA\MUN2020\Food Coupon System\FOOD.xlsx')
            elif df['Dinner Status'].iloc[index]=='Done':
                ctypes.windll.user32.MessageBoxW(0,'Entry Already Exists!!!','ERROR', 64)
        except ValueError as e:
            print(e)

    def destructor(self):
        self.root.destroy()
        self.vs.release()  # release web camera
        cv2.destroyAllWindows()  # it is not mandatory in this application

