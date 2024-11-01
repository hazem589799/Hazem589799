from tkinter import *
from time import sleep
import os
import sys
import shutil
import pandas as pd
from tkinter import ttk
from tkinter import messagebox
import subprocess
from time import *
import subprocess
import requests
import pyttsx3

engine = pyttsx3.init()

engine.setProperty('rate', 150)

engine.setProperty('volume', 1)
pa = os.getcwd()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
text=f"hello {pa}, how are you? , nice to see you again!"
engine.say(text)
engine.runAndWait()



try:
    url = "https://github.com/hazem589799/Hazem589799/archive/refs/heads/main.zip"
    local_zip_path = "Hazem589799.zip"
    new_program_path = "Hazem589799/hazem.py"
                
    response = requests.get(url)
    with open(local_zip_path, 'wb') as file:
        file.write(response.content)
                
    with zipfile.ZipFile(local_zip_path, 'r') as zip_ref:
                    zip_ref.extractall()
                
    subprocess.call([
                    'pyinstaller', '--onefile', '--distpath', '.', '--workpath', '.', '--specpath', '.', new_program_path
                ])
                
    exe_file = "HA.exe"
    original_path = os.path.abspath(__file__)
    original_dir = os.path.dirname(original_path)
    new_exe_path = os.path.join(original_dir, exe_file)
                
    if os.path.isfile(new_exe_path):
        os.remove(new_exe_path)
            
    shutil.move(os.path.join('.', exe_file), new_exe_path)
    os.remove(local_zip_path)
    shutil.rmtree('Hazem589799-main')
    subprocess.call([new_exe_path])
    wd.destroy()
    
except Exception as e:
    messagebox.showerror("خطأ", f"لسه مفيش تحديثات اترفعت على الانترنت  لا تقلق سيتم رفعها فى اسرع وقت: {e}")
hd = Tk()
hd.geometry('200x650+1+1')
hd.configure(bg = 'red')
label1 =Label(hd , text="برنامج المتابعات شركة زانوسى" ,fg='white' ,bg='red'  , font=("tajawal" , 13 , "bold"))
label1.pack(pady=5)
def file_select():
    global file
    file = filedialog.askopenfilename()
button = Button(hd, text="اختر ملف", command=file_select)
button.pack(padx=10, pady=10)
def daman_with_mail():
    row1 = pd.read_excel(rf"{file}")
    for index, ro in row1.iterrows():
        name = ro['name']
        number = ro['phone']
        device = ro['device']
        message = f"""
مساء الخير {name} مع حضرتك متابعة من شركة زانوسي الف مبروك علي المنتج و نتمني أن نكون عند حسن ظن سيادتكم كما أنه يوجد الكثير من الاكسسوار لل{device} ممكن الاستفاده منها بخصم ١٠٪ لحامل هذه الرساله للتوضيح يرجي الاتصال علي رقم المتابعة او الواتس اب
        """
        try:
            sendwhatmsg_instantly(f"+{number}", message , tab_close=True)
        except Exception as e:
            messagebox.showerror(f"حصل مشكلة في إرسال الرسالة لـ {name}: {str(e)}")
def adress_with_mail():
    row1 = pd.read_excel(rf"{file}")
    for index, ro in row1.iterrows():
        name = ro['name']
        number = ro['phone']
        device = ro['device']
        message = f"""
مساء الخير {name}  مع حضرتك متابعة شركة زانوسي ممكن ارقام أخري للتواصل لعدم القدرة علي الوصول للعنوان والارقام غير متاحه أو لا تجمع        """
        try:
            sendwhatmsg_instantly(f"+{number}", message , tab_close=True)
        except Exception as e:
            messagebox.showerror(f"حصل مشكلة في إرسال الرسالة لـ {name}: {str(e)}")

def siana_with_mail():
    row1 = pd.read_excel(rf"{file}")
    for index, ro in row1.iterrows():
        name = ro['name']
        number = ro['phone']
        device = ro['device']
        message = f"""
مساء الخير {name} مع حضرتك قسم المتابعة من شركة زانوسي يرجي توضيح مدي رضاءك عن الخدمة المقدمة ( سيئه - مقبولة - جيدة - ممتازة )
        """
        try:
            sendwhatmsg_instantly(f"+{number}", message , tab_close=True)
        except Exception as e:
            messagebox.showerror(f"حصل مشكلة في إرسال الرسالة لـ {name}: {str(e)}")


button1 = Button(hd , text="متابعات الضمانات "  , command=daman_with_mail)
button1.pack(pady=10)
button1 = Button(hd , text="متابعات عنوان غير واضح "  , command=adress_with_mail)
button1.pack(pady=10)
button2 = Button(hd , text="متابعات الصيانة"  , command=adress_with_mail)
button2.pack(pady=10)
hd.mainloop()
