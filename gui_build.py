# -*- coding: utf-8 -*-
"""
Created on Wed Jun 10 10:34:06 2020

@author: DELL
"""

import tkinter as tk
import tkinter.messagebox as msg
import tkinter.simpledialog as disp
import openpyxl
from openpyxl import load_workbook
import os

root=tk.Tk() # the main window
root.title('Body Adiposity Index Calculator')


HEIGHT=700
WIDTH=800



def xl_update():
    filename=xl_entry.get()
    if (filename==""):
        print('error')
        msg.showerror('Error','Please enter full file location ')
    else:
        try:
    
            if os.path.isfile(filename):
                wb=openpyxl.load_workbook(filename=filename)
                ws=wb.active
                age_xl=age_entry.get()
                sex_xl=sex_entry.get()
                height_xl=height_entry.get()
                hc_xl=hc_entry.get()
                [bai_val, cat]=bai(float(height_xl),float(hc_xl),str(sex_xl),float(age_xl))
                ws.append([float(age_xl),sex_xl,float(height_xl),float(hc_xl),float(bai_val),cat])
                wb.save(filename)
                msg.showinfo('update',f'Details are updated at {filename}')
            else:
                msg.showerror('Error','file location is not correct ')
        except PermissionError: 
            msg.showerror('Error','Excel file is already in use, please close before update')



def frange(start, stop=None, step=None):
    # if stop and step argument is None set start=0.0 and step = 1.0
    start = float(start)
    if stop == None:
        stop = start + 0.0
        start = 0.0
    if step == None:
        step = 1.0

    # print("start= ", start, "stop= ", stop, "step= ", step)

    count = 0
    while True:
        temp = float(start + count * step)
        if step > 0 and temp >= stop:
            break
        elif step < 0 and temp <= stop:
            break
        yield temp
        count += 1


def bai(height,hc,sex,age):
  
    try:
        
        cal=(hc/(height)**1.5)-18
        cal=round(cal,1)  
        print('bai is : ', cal)
        # if (sex!='M' or 'm' or 'f' or 'F'):
        #     msg.showerror('Error','Please enter correct details')
        # elif (height==''):
        #     msg.showerror('Error','Please enter correct details')            
        # else:
            
        if sex =='M' :
            if age in frange(0,20,0.1):
                analysis=' below 20, BAI is aplicable beyond 20yr age'
            elif age in frange(20,39,0.1):
                if cal <8:
                    analysis='Underweight'
                elif cal in frange(8,21,0.1):
                    analysis='healthy'
                elif cal in frange(21,26,0.1):
                    analysis='Overweight'
                elif cal >26:
                    analysis='Obese'            
            elif age in frange(40,59):
                if cal <11:
                    analysis='Underweight'
                elif cal in frange(11,23,0.1):
                    analysis='healthy'
                elif cal in frange(23,30,0.1):
                    analysis='Overweight'
                elif cal >29:
                    analysis='Obese'
            elif age in frange(60,79,0.1):
                if cal <13:
                    analysis='Underweight'
                elif cal in frange(13,25,0.1):
                    analysis='healthy'
                elif cal in frange(25,32,0.1):
                    analysis='Overweight'
                elif cal >31:
                    analysis='Obese'        
                 
                    
        elif sex=='F':
            
            if age in frange(0,20,0.1):
                analysis='Age is below 20, BAI cannot be calculated'
            elif age in frange(20,39,0.1):
                if cal <21:
                    analysis='Underweight'
                elif cal in frange(21,33,0.1):
                    analysis='healthy'
                elif cal in frange(33,40,0.1):
                    analysis='Overweight'
                elif cal >39:
                    analysis='Obese'            
            elif age in frange(40,59,0.1):
                if cal <23:
                    analysis='Underweight'
                elif cal in frange(23,35,0.1):
                    analysis='healthy'
                elif cal in frange(35,42,0.1):
                    analysis='Overweight'
                elif cal >41:
                    analysis='Obese'
            elif age in frange(60,79,0.1):
                if cal <25:
                    analysis='Underweight'
                elif cal in frange(25,38,0.1):
                    analysis='healthy'
                elif cal in frange(38,44,0.1):
                    analysis='Overweight'
                elif cal >43:
                    analysis='Obese'  
     
      
        elif(sex==''):
            msg.showerror('Error','Please enter correct details')
            
           
        if age in frange(0,20,0.1):
            out_label2['text']='NA'
            out_label3['text']=f'You are {analysis}'
        elif age>20:
            out_label2['text']=str(cal)
            out_label3['text']=f'You are {analysis}'
        # return cal, analysis

        # return age, height,hc,sex,cal,analysis

    except ValueError :
        msg.showerror('Error','Please enter correct details')
    except UnboundLocalError :
        out_label2['text']='NA'
        msg.showerror('Error','Please enter "M" for male and "F" for female ')
        
        
                
            
    


#input frame

canvas=tk.Canvas(root,height=HEIGHT, width=WIDTH,bd='4') # canvas create
canvas.pack() # pack is one way to organize the widget (canvas in this case), other ways are 'place' and 'grid'
input_frame=tk.Frame(root, bg='#4e5c6e') # can place color name or hex code
input_frame.place(relx=0.1,rely=0.1, relwidth=0.8,relheight=0.4) # place is the most convenient way to organize widgets as it organizes with relative dimension of the screen

age_label=tk.Label(input_frame,text='Enter your age',bg='#d6e0b8',font=('',12))
age_label.place(relx=0.1,rely=0.1, relwidth=0.3,relheight=0.1)
age_label2=tk.Label(input_frame,text='year',bg='#d6e0b8',font=('',10))
age_label2.place(relx=0.77,rely=0.1, relwidth=0.1,relheight=0.1)
age_entry=tk.Entry(input_frame,bg='#b8d7e0',justify='center' ,font=('',12))
age_entry.place(relx=0.55, rely=0.1, relwidth=0.2,relheight=0.1)


height_label=tk.Label(input_frame,text='Enter your height',bg='#d6e0b8',font=('',12))
height_label.place(relx=0.1,rely=0.25, relwidth=0.3,relheight=0.1)
height_label2=tk.Label(input_frame,text='meter',bg='#d6e0b8',font=('',10))
height_label2.place(relx=0.77,rely=0.25, relwidth=0.1,relheight=0.1)

height_entry=tk.Entry(input_frame,bg='#b8d7e0',justify='center' ,font=('',12))
height_entry.place(relx=0.55, rely=0.25, relwidth=0.2,relheight=0.1)

hc_label=tk.Label(input_frame,text='Enter your hip circumference',bg='#d6e0b8',font=('',12))
hc_label.place(relx=0.1,rely=0.4, relwidth=0.4,relheight=0.1)
hc_label2=tk.Label(input_frame,text='centimeter',bg='#d6e0b8',font=('',10))
hc_label2.place(relx=0.77,rely=0.4, relwidth=0.1,relheight=0.1)

hc_entry=tk.Entry(input_frame,bg='#b8d7e0',justify='center' ,font=('',12))
hc_entry.place(relx=0.55, rely=0.4, relwidth=0.2,relheight=0.1)

sex_label=tk.Label(input_frame,text='Specify your gender (M/F)',bg='#d6e0b8',font=('',12))
sex_label.place(relx=0.1,rely=0.55, relwidth=0.4,relheight=0.1)
# hc_label2=tk.Label(input_frame,text='centimeter',bg='#d6e0b8',font=('',10))
# hc_label2.place(relx=0.77,rely=0.55, relwidth=0.1,relheight=0.1)

sex_entry=tk.Entry(input_frame,bg='#b8d7e0',justify='center' ,font=('',12))
sex_entry.place(relx=0.55, rely=0.55, relwidth=0.2,relheight=0.1)


button= tk.Button(input_frame,text='Calculate Body Adiposity Index', bg='#f26168',fg='#030303',font=('',18), 
                  command=lambda: bai(float(height_entry.get()),float(hc_entry.get()),str(sex_entry.get()),float(age_entry.get()))) # button create
button.place(relx=0.1,rely=0.7,relwidth=0.77,relheight=0.2) # pack the button into the root window

#output frame
out_frame=tk.Frame(root, bg='#4e5c6e',bd=5) # can place color name or hex code
out_frame.place(relx=0.1,rely=0.52, relwidth=0.8,relheight=0.2)
out_label1=tk.Label(out_frame, text='CALCULATED BAI : ',bg='#4e5c6e', fg='#cff7cb',font=('',12))
out_label1.place(relx=0.05,rely=0.1, relwidth=0.9,relheight=0.2)
out_label2=tk.Label(out_frame,bg='#4e5c6e', fg='#cff7cb',font=('',16))
out_label2.place(anchor='n',relx=0.5,rely=0.3, relwidth=0.4,relheight=0.3)
out_label3=tk.Label(out_frame,bg='#4e5c6e', fg='#cff7cb',font=('',14))
out_label3.place(relx=0.05,rely=0.65, relwidth=0.9,relheight=0.2)


xl_frame=tk.Frame(root, bg='#4e5c6e',bd=5) # can place color name or hex code
xl_frame.place(relx=0.1,rely=0.74, relwidth=0.8,relheight=0.2)
xl_label1=tk.Label(xl_frame, text='Excel file location: ',bg='#cff7cb', fg='#4e5c6e',font=('',10))
xl_label1.place(relx=0,rely=0.1, relwidth=0.2,relheight=0.2)
xl_entry=tk.Entry(xl_frame,bg='#b8d7e0',justify='center' ,font=('',12))
xl_entry.place(relx=0.22,rely=0.1,relwidth=0.77,relheight=0.2)
save_button= tk.Button(xl_frame,text='Save Result', bg='#f26168',fg='#030303',font=('',18), command=xl_update)
save_button.place(relx=0.15,rely=0.5,relwidth=0.7,relheight=0.4)





root.mainloop() # get into the main window screen










