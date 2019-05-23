from tkinter import *  
from tkinter import ttk
from tkinter import messagebox
import xlrd
import xlsxwriter
import tkinter.filedialog
#from xlutils.copy import copy

root=Tk()
root.geometry('440x440+500+200')




frame0=Frame(height=65,width=400)
frame00=Frame(height=65,width=400)
frame1=Frame(height=75,width=400)
frame2=Frame(height=40,width=400)
frame3=Frame(height=120,width=400)
frame4=Frame(height=120,width=400)
frame5=Frame(height=120,width=400)
frame6=Frame(height=120,width=400)
frame7=Frame(height=120,width=400)
frame8=Frame(height=120,width=400)
frame9=Frame(height=120,width=400)
frame10=Frame(height=90,width=400)
frame100=Frame(height=140,width=400)
frame11=Frame(height=70,width=400)
frame110=Frame(height=140,width=400)
frame111=Frame(height=40,width=400)
frame12=Frame(height=120,width=400)
frame13=Frame(height=120,width=400)
frame14=Frame(height=290,width=400)
frame140=Frame(height=120,width=400)
frame15=Frame(height=120,width=400)
frame16=Frame(height=90,width=400)
frame160=Frame(height=120,width=400)
frame17=Frame(height=50,width=400)
frame170=Frame(height=120,width=400)
frame18=Frame(height=140,width=400)
frame180=Frame(height=120,width=400)
frame19=Frame(height=70,width=400)
frame190=Frame(height=120,width=400)
frame20=Frame(height=90,width=400)
frame200=Frame(height=120,width=400)
frame21=Frame(height=70,width=400)
frame210=Frame(height=120,width=400)
frame22=Frame(height=90,width=400)
frame220=Frame(height=120,width=400)
frame23=Frame(height=50,width=400)
frame230=Frame(height=50,width=400)
frame24=Frame(height=75,width=400)
frame1.grid_propagate(0)
frame2.grid_propagate(0)
frame3.grid_propagate(0)
frame4.grid_propagate(0)
frame5.grid_propagate(0)
frame6.grid_propagate(0)
frame7.grid_propagate(0)
frame8.grid_propagate(0)
frame9.grid_propagate(0)
frame10.grid_propagate(0)
frame100.grid_propagate(0)
frame11.grid_propagate(0)
frame12.grid_propagate(0)
frame13.grid_propagate(0)
frame14.grid_propagate(0)
frame140.grid_propagate(0)
frame15.grid_propagate(0)
frame16.grid_propagate(0)
frame160.grid_propagate(0)
frame17.grid_propagate(0)
frame170.grid_propagate(0)
frame18.grid_propagate(0)
frame180.grid_propagate(0)
frame19.grid_propagate(0)
frame190.grid_propagate(0)
frame20.grid_propagate(0)
frame200.grid_propagate(0)
frame21.grid_propagate(0)
frame210.grid_propagate(0)
frame22.grid_propagate(0)
frame220.grid_propagate(0)
frame23.grid_propagate(0)
frame230.grid_propagate(0)
count=0
v=IntVar()
def pre():
    global count
    for i in range(len(list_ans)):#if there is no value in Entry, make it back to 0
        try:
            if i != 0 or 2 or 1:
                list_ans[i].get()!=''
        
        except TclError:
            list_ans[i].set(00)
    
    count-=1
    if count<1:
        count=1
    if count==1:
        var.set('1. In which country is your farm located?')
        frame5.grid_forget()
        frame3.grid(sticky=W)
    if count==2:
        var.set('2. What crop do you produce?')
        
        frame4.grid_forget()
        frame5.grid(sticky=W)
    if count==3:
        var.set('3. What type of farm do you have (choose the one \nclosest to the options)?')
        frame6.grid_forget()
        frame4.grid(sticky=W)
    if count==4:
        var.set('4. How many square meters do you use for ' + ans3.get() +'?\n(all the following questions should be \nanswered based on this area)')
        frame7.grid_forget()
        frame6.grid(sticky=W)
    if count==5:
        var.set('5. How many kilograms of ' +ans3.get()+ ' do you produce per year?')
        frame8.grid_forget()
        frame7.grid(sticky=W)
    if count==6:
        var.set('6. How much renewable and non-renewable electricity \ndo you buy per year? \n (own production not included)')
        frame9.grid_forget()
        frame8.grid(sticky=W)
    if count==7:
        var.set('7. Do you produce your own renewable energy and how much \ndo you produce?')
        frame10.grid_forget()
        frame100.grid_forget()
        frame9.grid(sticky=W)
    if count==8:
        var.set("8. Can you specify what and how much the electricity is spend on? \nif you can't fill in zeros")
        frame11.grid_forget()
        frame110.grid_forget()
        frame111.grid_forget()
        frame10.grid(sticky=W)
        frame100.grid(sticky=W)
    if count==9:
        var.set("9. Do you use any fossil fuels(excluding transportation), \nand how much do you use(if you don't know fill in zero)")
        frame14.grid_forget()
        frame140.grid_forget()
        frame11.grid(sticky=W)
        frame111.grid(sticky=W)
        frame110.grid(sticky=W)

    if count==10:
        var.set('10. How much kilograms do you use of the following \nNPK chemicals per year?')
        frame16.grid_forget()
        frame160.grid_forget()
        frame14.grid(sticky=W)
        frame140.grid(sticky=W)

    if count==11:
        var.set('11. Do you use substrate and how much per year? (kg)')
        frame17.grid_forget()
        frame170.grid_forget()
        frame16.grid(sticky=W)
        frame160.grid(sticky=W)
    if count==12:
        var.set('12. How much water do you use?')
        frame18.grid_forget()
        frame180.grid_forget()
        frame17.grid(sticky=W)
        frame170.grid(sticky=W)
    if count==13:
        var.set('13. How much pesticides do you use? ')
        frame19.grid_forget()
        frame190.grid_forget()
        frame18.grid(sticky=W)
        frame180.grid(sticky=W)
    if count==14:
        var.set('14. Is the product sold to the customer packaged? ')
        frame20.grid_forget()
        frame200.grid_forget()
        frame19.grid(sticky=W)
        frame190.grid(sticky=W)
    if count==15:
        var.set('15. How much green waste do you produce? ')
        frame22.grid_forget()
        frame220.grid_forget()
        frame20.grid(sticky=W)
        frame200.grid(sticky=W)

    if count==16:
        var.set('16. How far does your product travel to the distribution center \non average? ')
        frame23.grid_forget()
        frame230.grid_forget()
        frame22.grid(sticky=W)
        frame220.grid(sticky=W)
    
    

def next1():
    global count
    global v

    for i in range(len(list_ans)):#if there is no value in Entry, make it back to 0
        try:
            if i != 0 or 2 or 1:
                list_ans[i].get()!=''
                
        except TclError:
            list_ans[i].set(00)
        
    count+=1
    if count==2:
        
        var.set('2. What crop do you produce?')
        frame3.grid_forget()
        frame5.grid(sticky=W)
    if count==3:
        var.set('3. What type of farm do you have (choose the one \nclosest to the options)?')
        frame5.grid_forget()
        frame4.grid(sticky=W)
    if count==4:
        var.set('4. How many square meters do you use for ' + ans3.get() +'?\n(all the following questions should be \nanswered based on this area)')
        frame4.grid_forget()
        frame6.grid(sticky=W)
    
    if count==5:
        var.set('5. How many kilograms of ' +ans3.get()+ ' do you produce per year?')
        frame6.grid_forget()
        frame7.grid(sticky=W)
    if count==6:
        var.set('6. How much renewable and non-renewable electricity \ndo you buy per year? \n (own production not included)')
        frame7.grid_forget()
        frame8.grid(sticky=W)
    if count==7:
        var.set('7. Do you produce your own renewable energy and how much \ndo you produce?')
        frame8.grid_forget()
        frame9.grid(sticky=W)
    if count==8:
        var.set("8. Can you specify what and how much the electricity is spend on? \nif you can't fill in zeros")
        frame9.grid_forget()
        frame10.grid(sticky=W)
        frame100.grid(sticky=W)
    if count==9:
        var.set("9. Do you use any fossil fuels(excluding transportation), \nand how much do you use (if you don't know fill in zero)")
        frame10.grid_forget()
        frame100.grid_forget()
        frame11.grid(sticky=W)
        frame111.grid(sticky=W)
        frame110.grid(sticky=W)

    if count==10:
        var.set('10. How much kilograms do you use of the following \nNPK chemicals per year?')
        other = 100 - ans916.get() - ans915.get() - ans914.get() - ans913.get() - ans912.get() - ans911.get() - ans910.get() - ans99.get() - ans98.get() - ans97.get()
        ans917.set(other)
        frame11.grid_forget()
        frame111.grid_forget()
        frame110.grid_forget()
        frame14.grid(sticky=W)
        frame140.grid(sticky=W)

    if count==11:
        var.set('11. Do you use substrate and how much per year? (kg)?')
        frame14.grid_forget()
        frame140.grid_forget()
        frame16.grid(sticky=W)
        frame160.grid(sticky=W)
    if count==12:
        var.set('12. How much water do you use?')
        frame16.grid_forget()
        frame160.grid_forget()
        frame17.grid(sticky=W)
        frame170.grid(sticky=W)
    if count==13:
        var.set('13. How much pesticides do you use?')
        frame17.grid_forget()
        frame170.grid_forget()
        frame18.grid(sticky=W)
        frame180.grid(sticky=W)
    if count==14:
        var.set('14. Is the product sold to the customer packaged? ')
        frame18.grid_forget()
        frame180.grid_forget()
        frame19.grid(sticky=W)
        frame190.grid(sticky=W)
    if count==15:
        var.set('15. How much waste do you produce? ')
        frame19.grid_forget()
        frame190.grid_forget()
        frame20.grid(sticky=W)
        frame200.grid(sticky=W)
   
    if count==16:
        var.set('16. How far does your produce travel to the distribution center \non average? ')
        frame20.grid_forget()
        frame200.grid_forget()
        frame22.grid(sticky=W)
        frame220.grid(sticky=W)
    if count==17:
        var.set('17. How much of the product does not survive the transport \nstage to the store? ')
        frame22.grid_forget()
        frame220.grid_forget()
        frame23.grid(sticky=W)
        frame230.grid(sticky=W)
    if count==18:
        if ans211.get()<0 or ans211.get()>100:
            messagebox.showinfo('Notification','The range of the number should be 0-100')
            count-=1
        else:
            a=messagebox.askquestion(title='Notification',message='This is the last question,are you finished?')
            if a=='yes':
                frame1.grid_forget()
                frame23.grid_forget()
                frame230.grid_forget()
                frame2.grid_forget()
                frame24.pack(anchor=CENTER)
            if a=='no':
                count-=1
def quit1():
    root.destroy()
def enter(event):
    if count>0:
        next1()

wb=xlrd.open_workbook('Crops energy content.xlsx')
lis=[]
Blad1=wb.sheet_by_name('Blad1')
for i in range(1,25):
    lis.append(Blad1.col_values(2)[i])

v=IntVar()
var=StringVar()
var.set('1. In which country is your farm located?')
helloLabel = Label(frame1,textvariable= var)
helloLabel.grid(row=0,column=0,padx=10,pady=10,sticky=W)
#frame1
button2=Button(frame2,text=('previous'),command=pre)
button2.grid(row=0,column=0,padx=10)
shitlabel=Label(frame2,text='                                   ')
shitlabel.grid(row=0,column=1)
button1=Button(frame2,text=('next'),command=next1)
button1.grid(row=0,column=2,sticky=E)
root.bind('<Return>',enter)
#frame2
ans1=StringVar()
country=ttk.Combobox(frame3,textvariable=ans1,state='readonly')
country['values']=('Netherlands','China','Germany')
country.current(0)
country.grid(padx=10)

#ans1.get()
#frame3,q1
button1=Radiobutton(frame4,text='Greenhouse',variable=v,value=1)
button1.grid(row=1,column=0,padx=10,sticky=W)
button2=Radiobutton(frame4,text='Open field',variable=v,value=2)
button2.grid(row=2,column=0,padx=10,sticky=W)
button3=Radiobutton(frame4,text='Vertical farm',variable=v,value=3)
button3.grid(row=3,column=0,padx=10,sticky=W)
#v
#frame4,q2
ans3=StringVar()
crop=ttk.Combobox(frame5,textvariable=ans3,state='readonly')
crop['values']=lis
crop.current(0)
crop.grid(padx=10)
#ans3.get()   
#frame5,q3
ans4=IntVar()
Area=Entry(frame6,textvariable=ans4)
Area.grid(sticky=W,padx=20,pady=20)
#ans4
#frame6,q4
ans5=IntVar()
Kilogram=Entry(frame7,textvariable=ans5)
Kilogram.grid(sticky=W,padx=20,pady=20)
#ans5
#frame7.q5
ans61=IntVar()
ans62=IntVar()
greenlabel=Label(frame8,text='Renewable (kWh)')
greenlabel.grid(row=1,column=0,padx=20,sticky=W)
greenentry=Entry(frame8,width=10,textvariable=ans61)
greenentry.grid(row=1,column=1)
greylabel=Label(frame8,text='Non-renewable(kWh)')
greylabel.grid(row=2,column=0,padx=20,sticky=W)
greyentry=Entry(frame8,width=10,textvariable=ans62)
greyentry.grid(row=2,column=1)
#ans61,ans62
#frame8,q6
ans71=IntVar()
ans72=IntVar()
ans73=IntVar()
solarlabel=Label(frame9,text='Solar energy (kWh)')
solarlabel.grid(row=1,column=0,padx=20,sticky=W)
solarentry=Entry(frame9,width=10, textvariable =ans71)
solarentry.grid(row=1,column=1)
        
biomasslabel=Label(frame9,text='Biomass (kWh)')
biomasslabel.grid(row=2,column=0,padx=20,sticky=W)
biomassentry=Entry(frame9,width=10, textvariable =ans72)
biomassentry.grid(row=2,column=1)
        
windlabel=Label(frame9,text='Windpower (kWh)')
windlabel.grid(row=3,column=0,padx=20,sticky=W)
windentry=Entry(frame9,width=10, textvariable =ans73)
windentry.grid(row=3,column=1)
#frame9,q7
ans81=IntVar()
ans82=IntVar()
ans83=IntVar()
ans84=IntVar()
ans85=IntVar()
ans86=IntVar()
ans87=IntVar()
ans88=IntVar()
ans89=IntVar()
ans890=IntVar()
heatingl=Label(frame10,text='Heating (kWh)')
heatingl.grid(row=1,column=0,padx=5,sticky=W)
heatingy=Entry(frame10,width=10, textvariable = ans81)
heatingy.grid(row=1,column=1)
        
coolingl=Label(frame10,text='Cooling (kWh)')
coolingl.grid(row=2,column=0,padx=5,sticky=W)
coolingy=Entry(frame10,width=10,  textvariable = ans82)
coolingy.grid(row=2,column=1)

ventillationl=Label(frame10,text='Ventillation (kWh)')
ventillationl.grid(row=3,column=0,padx=5,sticky=W)
ventillationy=Entry(frame10,width=10,  textvariable = ans83)
ventillationy.grid(row=3,column=1)
        
lightingl=Label(frame10,text='Lighting (kWh)')
lightingl.grid(row=1,column=2,padx=5,sticky=W)
lightingy=Entry(frame10,width=10,  textvariable = ans84)
lightingy.grid(row=1,column=3)
        
machineryl=Label(frame10,text='Machinery (kWh)')
machineryl.grid(row=2,column=2,padx=5,sticky=W)
machineryy=Entry(frame10,width=10,  textvariable = ans85)
machineryy.grid(row=2,column=3)
        
storagel=Label(frame10,text='Storage (kWh)')
storagel.grid(row=3,column=2,padx=5,sticky=W)
storagey=Entry(frame10,width=10,  textvariable = ans86)
storagey.grid(row=3,column=3)




Label(frame100, text='Selling renewable(kWh)').grid(row=0, column=0, sticky=W,padx=5)
Entry(frame100,width=10,textvariable=ans87).grid(row=0,column=1)
Label(frame100, text='Selling non-renewable(kWh)').grid(row=1, column=0, sticky=W,padx=5)
Entry(frame100,width=10,textvariable=ans88).grid(row=1,column=1)
Label(frame100, text='Other(kWh)').grid(row=2, column=0, sticky=W,padx=5)
Entry(frame100,width=10,textvariable=ans89).grid(row=2,column=1)
Checkbutton(frame100,text='I don\'t know',variable=ans890).grid(row=3,column=0,sticky=W,padx=5)
#frame10,q8

ans91 = IntVar()
ans92 = IntVar()
ans93 = IntVar()
ans94 = IntVar()
ans95 = IntVar()
ans96 = IntVar()
ans97 = IntVar()
ans98 = IntVar()
ans99 = IntVar()
ans910 = IntVar()
ans911 = IntVar()
ans912 = IntVar()
ans913 = IntVar()
ans914 = IntVar()
ans915 = IntVar()
ans916 = IntVar()
ans917 = IntVar()
other = 100 - ans916.get() - ans915.get() - ans914.get() - ans913.get() - ans912.get() - ans911.get() - ans910.get() - ans99.get() - ans98.get() - ans97.get()
ans917.set(other)
petroll = Label(frame11, text='Petrol (L)')
petroll.grid(row=0, column=0, padx=5, sticky=W)
petroly = Entry(frame11, width=5, textvariable=ans91)
petroly.grid(row=0, column=1)

diesell = Label(frame11, text='Diesel (L)')
diesell.grid(row=1,column=0, padx=5, sticky=W)
diesely = Entry(frame11, width=5, textvariable=ans92)
diesely.grid(row=1, column=1)


Ngasl = Label(frame11, text='Natural gas (L)')
Ngasl.grid(row=0, column=2, padx=10, sticky=W)
Ngasy = Entry(frame11, width=5, textvariable=ans94)
Ngasy.grid(row=0, column=3)

oill = Label(frame11, text='Oil (L)')
oill.grid(row=1, column=2, padx=10, sticky=W)
oily = Entry(frame11, width=5, textvariable=ans95)
oily.grid(row=1, column=3)

dd=2

def cal(event):
    try:
        if 0<=ans916.get()<=100 and 0<=ans915.get()<=100 and 0<=ans914.get()<=100 and 0<=ans913.get()<=100 and 0<=ans912.get()<=100 and 0<=ans911.get()<=100 and 0<=ans910.get()<=100 and 0<=ans99.get()<=100 and 0<=ans98.get()<=100 and 0<=ans97.get()<=100: 
            other = 100 - ans916.get() - ans915.get() - ans914.get() - ans913.get() - ans912.get() - ans911.get() - ans910.get() - ans99.get() - ans98.get() - ans97.get()
            # judge whether the value enter in is in the range (0,100)
            if other>=0:
                ans917.set(other)
            else:
                messagebox.showinfo('Notification','The range of the number should be (0,100)')
        else:
            global dd
            dd+=1
        
            if dd%3==0:
                messagebox.showinfo('Notification','The range of the number should be (0,100)')
            
    except TclError:
        for i in range(len(list_ans)):#if there is no value in Entry, make it back to 0
            try:
                if i != 0 or 2 or 1:
                    list_ans[i].get()!=''
                
            except TclError:
                list_ans[i].set(00)
    
Label(frame111, text='Estimate in percentages what the fossil fuels are used for').grid(row=0, column=0, sticky=W,padx=5)
Label(frame110, text='Heating').grid(row=1, column=0, sticky=W,padx=5)
q = Entry(frame110, width=5, textvariable=ans97)
q.grid(row=1, column=1)
Label(frame110, text='%').grid(row=1, column=2, sticky=W)
q.bind('<FocusOut>', cal)
Label(frame110, text='Cooling').grid(row=2, column=0, sticky=W,padx=5)
w = Entry(frame110, width=5, textvariable=ans98)
w.grid(row=2, column=1)
Label(frame110, text='%').grid(row=2, column=2, sticky=W)
w.bind('<FocusOut>', cal)
Label(frame110, text='Electricity').grid(row=3, column=0, sticky=W,padx=5)
e = Entry(frame110, width=5, textvariable=ans99)
e.grid(row=3, column=1)
Label(frame110, text='%').grid(row=3, column=2, sticky=W)
e.bind('<FocusOut>', cal)
Label(frame110, text='Tillage').grid(row=4, column=0, sticky=W,padx=5)
a = Entry(frame110, width=5, textvariable=ans910)
a.grid(row=4, column=1)
Label(frame110, text='%').grid(row=4, column=2, sticky=W)
a.bind('<FocusOut>', cal)
Label(frame110, text='Sowing').grid(row=5, column=0, sticky=W,padx=5)
s = Entry(frame110, width=5, textvariable=ans911)
s.grid(row=5, column=1)
Label(frame110, text='%').grid(row=5, column=2, sticky=W)
s.bind('<FocusOut>', cal)
Label(frame110, text='Weeding').grid(row=5, column=3, sticky=W,padx=20)
d = Entry(frame110, width=5, textvariable=ans912)
d.grid(row=5, column=4)
Label(frame110, text='%').grid(row=5, column=5, sticky=W)
d.bind('<FocusOut>', cal)
Label(frame110, text='Harvest').grid(row=1, column=3, sticky=W,padx=20)
z = Entry(frame110, width=5, textvariable=ans913)
z.grid(row=1, column=4)
Label(frame110, text='%').grid(row=1, column=5, sticky=W)
z.bind('<FocusOut>', cal)
Label(frame110, text='Fertilize').grid(row=2, column=3, sticky=W,padx=20)
x = Entry(frame110, width=5, textvariable=ans914)
x.grid(row=2, column=4)
Label(frame110, text='%').grid(row=2, column=5, sticky=W)
x.bind('<FocusOut>', cal)
Label(frame110, text='Irrigation').grid(row=3, column=3, sticky=W,padx=20)
c = Entry(frame110, width=5, textvariable=ans915)
c.grid(row=3, column=4)
Label(frame110, text='%').grid(row=3, column=5, sticky=W)
c.bind('<FocusOut>', cal)
Label(frame110, text='Pesticide').grid(row=4, column=3, sticky=W,padx=20)
r = Entry(frame110, width=5, textvariable=ans916)
r.grid(row=4, column=4)
Label(frame110, text='%').grid(row=4, column=5, sticky=W)
r.bind('<FocusOut>', cal)
Label(frame110, text='Other').grid(row=6, column=0, sticky=W,padx=5)
Entry(frame110, width=5, textvariable=ans917).grid(row=6, column=1)
Label(frame110, text='%').grid(row=6, column=2, sticky=W)
#frame11,q9

ans121=IntVar()
ans122=IntVar()
ans123=IntVar()
ans124=IntVar()
ans125=IntVar()
ans126=IntVar()
ans127=IntVar()
ans128=IntVar()
ans129=IntVar()
ans130=IntVar()
ans131=IntVar()
ans132=IntVar()
ans133=IntVar()

Label(frame14,text='Ammoniumnitrate (kg)').grid(row=1,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable =ans121 ).grid(row=1,column=1)

Label(frame14,text='Calciumammoniumnitrate (kg)').grid(row=2,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans122).grid(row=2,column=1)
        
Label(frame14,text='Ammoniumsulphate (kg)').grid(row=3,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans123).grid(row=3,column=1)
        
Label(frame14,text='Triplesuperphosphate (kg)').grid(row=4,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans124).grid(row=4,column=1)
        
Label(frame14,text='Single super phosphate (kg)').grid(row=5,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans125).grid(row=5,column=1)
        
Label(frame14,text='Ammonia(kg)').grid(row=6,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans126).grid(row=6,column=1)

Label(frame14,text='limestone (kg)').grid(row=7,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans127).grid(row=7,column=1)

Label(frame14,text='NPK 15-15-15(kg)').grid(row=8,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans128).grid(row=8,column=1)

Label(frame14,text='Urea(kg)').grid(row=9,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans129).grid(row=9,column=1)

Label(frame14,text='Manure(kg)').grid(row=10,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans130).grid(row=10,column=1)

Label(frame14,text='Phosphoric acid(kg)').grid(row=11,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans131).grid(row=11,column=1)

Label(frame14,text='mono-ammonium phosphate(kg)').grid(row=12,column=0,padx=5,sticky=W)
Entry(frame14,width=10, textvariable = ans132).grid(row=12,column=1)

Checkbutton(frame140,text='I don\'t know',variable=ans133).grid(padx=5)
#frame14,frame140,q10

ans141=IntVar()
ans142=IntVar()
ans143=IntVar()
ans144=IntVar()
ans145=IntVar()
ans146=IntVar()
ans147=IntVar()
Label(frame16,text='Rockwool(kg)').grid(row=1,column=0,padx=5,sticky=W)
Entry(frame16,width=10, textvariable =ans141 ).grid(row=1,column=1)

Label(frame16,text='Perlite(kg)').grid(row=2,column=0,padx=5,sticky=W)
Entry(frame16,width=10, textvariable = ans142).grid(row=2,column=1)

Label(frame16,text='Cocofiber(kg)').grid(row=1,column=2,padx=5,sticky=W)
Entry(frame16,width=10, textvariable = ans143).grid(row=1,column=3)
        
Label(frame16,text='hemp fiber(kg)').grid(row=2,column=2,padx=5,sticky=W)
Entry(frame16,width=10, textvariable = ans144).grid(row=2,column=3)

Label(frame16,text='Peat(kg)').grid(row=3,column=0,padx=5,sticky=W)
Entry(frame16,width=10, textvariable = ans145).grid(row=3,column=1)

Label(frame16,text='Peat Moss Kg)').grid(row=3,column=2,padx=5,sticky=W)
Entry(frame16,width=10, textvariable = ans146).grid(row=3,column=3)



Checkbutton(frame160,text='No substrate is used',variable=ans147).grid(padx=5)
#ans141,ans142,ans143,ans144,ans145
#frame16,frame160,q11
ans151=IntVar()
ans152=IntVar()
Label(frame17,text='Tap water(L)').grid(row=1,column=0,padx=5,sticky=W)
Entry(frame17,width=10, textvariable =ans151 ).grid(row=1,column=1)

Checkbutton(frame170,text='I don\'t know',variable=ans152).grid(sticky=W,padx=5)
#ans151,ans152,ans153,ans154
#frame17,q12
ans161=IntVar()
ans162=IntVar()
ans163=IntVar()
ans164=IntVar()
ans165=IntVar()
ans166=IntVar()
Label(frame18,text='Atrazine(kg)').grid(row=1,column=0,padx=5,sticky=W)
Entry(frame18,width=10, textvariable =ans161 ).grid(row=1,column=1)

Label(frame18,text='Glyphosphate(kg)').grid(row=2,column=0,padx=5,sticky=W)
Entry(frame18,width=10, textvariable = ans162).grid(row=2,column=1)

Label(frame18,text='Metolachlor(kg)').grid(row=3,column=0,padx=5,sticky=W)
Entry(frame18,width=10, textvariable =ans163 ).grid(row=3,column=1)

Label(frame18,text='Herbicide(kg)').grid(row=4,column=0,padx=5,sticky=W)
Entry(frame18,width=10, textvariable = ans164).grid(row=4,column=1)

Label(frame18,text='Insectiside(kg)').grid(row=5,column=0,padx=5,sticky=W)
Entry(frame18,width=10, textvariable = ans165).grid(row=5,column=1)
Checkbutton(frame180,text='I don\'t know',variable=ans166).grid(sticky=W,padx=5)

#frame18,q13
ans171=IntVar()
ans172=IntVar()
ans173=IntVar()
Radiobutton(frame19,text='Yes, it is',variable=ans171,value=1).grid(sticky=W,padx=5)

Radiobutton(frame19,text='No, it isn\'t',variable=ans171,value=2).grid(sticky=W,padx=5)

#frame19.q14
ans181=IntVar()
ans182=IntVar()
ans183=IntVar()
ans184=IntVar()
Label(frame20,text='Green waste(kg)').grid(row=1,column=0,padx=5,sticky=W)
Entry(frame20,width=10, textvariable =ans181 ).grid(row=1,column=1)

Label(frame20,text='Gray waste(kg)').grid(row=2,column=0,padx=5,sticky=W)
Entry(frame20,width=10, textvariable = ans182).grid(row=2,column=1)

Label(frame20,text='Paper(kg)').grid(row=3,column=0,padx=5,sticky=W)
Entry(frame20,width=10, textvariable = ans183).grid(row=3,column=1)

Checkbutton(frame200,text='I don\'t know',variable=ans184).grid(sticky=W,padx=5)

#frame20,q15


ans201=IntVar()
ans202=IntVar()
ans203=IntVar()
ans204=IntVar()
Label(frame22,text='Plane (km)').grid(row=1,column=0,padx=40,sticky=W)
Entry(frame22,width=10, textvariable =ans201 ).grid(row=1,column=1)

Label(frame22,text='Truck (km)').grid(row=2,column=0,padx=40,sticky=W)
Entry(frame22,width=10, textvariable = ans202).grid(row=2,column=1)

Label(frame22,text='Ship (km)').grid(row=3,column=0,padx=40,sticky=W)
Entry(frame22,width=10, textvariable = ans203).grid(row=3,column=1)

Checkbutton(frame220,text='I don\'t know',variable=ans204).grid(sticky=W,padx=40)

#frame22,q16
ans211=IntVar()
ans212=IntVar()
Entry(frame23,width=5,textvariable=ans211).grid(row=0,column=0,padx=20,pady=20)
Label(frame23,text='%').grid(row=0,column=1,sticky=W)
Checkbutton(frame230,text='I don\'t know',variable=ans212).grid(sticky=W,padx=40)

#frame23,q17
Button(frame24,text='Quit',font=12,command=quit1).pack(side=BOTTOM,anchor=CENTER)
Label(frame24,text='\n\n\n\nThanks for your time,\nthe result is in the folder \nsaved as Excel named as \nthe farm name and crop name',font=12).pack(fill=BOTH,side=BOTTOM,anchor=CENTER)

def start():#The command of 'Start' button at the beginning
    frame0.pack_forget()
    frame00.pack(anchor=CENTER)
def next2():#The command of 'Next' button after you input the farm name
    global count
    frame00.pack_forget()
    frame1.grid()
    frame2.grid()
    frame3.grid()
    count+=1
startbutton=Button(frame0,text='Start',command=start,font=12)
startbutton.pack(fill=X,side=BOTTOM,anchor=CENTER)

startlabel=Label(frame0,text='\n\n\n\nQuestionaire\n\n\n',font=12)
startlabel.pack(fill=BOTH,side=BOTTOM)
  

frame0.pack(anchor=CENTER)
farm_name=StringVar()
Button(frame00,text='Next',command=next2).pack(fill=BOTH,side=BOTTOM,anchor=CENTER)
Entry(frame00,textvariable=farm_name).pack(fill=BOTH,side=BOTTOM,anchor=CENTER)
Label(frame00,text='\n\n\n\nEnter the name of your farm').pack(fill=BOTH,side=BOTTOM)
list_ans=[farm_name,ans1,ans3,v,ans4,ans5,ans61,ans62,ans71,ans72,ans73,ans81,ans82,ans83,ans84,ans85,ans86,ans87,ans88,ans89,ans890,ans91,ans92,ans94,ans95,ans97,ans98,ans99,ans910,ans911,ans912,ans913,ans914,ans915,ans916,ans917,ans121,ans122,ans123,ans124,ans125,ans126,ans127,ans128,ans129,ans130,ans131,ans132,ans133,ans141,ans142,ans143,ans144,ans145,ans146,ans147,ans151,ans152,ans161,ans162,ans163,ans164,ans165,ans166,ans171,ans172,ans181,ans182,ans183,ans184,ans201,ans202,ans203,ans204,ans211,ans212]
#make a list contain all the variables
path1=StringVar()
path2=StringVar()

root.file_opt = options = {}# define the 'file' Menu
options['defaultextension'] = '.txt'  
options['filetypes'] = [('all files', '.*'), ('text files', '.txt')]  

options['initialfile'] = 'myfile.txt'  
options['parent'] = root  
options['title'] = 'This is a title'    
def file_open():
    try:
        path1=tkinter.filedialog.askopenfilename()
        f=open(path1)
        lines=f.readlines()
        num=-1
        for line in lines:
            num+=1
            if num<3:
                list_ans[num].set(str(line.strip('\n')))            
            if num>=3:
                list_ans[num].set(int(line.strip('\n')))
    except:
        pass
            
def file_save():
    try:
        path2=tkinter.filedialog.asksaveasfilename(**root.file_opt)
        f1=open(path2,'w')
        for i in range(len(list_ans)):
            f1.write(str(list_ans[i].get())+'\n')
        f1.close()
    except:
        pass
menu=Menu(root)
filemenu=Menu(menu,tearoff=0)
filemenu.add_command(label='Load',command=file_open)
filemenu.add_command(label='Save',command=file_save)
filemenu.add_command(label='Quit',command=root.quit)
menu.add_cascade(label='File',menu=filemenu)
root.config(menu=menu)


root.mainloop()



if count>=17:# when count >=17 , it means the questionnaire is finished, start to read database
    workbook = xlrd.open_workbook('Country database.xlsx') 
    sheet = workbook.sheet_by_name(ans1.get())
    Eoc=Blad1.col_values(4)[lis.index(ans3.get())+1]#Energy of crop
    non_count=str()
    if ans890.get()==1: #if choose 'I don‘t know’ option， make the value back to zero
        ans81.set(0);ans82.set(0);ans83.set(0);ans84.set(0);ans8.set(0);ans86.set(0);ans87.set(0);ans88.set(0);ans89.set(0)
        non_count=('Specification of electricity,')
    if ans133.get()==1:
        ans121.set(0);ans122.set(0);ans123.set(0);ans124.set(0);ans125.set(0);ans126.set(0);ans127.set(0);ans128.set(0);ans129.set(0);ans130.set(0);ans131.set(0);ans132.set(0)
        non_count=(non_count+'NPK chemicals,')
    if ans147.get()==1:
        ans141.set(0);ans142.set(0);ans143.set(0);ans144.set(0);ans145.set(0);ans146.set(0)
        non_count=(non_count+'Substrate,')
    if ans152.get()==1:
        ans151.set(0)
        non_count=(non_count+'Water,')
    if ans166.get()==1:
        ans161.set(0);ans162.set(0);ans163.set(0);ans164.set(0);ans165.set(0)
        non_count=(non_count+'Pesticides,')
    if ans184.get()==1:
        ans181.set(0);ans182.set(0);ans183.set(0)
        non_count=(non_count+'Waste,')
    if ans204.get()==1:
        ans201.set(0);ans202.set(0);ans203.set(0)
        non_count=(non_count+'NPK chemicals,')
    if ans212.get()==1:
        ans211.set(0)
        non_count=(non_count+'Waste during transportation')
   
    #electricity
    C1=float(sheet.cell_value(1, 1)) #co2 equivalent green electricty from net
    C2=float(sheet.cell_value(1, 2)) #energy equivalent green electricty from net
    C3=float(sheet.cell_value(2, 1)) #co2 equivalent gray electricity from net
    C4=float(sheet.cell_value(2, 2)) #energy equivalent gray electricity from net
    C5=float(sheet.cell_value(4, 1)) #Co2 equivalent solar electricity
    C6=float(sheet.cell_value(4, 2)) #energy equivalent solar electricity
    C7=float(sheet.cell_value(5, 1)) #Co2 equivalent wind electricity
    C8=float(sheet.cell_value(5, 2)) #energy equivalent wind electricity
    C9=float(sheet.cell_value(6, 1)) #Co2 equivalent biomass electricity
    C10=float(sheet.cell_value(6, 2))#energy equivalent buimass electricity

    #fuels
    Fo1=sheet.cell_value(1, 5)  #petrol Co2 equivalent/L
    Fo2=sheet.cell_value(1, 6)  #petrol Energy equivalent/L
    Fo3=sheet.cell_value(2, 5)  #diesel Co2 equivalent/L
    Fo4=sheet.cell_value(2, 6)  #diesel Energy equivalent/L
    Fo7=sheet.cell_value(4, 5)  #natural gas Co2 equivalent/L
    Fo8=sheet.cell_value(4, 6)  #natural gas Energy equivalent/L
    Fo9=sheet.cell_value(5, 5)  #oil Co2 equivalent/L
    Fo10=sheet.cell_value(5, 6) #oil Energy equivalent/L
    Fo11=sheet.cell_value(6, 6) #hard coal Co2 equivalent/L
    F012=sheet.cell_value(6, 6) #hard coal Energy equivalent/L

    #fertelizer

    Fe1=sheet.cell_value(1, 9)   #amonium nitrate Co2 equivalent/L
    Fe2=sheet.cell_value(1, 10)  #amonium nitrate energy equivalent/L
    Fe3=sheet.cell_value(2, 9)   #Calcium Ammonium Nitrate Co2 equivalent/L
    Fe4=sheet.cell_value(2, 10)  #Calcium Ammonium Nitrate energy equivalent/L
    Fe5=sheet.cell_value(3, 9)   #Ammonium Sulphate Co2 equivalent/L
    Fe6=sheet.cell_value(3, 10)  #Ammonium Sulphate energy equivalent/L
    Fe7=sheet.cell_value(4, 9)   #Triple Superphosphate Co2 equivalent/L
    Fe8=sheet.cell_value(4, 10)  #Triple Superphosphate energy equivalent/L
    Fe9=sheet.cell_value(5, 9)   #Single super phosphate Co2 equivalent/L
    Fe10=sheet.cell_value(5, 10) #Single super phosphate energy equivalent/L
    Fe11=sheet.cell_value(6, 9)  #Ammonia Co2 equivalent/L
    Fe12=sheet.cell_value(6, 10) #Ammonia energy equivalent/L
    Fe13=sheet.cell_value(7, 9)  #limestone Co2 equivalent/L
    Fe14=sheet.cell_value(7, 10) #limestone energy equivalent/L
    Fe15=sheet.cell_value(8, 9)  #NPK 15-15-15 Co2 equivalent/L
    Fe16=sheet.cell_value(8, 10) #NPK 15-15-15 energy equivalent/L
    Fe17=sheet.cell_value(9, 9)  #Urea Co2 equivalent/L
    Fe18=sheet.cell_value(9, 10) #Urea energy equivalent/L
    Fe19=sheet.cell_value(10, 9) #cow manure Co2 equivalent/L
    Fe20=sheet.cell_value(10, 10)#cow manure energy equivalent/L
    Fe21=sheet.cell_value(11, 9) #phosphoric acid Co2 equivalent/L
    Fe22=sheet.cell_value(11, 10)#phosphoric acid energy equivalent/L
    Fe23=sheet.cell_value(12, 9) #Mono-ammonium phosphate Co2 equivalent/L
    Fe24=sheet.cell_value(12, 10)#Mono-ammonium phosphate energy equivalent/L

    #substrate
    S1=sheet.cell_value(1, 13)  #Rockwool Co2 equivalent/
    S2=sheet.cell_value(1, 14)  #Rockwool energy equivalent/L
    S3=sheet.cell_value(2, 13)  #Perlite Co2 equivalent/L
    S4=sheet.cell_value(2, 14)  #Perlite energy equivalent/L
    S5=sheet.cell_value(3, 13)  #Coco Fiber (coir pith) Co2 equivalent/L
    S6=sheet.cell_value(3, 14)  #Coco Fiber (coir pith) energy equivalent/L
    S7=sheet.cell_value(4, 13)  #Hemp fiber Co2 equivalent/L
    S8=sheet.cell_value(4, 14)  #Hemp fiber energy equivalent/L
    S9=sheet.cell_value(5, 13)  #Peat Co2 equivalent/L
    S10=sheet.cell_value(5, 14) #Peat energy equivalent/L
    S11=sheet.cell_value(6, 13) #Peat moss Co2 equivalent/L
    S12=sheet.cell_value(6, 14) #Peat moss energy equivalent/L

    #water
    W1=sheet.cell_value(1, 17)  #Tapwater Co2 equivalent/L
    W2=sheet.cell_value(1, 18)  #Tapwater energy equivalent/L

    #pesticide 
    P1=sheet.cell_value(1, 21) #Atrazine water Co2 equivalent/L
    P2=sheet.cell_value(1, 22) #Atrazine energy equivalent/L
    P3=sheet.cell_value(2, 21) #Glyphosphate water Co2 equivalent/L
    P4=sheet.cell_value(2, 22) #Glyphosphate energy equivalent/L
    P5=sheet.cell_value(3, 21) #Metolachlor Co2 equivalent/L
    P6=sheet.cell_value(3, 22) #Metolachlor energy equivalent/L
    P7=sheet.cell_value(4, 21) #Herbicide Co2 equivalent/L
    P8=sheet.cell_value(4, 22) #Herbicide energy equivalent/L
    P9=sheet.cell_value(5, 21) #Insectiside Co2 equivalent/L
    P10=sheet.cell_value(5, 22)#Insectiside energy equivalent/L

    #waste
    W1=sheet.cell_value(1, 25) #green waste Co2 equivalent/kg
    W2=sheet.cell_value(1, 26) #green waste energy equivalent/L
    W3=sheet.cell_value(2, 25) #other waste Co2 equivalent/kg
    W4=sheet.cell_value(2, 26) #other waste energy equivalent/L
    W5=sheet.cell_value(3, 25) #paper waste Co2 equivalent/kg
    W6=sheet.cell_value(3, 26) #paper waste energy equivalent/L
    
    #Transport plane
    Tvp1=sheet.cell_value(1, 29)  # regression factor plane
    Tvp2=sheet.cell_value(2, 29)  # regression factor plane
    Tvp3=sheet.cell_value(4, 29)  # extra distance travelled plane
    Tvp4=sheet.cell_value(5, 29)  # extra langdings made plane
    Tvp5=sheet.cell_value(6, 29)  # landing take of kerosene usage plane
    Tvp6=sheet.cell_value(7, 29)  # co2 emissions 1 L of kerosene (Co2-eq/kg) plane	
    Tvp7=sheet.cell_value(8, 29)  # radiative forcing factor plane 	
    Tvp8=sheet.cell_value(9, 29)  # possible cargo plane
    Tvp9=sheet.cell_value(10, 29) # average percent full plane
    T1=((Tvp1*(ans201.get()*Tvp3)**2+Tvp2*(ans201.get()*Tvp3)*Tvp6*Tvp7)+Tvp4*Tvp5*Tvp6)/(ans201.get()+0.001)/(Tvp8*Tvp9) #plane Co2 equivalent/(Ton*km)
    T2=sheet.cell_value(1, 32)   #Plane energy equivalent/(ton*Km)

    #transport Truck
    Tvt1=sheet.cell_value(12, 29) # regression factor Truck
    Tvt2=sheet.cell_value(13, 29) # regression factor Truck
    Tvt3=sheet.cell_value(14, 29) # regression factor Truck
    Tvt4=sheet.cell_value(16, 29) # extra distance travelled Truck 
    Tvt5=sheet.cell_value(17, 29) # regression factor Truck
    Tvt6=sheet.cell_value(18, 29) # Co2 emission 1 L of diesel (Co2-eq/kg) Truck	
    Tvt7=sheet.cell_value(19, 29) # max cargo (ton) Truck
    Tvt8=sheet.cell_value(20, 29) # average percent full Truck
    T3=((Tvt1*Tvt7+Tvt2*Tvt3*ans202.get()*Tvt4)*Tvt5*Tvt6)/(ans202.get()+0.001)/(Tvt8*Tvt7) #truck C02 equivalent/(Ton*km)
    T4=sheet.cell_value(12, 32) #Truck energy equivalent/(ton*Km)

    #transport ship
    Tp1=sheet.cell_value(22, 29) # regression factor plane
    Tp2=sheet.cell_value(23, 29) # regression factor plane
    Tp3=sheet.cell_value(25, 29) # extra distance travelled plane
    Tp4=sheet.cell_value(26, 29) # regression factor plane
    Tp5=sheet.cell_value(27, 29) # litre of oil used in harbours	
    Tp6=sheet.cell_value(28, 29) # extra stops see harbours	
    Tp7=sheet.cell_value(29, 29) # Co2 emission 1 L of Oil (Co2-eq/kg)	
    Tp8=sheet.cell_value(30, 29) # Cargo (ton)
    Tp9=sheet.cell_value(31, 29) # average percentage full
    T5=(((Tp1*Tp8+Tp2*Tp4*ans203.get()*Tp3+(Tp5*(Tp6+1))*Tp7)/(ans203.get()+0.001)))/(Tp8*Tp9) #ship C02 equivalent/(Ton*km)
    T6=sheet.cell_value(22, 32) #Ship energy equivalent/(ton*Km)

    #Package
    if ans171.get()==1:
        Pac1=float(sheet.cell_value(1,39)) # Packaging Co2 equivalent/kg
        Pac2=float(sheet.cell_value(1,40)) # Packaging energy equivalent/kg
    else:
        Pac1=0
        Pac2=0

    
    Eco2=(C1*ans61.get())+(C3*ans62.get())+(C5*ans71.get())+(C7*ans73.get())+(C9*ans72.get())-(ans87.get()*C1)-(ans88.get()*C3)  #caculation for total C02 of electrcity usage
    
    Eenergy=(C2*ans61.get())+(C4*ans62.get())+(C6*ans71.get())+(C8*ans73.get())+(C10*ans72.get())-(ans87.get()*C2)-(ans88.get()*C4) #caculation for total energy of electricity usage

    Fco2=(Fo1*ans91.get())+(Fo3*ans92.get())+(Fo7*ans94.get())+(Fo9*ans95.get()) #calculation for total Co2 of fossile fuels use

    Fenergy=(Fo2*ans91.get())+(Fo4*ans92.get())+(Fo8*ans94.get())+(Fo10*ans95.get()) #calculation for total energy of fissile fuele use

    FERco2=((Fe1*ans121.get())+(Fe3*ans122.get())+(Fe5*ans123.get())+(Fe7*ans124.get())+(Fe9*ans125.get())+(Fe11*ans126.get())+(Fe13*ans127.get())+(Fe15*ans128.get())+(Fe17*ans129.get())+(Fe19*ans130.get())+(Fe21*ans131.get())+(Fe22*ans132.get())) #calculation for total Co2 of fertelizers

    FERenergy=((Fe2*ans121.get())+(Fe4*ans122.get())+(Fe6*ans123.get())+(Fe8*ans124.get())+(Fe10*ans125.get())+(Fe12*ans126.get())+(Fe14*ans127.get())+(Fe16*ans128.get())+(Fe18*ans129.get())+(Fe20*ans130.get())+(Fe22*ans131.get())+(Fe24*ans132.get()))  #calculation for total energy of fertelizers

    Sco2=(S1*ans141.get())+(S3*ans142.get())+(S5*ans143.get())+(S7*ans144.get())+(S9*ans145.get())+(S11*ans146.get()) #calculation for total Co2 of substrates

    Senergy=(S2*ans141.get())+(S4*ans142.get())+(S6*ans143.get())+(S8*ans144.get())+(S10*ans145.get())+(S12*ans146.get()) #calculation for total energy of substrates

    Wco2= (W1*ans151.get()) #calculation for total Co2 of water

    Wenergy = (W2*ans151.get()) #calculation for total energy of water

    Pco2 = (P1*ans161.get())+(P3*ans162.get())+(P5*ans163.get())+(P7*ans164.get())+(P9*ans165.get()) #calculation for total Co2 of pesticides 

    Penergy = (P2*ans161.get())+(P4*ans162.get())++(P6*ans163.get())++(P8*ans164.get())+(P10*ans165.get()) #calculation for total energy of pesticides

    Tco2= (T1* ans201.get())+(T3*ans202.get())+(T5*ans203.get()) #calculation for total Co2 of transport

    Tenergy= (T2* ans201.get())+(T4*ans202.get())+(T6*ans203.get()) #calculation for total energy of transport
    
    Pacco2=(ans5.get()*(100-ans211.get())/100)*Pac1

    Pacenergy=(ans5.get()*(100-ans211.get())/100)*Pac2

    Totalco2 = Eco2+Fco2+FERco2+ Sco2+Wco2+Pco2+Tco2+ Pacco2
    Totalenergy = Eenergy+Fenergy+FERenergy+Senergy+Wenergy+Penergy+Tenergy+Pacenergy

    Totalco2_per_kg_product= Totalco2/(ans5.get()*(100-ans211.get())/100)
    Totalenergy_per_kg_product= Totalenergy/(ans5.get()*(100-ans211.get())/100)

    Totalco2_per_KJ_product=Totalco2/((ans5.get()*(100-ans211.get())/100)*Eoc)
    Totalenergy_per_KJ_product=Totalenergy/((ans5.get()*(100-ans211.get())/100)*Eoc)

    wb = xlsxwriter.Workbook(farm_name.get()+ans3.get()+'.xlsx')
    ws = wb.add_worksheet('answer')
    ws.write(0,2,'Total CO2 emitted(Kg)')
    ws.write(0,4,'Total energy used(MJ)')
    ws.write(1,0,'Electricity ')
    ws.write(1,2,Eco2)
    ws.write(1,4,Eenergy)

    ws.write(2,0,'Fossil fuels')
    ws.write(2,2,Fco2)
    ws.write(2,4,Fenergy)

    ws.write(3,0,'Fertelizer')
    ws.write(3,2,FERco2)
    ws.write(3,4,FERenergy)

    ws.write(4,0,'Substrates')
    ws.write(4,2,Sco2)
    ws.write(4,4,Senergy)

    ws.write(5,0,'Water')
    ws.write(5,2,Wco2)
    ws.write(5,4,Wenergy)

    ws.write(6,0,'Pesticides')
    ws.write(6,2,Pco2)
    ws.write(6,4,Penergy)

    ws.write(7,0,'Transport')
    ws.write(7,2,Tco2)
    ws.write(7,4,Tenergy)

    ws.write(8,0,'Package')
    ws.write(8,2,Pacco2)
    ws.write(8,4,Pacenergy)

    ws.write(1,6, 'Total CO2 emitted (Kg)')
    ws.write(1,9,Totalco2)
    ws.write(2,6, 'Total energy used(MJ)')
    ws.write(2,9,Totalenergy)

    ws.write(3,6, 'Total CO2 emitted per kg product (Kg/Kg)')
    ws.write(3,9,Totalco2_per_kg_product)
    ws.write(4,6, 'Total Energy used per kg product (KJ/Kg)')
    ws.write(4,9,Totalenergy_per_kg_product)

    ws.write(5,6, 'Total CO2 emitted per KJ product (Kg/KJ)')
    ws.write(5,9,Totalco2_per_KJ_product)
    ws.write(6,6, 'Total energy used per KJ product (Kg/KJ)')
    ws.write(6,9,Totalenergy_per_KJ_product)

    ws.write(43,1,'Heating')
    ws.write(44,1,'Cooling')
    ws.write(45,1,'Electricity')
    ws.write(46,1,'Tillage')
    ws.write(47,1,'Sowing')
    ws.write(48,1,'Weeding')
    ws.write(49,1,'Harvest')
    ws.write(50,1,'Fertilize')
    ws.write(51,1,'Irrigation')
    ws.write(52,1,'Pesticide')
    ws.write(53,1,'Other')
    ws.write(43,2,ans97.get())
    ws.write(44,2,ans98.get())
    ws.write(45,2,ans99.get())
    ws.write(46,2,ans910.get())
    ws.write(47,2,ans911.get())
    ws.write(48,2,ans912.get())
    ws.write(49,2,ans913.get())
    ws.write(50,2,ans914.get())
    ws.write(51,2,ans915.get())
    ws.write(52,2,ans916.get())
    ws.write(53,2,ans917.get())
    ws.write(43,9,'Heating')
    ws.write(44,9,'Cooling')
    ws.write(45,9,'Ventillation')
    ws.write(46,9,'Lighting')
    ws.write(47,9,'Machinery')
    ws.write(48,9,'Storage')
    ws.write(49,9,'Selling')
    ws.write(50,9,'Other')

    ws.write(43,10,ans81.get())
    ws.write(44,10,ans82.get())
    ws.write(45,10,ans83.get())
    ws.write(46,10,ans84.get())
    ws.write(47,10,ans85.get())
    ws.write(48,10,ans86.get())
    ws.write(49,10,ans87.get())
    ws.write(50,10,ans88.get())
    
    if ans890.get()==1 or ans131.get()==1 or ans147.get()==1 or ans152.get()==1 or ans166.get()==1 or ans184.get()==1 or ans204.get()==1:
        ws.write(10,1,non_count+'is not taken into account because of lacking data')
    if v.get()==1:
        ws.write(8,6,'Greenhouse')
    if v.get()==2:
        ws.write(8,6,'Open Field')
    if v.get()==3:
        ws.write(8,6,'Vertical Farm')
        
    chart_col= wb.add_chart({'type':'column'})
    chart_col.add_series({
        'name':['answer',0,2],
        'categories' :['answer',1,0,10,0],
        'values':['answer',1,2,10,2],
        'line':{'color':'yellow'}
        })
    chart_col.set_title({'name': 'Total CO2 emitted from different sources'})
    chart_col.set_y_axis({'name': 'Total CO2 emitted'})
    chart_col.set_x_axis({'name':  'Different sources'})

    chart_col.set_style(2)
    ws.insert_chart('A13', chart_col, {'x_offset': 20, 'y_offset': 8})

    chart_col= wb.add_chart({'type':'column'})
    chart_col.add_series({
        'name':['answer',0,4],
        'categories' :['answer',1,0,10,0],
        'values':['answer',1,4,10,4],
        'line':{'color':'yellow'}
        })
    chart_col.set_title({'name': 'Total energy used from different sources'})
    chart_col.set_y_axis({'name': 'Total energy used'})
    chart_col.set_x_axis({'name':  'Different sources'})

    chart_col.set_style(2)
    ws.insert_chart('I13', chart_col, {'x_offset': 20, 'y_offset': 8})

    chart_col=wb.add_chart({'type':'pie'})
    chart_col.add_series({
        'name':['answer',0,1],
        'categories' :['answer',1,0,10,0],
        'values':['answer',1,2,10,2],
        'points':[{'fill':{'color':'blue'}},
                  {'fill':{'color':'yellow'}},
                  {'fill':{'color':'red'}},
                  {'fill':{'color':'gray'}},
                  {'fill':{'color':'black'}},
                  {'fill':{'color':'purple'}},
                  {'fill':{'color':'pink'}},
                  ],
        })
    chart_col.set_title({'name': 'Total CO2 emitted from different \nsources'})
    chart_col.set_style(2)
    ws.insert_chart('A28', chart_col, {'x_offset': 20, 'y_offset': 8})

    chart_col=wb.add_chart({'type':'pie'})
    chart_col.add_series({
        'name':['answer',0,2],
        'categories' :['answer',1,0,10,0],
        'values':['answer',1,4,10,4],
        'points':[{'fill':{'color':'blue'}},
                  {'fill':{'color':'yellow'}},
                  {'fill':{'color':'red'}},
                  {'fill':{'color':'gray'}},
                  {'fill':{'color':'black'}},
                  {'fill':{'color':'purple'}},
                  {'fill':{'color':'pink'}},
                  ],
        })
    chart_col.set_title({'name': 'Total energy used from different sources'})
    chart_col.set_style(2)
    ws.insert_chart('I28', chart_col, {'x_offset': 20, 'y_offset': 8})

    chart_col=wb.add_chart({'type':'pie'})
    chart_col.add_series({
        'name':'Fossil fuels used',
        'categories' :['answer',43,1,53,1],
        'values':['answer',43,2,53,2],
        'points':[{'fill':{'color':'blue'}},
                  {'fill':{'color':'yellow'}},
                  {'fill':{'color':'red'}},
                  {'fill':{'color':'gray'}},
                  {'fill':{'color':'black'}},
                  {'fill':{'color':'purple'}},
                  {'fill':{'color':'pink'}},
                  {'fill':{'color':'cyan'}},
                  {'fill':{'color':'magenta'}},
                  {'fill':{'color':'brown'}},
                  ],
        })
    chart_col.set_title({'name': 'Fossil fuels used for different aspects'})
    chart_col.set_style(5)
    ws.insert_chart('A43', chart_col, {'x_offset': 25, 'y_offset': 10})

    chart_col=wb.add_chart({'type':'pie'})
    chart_col.add_series({
        'name':'Electricity used',
        'categories' :['answer',43,9,50,9],
        'values':['answer',43,10,50,10],
        'points':[{'fill':{'color':'blue'}},
                  {'fill':{'color':'yellow'}},
                  {'fill':{'color':'red'}},
                  {'fill':{'color':'gray'}},
                  {'fill':{'color':'black'}},
                  {'fill':{'color':'purple'}},
                  {'fill':{'color':'pink'}},
                  {'fill':{'color':'cyan'}},
                  
                  ],
        })
    chart_col.set_title({'name': 'Electricity used for different aspects'})
    chart_col.set_style(4)
    ws.insert_chart('I43', chart_col, {'x_offset': 25, 'y_offset': 10})

    wb.close()


    ## this change was made to get the git repository running

    ## JW: still doing examples of VCS‘


      

