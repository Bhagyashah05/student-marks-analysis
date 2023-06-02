#importing required libraries
import numpy as np
import pandas as pd
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

#creating the Tkinter window
w=tk.Tk()
w.title("Student Result Analysis")
w.geometry('1200x800')
w.configure(background='#141401')

#adding a figure to show the graph in the window 
fig=Figure()
a=fig.add_subplot(111)
canvas=FigureCanvasTkAgg(fig,master=w)
plot_widget=canvas.get_tk_widget()

#main program where calculations are done and results are printed/plotted
def main():
    path=e.get()
    d=pd.read_excel(str(path),index_col=0, engine="openpyxl")
    total=d['TOTAL'].values
    value,count=np.unique(total,return_counts=True)
    t=dict(zip(value,count))
    a.bar(t.keys(),t.values())
    a.set_xlabel('Total Marks',fontsize=15)
    a.set_ylabel('Number of Students',fontsize=15)
    a.set_title('RESULT ANALYSIS',fontsize=20)
    fig.canvas.draw_idle()

    avg_total=d['TOTAL'].values.mean()
    avg_eng=d['ENGLISH'].values.mean()
    avg_maths=d['MATHEMATICS'].values.mean()
    avg_phy=d['PHYSICS'].values.mean()
    avg_chem=d['CHEMISTRY'].values.mean()
    avg_bio=d['BIOLOGY'].values.mean()

    dic_total=dict(zip(d.index,d['TOTAL'].values))
    max_total=max(d['TOTAL'].values)
    T.insert(tk.INSERT,"\n1. ")
    for i_total in dic_total.keys():
        if dic_total[i_total]==max_total:
            T.insert(tk.INSERT,str(i_total)+",")
    T.insert(tk.INSERT," secured the highest marks in total "+str(max_total)+".\n\n")
        
    dic_eng=dict(zip(d.index,d['ENGLISH'].values))
    max_eng=max(d['ENGLISH'].values)
    T.insert(tk.INSERT,"2. ")
    for i_eng in dic_eng.keys():
        if dic_eng[i_eng]==max_eng:
            T.insert(tk.INSERT,str(i_eng)+",")
    T.insert(tk.INSERT," secured the highest marks in English "+str(max_eng)+".\n\n")
        
    dic_maths=dict(zip(d.index,d['MATHEMATICS'].values))
    max_maths=max(d['MATHEMATICS'].values)
    T.insert(tk.INSERT,"3. ")
    for i_maths in dic_maths.keys():
        if dic_maths[i_maths]==max_maths:
            T.insert(tk.INSERT,str(i_maths)+",")
    T.insert(tk.INSERT," secured the highest marks in Mathematics "+str(max_maths)+".\n\n")
        
    dic_phy=dict(zip(d.index,d['PHYSICS'].values))
    max_phy=max(d['PHYSICS'].values)
    T.insert(tk.INSERT,"4. ")
    for i_phy in dic_phy.keys():
        if dic_phy[i_phy]==max_phy:
            T.insert(tk.INSERT,str(i_phy)+",")
    T.insert(tk.INSERT," secured the highest marks in Physics "+str(max_phy)+".\n\n")
        
    dic_chem=dict(zip(d.index,d['CHEMISTRY'].values))
    max_chem=max(d['CHEMISTRY'].values)
    T.insert(tk.INSERT,"5. ")
    for i_chem in dic_chem.keys():
        if dic_chem[i_chem]==max_chem:
            T.insert(tk.INSERT,str(i_chem)+",")
    T.insert(tk.INSERT," secured the highest marks in Chemistry "+str(max_chem)+".\n\n")
        
    dic_bio=dict(zip(d.index,d['BIOLOGY'].values))
    max_bio=max(d['BIOLOGY'].values)
    T.insert(tk.INSERT,"6. ")
    for i_bio in dic_bio.keys():
        if dic_bio[i_bio]==max_bio:
            T.insert(tk.INSERT,str(i_bio)+",")
    T.insert(tk.INSERT," secured the highest marks in Biology "+str(max_bio)+".\n\n")
        
    above_avg=0
    above_pass=0
    pass_marks=400
    for j in d['TOTAL'].values:
        if j>avg_total:
            above_avg+=1
    for k in d['TOTAL'].values:
        if k>=pass_marks:
            above_pass+=1


    T.insert(tk.INSERT,"7. Average marks of students is "+str("%.2f"%avg_total)+" and "+str(above_avg)+" students have secured marks above average marks.\n\n")
    T.insert(tk.INSERT,"8. Average marks in English is "+str("%.2f"%avg_eng)+".\n\n")
    T.insert(tk.INSERT,"9. Average marks in Mathematics is "+str("%.2f"%avg_maths)+".\n\n")
    T.insert(tk.INSERT,"10. Average marks in Physics is "+str("%.2f"%avg_phy)+".\n\n")
    T.insert(tk.INSERT,"11. Average marks in Chemistry is "+str("%.2f"%avg_chem)+".\n\n")
    T.insert(tk.INSERT,"12. Average marks in Biology is "+str("%.2f"%avg_bio)+".\n\n")
    T.insert(tk.INSERT,"13. The pass marks is "+str(pass_marks)+" and "+str(above_pass)+" students have passed.\n\n")
    T.config(state="disabled")

#creating different Tkinter widgets and defining their positions in the tkinter window
label=tk.Label(w,text='STUDENT RESULT ANALYSIS',font=("Arial Bold",15),fg="red",bg="#FDFD96").grid(row=2,column=2)
enter=tk.Label(w,text='Enter path of Excel file',font=15,bg="#FDFD96").grid(row=4,column=2)
e=tk.Entry(w,width=60)
e.grid(row=4,column=3)
b=tk.Button(w,text='Go',command=main).grid(row=4,column=4)
plot_widget.grid(row=5,column=2,rowspan=15,columnspan=50)
b1=tk.Button(w,text='Quit',command=w.destroy)
b1.grid(row=22,column=3)
T=tk.Text(w,fg="blue",bg="pink",width=55,height=40)
T.grid(row=5,column=1)
w.mainloop()
