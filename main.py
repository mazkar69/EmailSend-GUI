from tkinter import * 
from tkinter import messagebox,filedialog
from PIL import ImageTk,Image
import os
import pandas as pd
import smtplib



class Email():



    def __init__(self,root):
        self.check=0        #This variable will work like the radio button when we choice the bulk option it will be 1 and when we choice the single option it will be the 0 
        self.root=root
        root.title("BULK EMAIL APPLICATION | By AzKaR")
        root.geometry("800x450+200+50")
        icon=PhotoImage(file="email_32.png")
        

        root.resizable(0,0)
        root.config(bg='white')
        root.iconphoto(False, icon)

        #===============notification status===============
        des=Label(self.root,text='To send the Bulk Mail,Change the Mode',padx=15,font=('Calibri (Body)',12,'italic'),bg='#FFD966',fg='black').place(x=0,y=0,relwidth=True)

        #==================Menu bar ================

        mainmenu=Menu(self.root)

        m0=Menu(root,tearoff=0)
        m0.add_command(label="Single",font=("times new roman",12,'italic'),command=self.single)
        m0.add_command(label="Bulk",font=("times new roman",12,'italic'),command=self.bulk)

        m1=Menu(mainmenu,tearoff=0)
        m1.add_command(label="Uses")
        m1.add_command(label="Contact Us")
        self.root.config(menu=mainmenu)

        mainmenu.add_cascade(label="Mode",menu=m0)
        mainmenu.add_command(label="setting",command=self.setting_window)
        mainmenu.add_cascade(label="Help",menu=m1)
        root.config(menu=mainmenu)

        #================Laleb===============================

        to=Label(root,text='To (Email Address)',font=('times new roman',18),bg='white').place(x=50,y=50)
        subj=Label(root,text='SUBJECT',font=('times new roman',18),bg='white').place(x=50,y=100)
        msg=Label(root,text='MASSAGE',font=('times new roman',18),bg='white').place(x=50,y=150)
        

        #==================Entry field==================
        self.txt_to=Entry(root,font=('times new roman',14),bg='lightyellow')
        self.txt_to.place(x=350,y=50,w=350,h=30)

        self.subj=Entry(root,font=('times new roman',14),bg='lightyellow')
        self.subj.place(x=350,y=100,w=400,h=30)

        self.msg=Text(root,font=('times new roman',15,'italic'),bg='lightyellow')
        self.msg.place(x=350,y=150,w=400,h=100)


        #==============button====================
        self.clear=Button(root,text='CLEAR',command=self.clear1,activebackground="#262626",font=('times new roman',18,'bold'),cursor='hand2',bg='#262626',fg='white',activeforeground="white")
        self.clear.place(x=400,y=300,w=120,h=30)

        self.send=Button(root,text='SEND',command=self.sendfunc,activebackground="#00B0F0",font=('times new roman',18,'bold'),cursor='hand2',bg='#00B0F0',fg='white',activeforeground="white")
        self.send.place(x=550,y=300,w=120,h=30)

        #==========browse btn=============

        self.upload_img=ImageTk.PhotoImage(file="upload.png")
        

        self.browse_btn=Button(root,image=self.upload_img,cursor='hand2',state=DISABLED,command=self.browse_file)
        self.browse_btn.place(x=720,y=45)

        #==================status bar=====================
        self.lbl_total=Label(root,text='',font=('times new roman',18),bg='white')
        self.lbl_total.place(x=50,y=370)

        self.lbl_sent=Label(root,text='',font=('times new roman',18),bg='white',fg='green')
        self.lbl_sent.place(x=300,y=370)

        self.lbl_left=Label(root,text='',font=('times new roman',18),bg='white',fg="orange")
        self.lbl_left.place(x=420,y=370)

        self.lbl_failed=Label(root,text='',font=('times new roman',18),bg='white',fg="red")
        self.lbl_failed.place(x=550,y=370)

        self.check_file_exits()


    #===========================================Methoda are written bellow=======================================

    def single(self):                               
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0,END)
        self.subj.delete(0,END)
        self.msg.delete(1.0,END)
        self.lbl_total.config(text="")
        self.lbl_left.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_failed.config(text="")
        self.check=0
        # print(self.check)
        self.browse_btn.config(state=DISABLED)
        

    def browse_file(self):                      #To select the exce file
        
        op=filedialog.askopenfile(title="Select Excel File for Emails.")
        if op!=None:
            
            data=pd.read_excel(op.name)
            self.txt_to.config(state=NORMAL)
            self.txt_to.insert(0,str(op.name.split("/")[-1]))
            self.txt_to.config(state=DISABLED)
    
            if 'Email' or 'email' or 'emails' or 'Emails' in data.columns:
                
                email=list(data['Email'])
                # print(email)
                self.c=[]
                for i in email:
                    if pd.isnull(i)==False:
                        self.c.append(i)
                        # print(i)
                self.lbl_total.config(text=f"Total: {str(len(self.c))}")
                self.lbl_sent.config(text=f"SENT:")
                self.lbl_left.config(text=f"LEFT:")
                self.lbl_failed.config(text=f"FAILED:")

            else:

                messagebox.showerror("Error","Please Select the file which have email column")


    def bulk(self):
        self.txt_to.delete(0,END)
        self.subj.delete(0,END)
        self.msg.delete(1.0,END)
        self.check=1
        # print(self.check)
        self.browse_btn.config(state=NORMAL)
        self.txt_to.config(state=DISABLED)

    def clear1(self):                           #clear button funtion
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0,END)
        self.subj.delete(0,END)
        self.msg.delete(1.0,END)
        self.lbl_total.config(text="")
        self.lbl_left.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_failed.config(text="")
        self.browse_btn.config(state=DISABLED)
        
    def sendfunc(self):                             #when we press on the send button this function will triggered
        x=len(self.msg.get(1.0,END))
        # print(x)
        if self.txt_to.get()=="" or self.subj.get()=="" or x==1:                #This is for authantication we any field are empty it will show the error
            messagebox.showerror("Error","All fields are required")
            # print("All field are required!")

        else:
            if self.check==0:                       #This self.check is to know that whether the the user has selected the bulk or singe email for single it is zero and for nulk it is 1 , when user click on the bulk button it will chage the 1 .
                status=self.email_send_funct(self.txt_to.get(),self.subj.get(),self.msg.get(1.0,END),self.from_,self.pass_)        #passig the value to the email send function that will send the email , this funciton will return the s charector for success and f for failed
                if status=='s':
                    messagebox.showinfo("Success","Email Has Been Send Successful.")
                if status=='f':
                    messagebox.showerror("Failed",'Email not sent')


            if self.check==1:
                self.s_count=0
                self.f_count=0
                self.l_count=0
                for i in self.c:                #this one is for bulk , this loop will send one email at a time as self.c contail the list of emails
                    status=self.email_send_funct(i,self.subj.get(),self.msg.get(1.0,END),self.from_,self.pass_)
                    if status=='s':
                        self.s_count+=1
                    if status=='f':
                        self.f_count+=1

                    self.status_bar()               #calling the status bar 
                    
                messagebox.showinfo("Success",'Email Has Sent, Please Check Status')                #When all the msg send sent this masse box will shoe for success .

    def status_bar(self):
        self.lbl_total.config(text=f"STATUS: {str(len(self.c))} =>>")
        self.lbl_sent.config(text=f"SENT: {str(self.s_count)}")
        self.lbl_left.config(text=f"LEFT: {str(len(self.c)-(self.s_count+self.f_count))}")
        self.lbl_failed.config(text=f"FAILED {str(self.f_count)}")

        self.lbl_total.update()
        self.lbl_sent.update()
        self.lbl_left.update()
        self.lbl_failed.update()

    def email_send_funct (self,to_,subj_,msg_,from_,pass_):      
        s=smtplib.SMTP("smtp.gmail.com",587)   #Creating the session for gmail
        s.starttls()    #Transport layer
        s.login(from_,pass_)        #log in 
        msg=f"Subject:{subj_}\n\n{msg_}"
        s.sendmail(from_,to_,msg)
        x=s.ehlo()
            
        if x[0]==250:
                return 's'
        else:
                return 'f'

        s.close()


    def setting_window(self):                   #Settng window , top level window settng and properties 
        self.check_file_exits()
        self.root1=Toplevel()
        self.root1.resizable(0,0 )
        self.root1.title("Setting")
        self.root1.geometry('700x350+250+90')
        self.root1.focus_force()
        self.root1.grab_set()
        self.root1.config(bg='white')
        self.img2=ImageTk.PhotoImage(file="setting.png")
        self.root1.iconphoto(False, self.img2)

        self.title=Label(self.root1,text='Credintial Setting',padx=15,font=('Goudy Old Style',30,'bold'),bg='#222A35',fg='white').place(x=0,y=0,relwidth=True)

        des=Label(self.root1,text='Enter the Email Address and Password from which You want to Send all the Emails',padx=15,font=('Calibri (Body)',12,'italic'),bg='#FFD966',fg='black').place(x=0,y=50,relwidth=True)


        from_=Label(self.root1,text='Email Address',font=('times new roman',18),bg='white').place(x=50,y=100)
        password=Label(self.root,text='PASSWORD',font=('times new roman',18),bg='white').place(x=50,y=200)

        self.from_to=Entry(self.root1,font=('times new roman',14),bg='lightyellow')
        self.from_to.place(x=250,y=100,w=330,h=30)

        password=Label(self.root1,text='PASSWORD',font=('times new roman',18),bg='white').place(x=50,y=150)
        self.ent_password=Entry(self.root1,font=('times new roman',14),bg='lightyellow',show='*')
        self.ent_password.place(x=250,y=150,w=330,h=30)

        self.clear2=Button(self.root1,text='CLEAR',activebackground="#262626",font=('times new roman',18,'bold'),cursor='hand2',bg='#262626',fg='white',activeforeground="white",command=self.clear3)
        self.clear2.place(x=300,y=210,w=120,h=30)
        
        self.save=Button(self.root1,text='Save',activebackground="#00B0F0",font=('times new roman',18,'bold'),cursor='hand2',bg='#00B0F0',fg='white',activeforeground="white",command=self.save_setting)
        self.save.place(x=430,y=210,w=120,h=30)

        #These two line will help to add the email and password on the Entry field . 
        self.from_to.insert(0,self.from_)      #we already assing the value of self.from_ and the vlue of self.pass_ ,by calling the check_file_exixt function on the top ,.
        self.ent_password.insert(0,self.pass_)


    def clear3(self):                   #This clear method is for the secound window
        self.from_to.delete(0,END)
        self.ent_password.delete(0,END)

    def save_setting(self):                     #this will save the email and password in the file
        if self.ent_password.get()=="" or self.from_to.get()=="":
            messagebox.showerror("Error","Please enter the both field",parent=self.root)
        else:
            f=open("important.txt",'w')
            f.write(self.from_to.get()+":"+self.ent_password.get())
            f.close()
            messagebox.showinfo("Success",'Saved Successfully')
            self.check_file_exits()

    
    def check_file_exits(self):                         #This function is to check that file is available or not if availabe then assing the email and password to the variable .
        if os.path.exists("important.txt")==False:
            f=open("important.txt",'w')
            f.write(":")
            f.close()

        f2=open('important.txt','r')
        f3=f2.read()
        self.cren=f3.split(':')
        # print(self.cren)
        self.from_=self.cren[0]
        self.pass_=self.cren[1]
        f2.close()

    

root=Tk()

ob1=Email(root)




root.mainloop()


