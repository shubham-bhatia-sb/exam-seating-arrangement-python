
from tkinter import *
from tkinter import messagebox
import xlsxwriter
from tkinter.ttk import * 
root = Tk();
class App :
    def __init__(self,root) :
        root.title("SEATING PLAN MAKER VER : 1.0.1 ")
        style = Style();
        style.configure('TButton',background='#03F9A4',foreground="blue");
        style.configure('TEntry',foreground="green");
        root.iconbitmap(r"favicon.ico");
        root.minsize(height=280,width=400);
        self.menu = StringVar();
        self.menu.set("SELECT SEATING PLAN TYPE:");   
        Label(root,text="SELCT PLAN TYPE:",font=("Courier",10)).pack();
        canvas = Canvas(root,width=280, height = 150)      ;
        canvas.pack()      
        self.img = PhotoImage(file= r"type.png") ;
             
        canvas.create_image(10,10, anchor=NW, image= self.img);
        #Label(root,text="COPYRIGHT BY SHUBHAM BHATIA & AMIT MISHRA:-)",font=("Courier",9)).place(relx=0.0,rely=1.0,anchor='sw');
        OptionMenu(root,self.menu,"SELECT:","TYPE1","TYPE2","TYPE3","TYPE4",command=self.show).pack();
 
    def show(self,n) :
        man = str(self.menu.get());
        self.btn = Button(root,text="LET's GET STARTED");
        self.btn.pack();
        if(man == "TYPE1") :
          self.btn.configure(command=self.TYPE1)
            
        elif(man == "TYPE2") :
            self.btn.configure(command=self.TYPE2)
        elif(man=="TYPE3") :
                self.btn.configure(command=self.TYPE3)
        elif(man=="TYPE4") :
                self.btn.configure(command=self.TYPE4)        
    
    #FUNCTIONS FOR DIFFERENT TYPES : 
    # TYPE1 : A   TYPE2 : AB       TYPE3 : AB     TYPE4 : AB
    #         A           AB               CD             BA
    


    def TYPE1(self) :
        self.btn.configure(state=DISABLED);
        label1= Label(root,text="EXAM DATE",font=("Courier",10));
        label1.pack();
        self.date = StringVar();
        entry1 = Entry(root,textvariable=self.date);
        entry1.pack()
        #SET ROOM NAME :::::
        label2= Label(root,text="ROOM NAME:",font=("Courier",10));
        label2.pack();
        self.room_name = StringVar();
        entry2 = Entry(root,textvariable=self.room_name);
        entry2.pack();
        #SET CLASS 1
        label3= Label(root,text="CLASS NAME",font=("Courier",10));
        label3.pack();
        self.class1 = StringVar();
        entry3 = Entry(root,textvariable=self.class1);
        entry3.pack()
        #STARTING ROLL FOR CLASS 1;
        label4= Label(root,text="STR ROLL FOR CLASS ",font=("Courier",10));
        label4.pack();
        self.sr1 = StringVar();
        entry4 = Entry(root,textvariable=self.sr1);
        entry4.pack()
        # file name ::::::
        label8= Label(root,text="ENTER FILE NAME : ",font=("Courier",10));
        label8.pack();
        self.file = StringVar();
        entry8 = Entry(root,textvariable=self.file);
        entry8.pack()
        # COMPILE BUTTON : 
        Button(root,text="COMPILE PLAN :-)" ,command=self.compile2).pack();

    def compile2(self) :
        # zaruri jankeri jo chahiye 
        dated = self.date.get();
        room = self.room_name.get();
        class_a = str(self.class1.get())+" ";
        
        roll_a = int(self.sr1.get());
        
        fn = self.file.get();
        filen = fn+".xlsx"
        # ye exel wala idea amit se btw tq amit :-)
        wrk = xlsxwriter.Workbook(filen);
        worksheet = wrk.add_worksheet();
        worksheet.write(0,4,"JAWAHAR NAVODAYA VIDYALAYA SEATING PLAN FOR ROOM : "+str(room))
        worksheet.write(1,1,"EXAM DATE: "+str(dated));
        
        worksheet.write(1,5,"ROOM :"+str(room));
        
        # SETTING ROWS 
        for i in range(8):
           worksheet.write(3,i+1,"ROW "+str(i+1));
       
        for i in range(5):
            worksheet.write(4+i,0,"COL:"+str(i+1));
        # enter code here row and columns are alrready added
         #for row 1
        for i in range(5):
            worksheet.write(4+i,1,(class_a+str(roll_a+i)));
        #for row 2
        for i in range(5):
            worksheet.write(4+i,2,(class_a+str(roll_a+5+i)));
        #for row 3
        for i in range(5):
            worksheet.write(4+i,3,(class_a+str(roll_a+5+5+i)));
        #for row 4
        for i in range(5):
            worksheet.write(4+i,4,(class_a+str(roll_a+5+5+5+i)));
        
        
        
        
        wrk.close();    
        # ADDING CHILDREN TO SHEET
        #code for TYPE1
        print("TYPE1")

# TYPE 2 ------------------------------------------------------------------

    def TYPE2(self) :
        self.btn.configure(state=DISABLED);
        # GUI -----
        label1= Label(root,text="EXAM DATE",font=("Courier",10)).pack();
        self.date = StringVar();
        entry1 = Entry(root,textvariable=self.date).pack();
        label2= Label(root,text="ROOM NAME:",font=("Courier",10));
        label2.pack();
        self.room_name = StringVar();
        entry2 = Entry(root,textvariable=self.room_name);
        entry2.pack();
        label3= Label(root,text="CLASS 1st NAME",font=("Courier",10));
        label3.pack();
        self.class1 = StringVar();
        entry3 = Entry(root,textvariable=self.class1);
        entry3.pack()
        #STARTING ROLL FOR CLASS 1;
        label4= Label(root,text="STR ROLL CLASS 1st",font=("Courier",10));
        label4.pack();
        self.sr1 = StringVar();
        entry4 = Entry(root,textvariable=self.sr1);
        entry4.pack()
        #SET CLASS 2
        label5= Label(root,text="CLASS 2nd NAME",font=("Courier",10));
        label5.pack();
        self.class2 = StringVar();
        entry5 = Entry(root,textvariable=self.class2);
        entry5.pack()
        #STARTING ROLL FOR CLASS 2;
        label6= Label(root,text="STR ROLL CLASS 2nd",font=("Courier",10));
        label6.pack();
        self.sr2 = StringVar();
        entry6 = Entry(root,textvariable=self.sr2);
        entry6.pack()
        # file name ::::::
        label8= Label(root,text="ENTER FILE NAME : ",font=("Courier",10));
        label8.pack();
        self.file = StringVar();
        entry8 = Entry(root,textvariable=self.file);
        entry8.pack()
        # COMPILE BUTTON : 
        Button(root,text="COMPILE PLAN :-)" ,command=self.compile1).pack();
        # result
    def compile1(self) :
        self.btn.configure(state=DISABLED);
        # zaruri jankeri jo chahiye 
        dated = self.date.get();
        room = self.room_name.get();
        class_a = str(self.class1.get())+" ";
        class_b = str(self.class2.get())+" ";
        roll_a = int(self.sr1.get());
        roll_b = int(self.sr2.get());
        fn = self.file.get();
        filen = fn+".xlsx"
        # ye exel wala idea amit se btw tq amit :-)
        wrk = xlsxwriter.Workbook(filen);
        worksheet = wrk.add_worksheet();
        worksheet.write(0,4,"JAWAHAR NAVODAYA VIDYALAYA SEATING PLAN FOR ROOM : "+str(room))
        worksheet.write(1,1,"EXAM DATE: "+str(dated));
        
        worksheet.write(1,5,"ROOM :"+str(room));
        
        # SETTING ROWS 
        for i in range(8):
           worksheet.write(3,i+1,"ROW "+str(i+1));
       
        for i in range(5):
            worksheet.write(4+i,0,"COL:"+str(i+1));
        # ADDING CHILDREN TO SHEET    
        r1=int(roll_a);
        r2=int(roll_b);
        for i in range(5) :
            worksheet.write(4+i,1,class_a+str(r1+i));  
        for i in range(5) :
            worksheet.write(4+i,2,class_b+str(r2+i));

        for i in range(5) :
            worksheet.write(4+i,3,class_a+str(r1+i+5));    
        for i in range(5) :
            worksheet.write(4+i,4,class_b+str(r2+i+5));

        for i in range(5) :
            worksheet.write(4+i,5,class_a+str(r1+i+10));
        for i in range(5) :
            worksheet.write(4+i,6,class_b+str(r2+i+10));

        for i in range(5) :
            worksheet.write(4+i,7,class_a+str(r1+i+15));
        for i in range(5) :
            worksheet.write(4+i,8,class_b+str(r2+i+15)); 
        worksheet.write(10,4,"TOOL CREATED BY SHUBHAM BHATIA :-)")                    
        wrk.close() ;
        messagebox.showinfo("SUCESS !","SUCESS RELOAD APPLICATION TO COMPILE NEW PLAN \n FILE SAVED WHERE YOU SAVED THIS APPLICATION")
        #code for TYPE2 is complete no need to edit otherwise it will mess up !;
        print("TYPE2")

# TYPE 3 ------------------------------------------------------------------------------------

    def TYPE3(self) :
        self.btn.configure(state=DISABLED);
       
        label1= Label(root,text="EXAM DATE",font=("Courier",10));
        label1.pack();
        self.date = StringVar();
        entry1 = Entry(root,textvariable=self.date);
        entry1.pack()
        #SET ROOM NAME :::::
        label2= Label(root,text="ROOM NAME:",font=("Courier",10));
        label2.pack();
        self.room_name = StringVar();
        entry2 = Entry(root,textvariable=self.room_name);
        entry2.pack();
        #SET CLASS 1
        label3= Label(root,text="CLASS 1st NAME",font=("Courier",10));
        label3.pack();
        self.class1 = StringVar();
        entry3 = Entry(root,textvariable=self.class1);
        entry3.pack()
        #STARTING ROLL FOR CLASS 1;
        label4= Label(root,text="STR ROLL CLASS 1st",font=("Courier",10));
        label4.pack();
        self.sr1 = StringVar();
        entry4 = Entry(root,textvariable=self.sr1);
        entry4.pack()
        #SET CLASS 2
        label5= Label(root,text="CLASS 2nd NAME",font=("Courier",10));
        label5.pack();
        self.class2 = StringVar();
        entry5 = Entry(root,textvariable=self.class2);
        entry5.pack()
        #STARTING ROLL FOR CLASS 2;
        label6= Label(root,text="STR ROLL CLASS 2nd",font=("Courier",10));
        label6.pack();
        self.sr2 = StringVar();
        entry6 = Entry(root,textvariable=self.sr2);
        entry6.pack()
        #SET CLASS  3RD #
        label10= Label(root,text="CLASS 3RD NAME",font=("Courier",10)).pack();
        self.class3 = StringVar();
        entry10 = Entry(root,textvariable=self.class3).pack();
        
        #STARTING ROLL FOR CLASS 3;
        label9= Label(root,text="STR ROLL CLASS 3RD",font=("Courier",10));
        label9.pack();
        self.sr3 = StringVar();
        entry9 = Entry(root,textvariable=self.sr3);
        entry9.pack()

         #SET CLASS 4
        label11= Label(root,text="CLASS 4TH NAME",font=("Courier",10)).pack();
        self.class4 = StringVar();
        entry11 = Entry(root,textvariable=self.class4).pack();
        #STARTING ROLL FOR CLASS 4TH;
        label12= Label(root,text="STR ROLL CLASS 4TH",font=("Courier",10)).pack();
        self.sr4 = StringVar();
        entry12 = Entry(root,textvariable=self.sr4).pack();
        
        # file name ::::::
        label8= Label(root,text="ENTER FILE NAME : ",font=("Courier",10));
        label8.pack();
        self.file = StringVar();
        entry8 = Entry(root,textvariable=self.file);
        entry8.pack()
        # COMPILE BUTTON : 
        Button(root,text="COMPILE PLAN :-)" ,command=self.compile3).pack();
       
    def compile3(self) :
        # zaruri jankeri jo chahiye 
        dated = self.date.get();
        room = self.room_name.get();
        class_a = str(self.class1.get())+" ";
        class_b = str(self.class2.get())+" ";
        class_c = str(self.class3.get())+" ";
        class_d = str(self.class4.get())+" ";

        roll_a = int(self.sr1.get());
        roll_b = int(self.sr2.get());
        roll_c = int(self.sr4.get());
        roll_d = int(self.sr2.get());
        fn = self.file.get();
        filen = fn+".xlsx"
        # ye exel wala idea amit se btw tq amit :-)
        wrk = xlsxwriter.Workbook(filen);
        worksheet = wrk.add_worksheet();
        worksheet.write(0,4,"JAWAHAR NAVODAYA VIDYALAYA SEATING PLAN FOR ROOM : "+str(room))
        worksheet.write(1,1,"EXAM DATE: "+str(dated));
        worksheet.write(1,5,"ROOM :"+str(room));
         
        # SETTING ROWS 
        for i in range(8):
           worksheet.write(3,i+1,"ROW "+str(i+1));
       
        for i in range(5):
            worksheet.write(4+i,0,"COL:"+str(i+1));
         # yaha code likho upar variable diye ve unko use karke ;
        
        #class1
        worksheet.write('B5', (class_a+(str(roll_a))))
        worksheet.write('B7', (class_a+(str(roll_a+1))))
        worksheet.write('B9', (class_a+(str(roll_a+2))))
        worksheet.write('D6', (class_a+(str(roll_a+3))))
        worksheet.write('D8', (class_a+(str(roll_a+4))))
        worksheet.write('F5', (class_a+(str(roll_a+5))))
        worksheet.write('F7', (class_a+(str(roll_a+6))))
        worksheet.write('F9', (class_a+(str(roll_a+7))))
        worksheet.write('H6', class_a+(str(roll_a+8)))
        worksheet.write('H8', (class_a+(str(roll_a+9))))

        #CLASS2
        worksheet.write('B6', (class_b+(str(roll_b))))
        worksheet.write('B8', (class_b+(str(roll_b+1))))
        worksheet.write('D5', (class_b+(str(roll_b+2))))
        worksheet.write('D7', (class_b+(str(roll_b+3))))
        worksheet.write('D9', (class_b+(str(roll_b+4))))
        worksheet.write('F6', (class_b+(str(roll_b+5))))
        worksheet.write('F8', (class_b+(str(roll_b+6))))
        worksheet.write('H5', (class_b+(str(roll_b+7))))
        worksheet.write('H7', (class_b+(str(roll_b+8))))
        worksheet.write('H9', (class_b+(str(roll_b+9))))

        #class 3
        worksheet.write('C5', (class_c+(str(roll_c))))
        worksheet.write('C7', (class_c+(str(roll_c+1))))
        worksheet.write('C9', (class_c+(str(roll_c+2))))
        worksheet.write('E6', (class_c+(str(roll_c+3))))
        worksheet.write('E8', (class_c+(str(roll_c+4))))
        worksheet.write('G5', (class_c+(str(roll_c+5))))
        worksheet.write('G7', (class_c+(str(roll_c+6))))
        worksheet.write('G9', (class_c+(str(roll_c+7))))
        worksheet.write('I6', (class_c+(str(roll_c+8))))
        worksheet.write('I8', (class_c+(str(roll_c+9))))

        #class 4
        worksheet.write('C6',(class_d+(str(roll_d))))
        worksheet.write('C8',(class_d+(str(roll_d+1))))
        worksheet.write('E5',(class_d+(str(roll_d+2))))
        worksheet.write('E7',(class_d+(str(roll_d+3))))
        worksheet.write('E9',(class_d+(str(roll_d+4))))
        worksheet.write('G6',(class_d+(str(roll_d+5))))
        worksheet.write('G8',(class_d+(str(roll_d+6))))
        worksheet.write('I5',(class_d+(str(roll_d+7))))
        worksheet.write('I7',(class_d+(str(roll_d+8))))
        worksheet.write('I9',(class_d+(str(roll_b+9))))


        wrk.close();
        messagebox.showinfo("SUCESS !","SUCESS RELOAD APPLICATION TO COMPILE NEW PLAN \n FILE SAVED WHERE YOU SAVED THIS APPLICATION")
    
        #code for TYPE5
        print("TYPE3")  

  # TYPE 4  ---------------------------------------------------------------
    
    def TYPE4(self):
        self.btn.configure(state=DISABLED);
         #SET date
        label1= Label(root,text="EXAM DATE",font=("Courier",10));
        label1.pack();
        self.date = StringVar();
        entry1 = Entry(root,textvariable=self.date);
        entry1.pack()
        #SET ROOM NAME :::::
        label2= Label(root,text="ROOM NAME:",font=("Courier",10));
        label2.pack();
        self.room_name = StringVar();
        entry2 = Entry(root,textvariable=self.room_name);
        entry2.pack();
        #SET CLASS 1
        label3= Label(root,text="CLASS 1st NAME",font=("Courier",10));
        label3.pack();
        self.class1 = StringVar();
        entry3 = Entry(root,textvariable=self.class1);
        entry3.pack()
        #STARTING ROLL FOR CLASS 1;
        label4= Label(root,text="STR ROLL CLASS 1st",font=("Courier",10));
        label4.pack();
        self.sr1 = StringVar();
        entry4 = Entry(root,textvariable=self.sr1);
        entry4.pack()
        #SET CLASS 2
        label5= Label(root,text="CLASS 2nd NAME",font=("Courier",10));
        label5.pack();
        self.class2 = StringVar();
        entry5 = Entry(root,textvariable=self.class2);
        entry5.pack()
        #STARTING ROLL FOR CLASS 2;
        label6= Label(root,text="STR ROLL CLASS 2nd",font=("Courier",10));
        label6.pack();
        self.sr2 = StringVar();
        entry6 = Entry(root,textvariable=self.sr2);
        entry6.pack();
        # file name ::::::
        label8= Label(root,text="ENTER FILE NAME : ",font=("Courier",10));
        label8.pack();
        self.file = StringVar();
        entry8 = Entry(root,textvariable=self.file);
        entry8.pack()
        # COMPILE BUTTON : 
        Button(root,text="COMPILE PLAN :-)" ,command=self.compile4).pack();

    def compile4(self) :
        # zaruri jankeri jo chahiye 
        dated = self.date.get();
        room = self.room_name.get();
        class_a = str(self.class1.get())+" ";
        class_b = str(self.class2.get())+" ";
       
        roll_a = int(self.sr1.get());
        roll_b = int(self.sr2.get());
       
        fn = self.file.get();
        filen = fn+".xlsx"
        # ye exel wala idea amit se btw tq amit :-)
        wrk = xlsxwriter.Workbook(filen);
        worksheet = wrk.add_worksheet();
        worksheet.write(0,4,"JAWAHAR NAVODAYA VIDYALAYA SEATING PLAN FOR ROOM : "+str(room))
        worksheet.write(1,1,"EXAM DATE: "+str(dated));
        worksheet.write(1,5,"ROOM :"+str(room));
         
        # SETTING ROWS 
        for i in range(8):
           worksheet.write(3,i+1,"ROW "+str(i+1));
       
        for i in range(5):
            worksheet.write(4+i,0,"COL:"+str(i+1));

         # yaha code likho upar variable diye ve unko use karke ;




         
        wrk.close();
        
        messagebox.showinfo("SUCCESS !","SUCCESS RELOAD APPLICATION TO COMPILE NEW PLAN \n FILE SAVED WHERE YOU SAVED THIS APPLICATION")
       
        
App(root);
root.mainloop();