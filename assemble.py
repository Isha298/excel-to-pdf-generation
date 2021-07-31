from tkinter import *
from tkinter import messagebox
# Reading an excel file using Python
import xlrd
from fpdf import FPDF
#tkinker checkbox,button

class Test:
    def __init__(self, tk):
        text = Label(tk, text="Please! select student no. to get their PDF.").place(x=50, y=20)
        self.checkvar1 = IntVar()
        self.checkvar2 = IntVar()
        self.checkvar3 = IntVar()
        self.checkvar4 = IntVar()
        self.checkvar5 = IntVar()

        self.chkbtn1 = Checkbutton(tk, text="Student 1", variable=self.checkvar1).place(x=50, y=50)
        self.chkbtn2 = Checkbutton(tk, text="Student 2", variable=self.checkvar2).place(x=50, y=80)
        self.chkbtn3 = Checkbutton(tk, text="Student 3", variable=self.checkvar3).place(x=50, y=110)
        self.chkbtn4 = Checkbutton(tk, text="Student 4", variable=self.checkvar4).place(x=50, y=140)
        self.chkbtn5 = Checkbutton(tk, text="Student 5", variable=self.checkvar5).place(x=50, y=170)
        self.btn1 = Button(tk, text="Submit", command=self.submit).place(x=50, y=250)

    def submit(self):
        list_of_choices=[]
        if self.checkvar1.get():
            list_of_choices.append(1)
        if self.checkvar2.get():
            list_of_choices.append(2)
        if self.checkvar3.get():
            list_of_choices.append(3)
        if self.checkvar4.get():
            list_of_choices.append(4)
        if self.checkvar5.get():
            list_of_choices.append(5)
        create_pdf(list_of_choices)
        messagebox.showinfo("Congrats!","Result Published")


def f3(sheet,student_no): 
 l=[[],[],[],[],[],[]]
 student_no-=1
 for e2 in range(2+25*student_no,27+25*student_no):
   l[0].append(str(sheet.cell_value(e2,13)))
   l[1].append(str(sheet.cell_value(e2,14)))
   l[2].append(str(sheet.cell_value(e2,15)))
   l[3].append(str(sheet.cell_value(e2,16)))
   l[4].append(int(sheet.cell_value(e2,17)))
   l[5].append(int(sheet.cell_value(e2,18)))
 return l


def f1(sheet,student_no):
 l=[]
 student_no-=1
 row=2+25*student_no
 l.append(str(int(sheet.cell_value(row,5))))
 l.append(str(sheet.cell_value(row,4)))
 l.append(str(sheet.cell_value(row,8)))
 strv=str(xlrd.xldate_as_datetime(sheet.cell_value(row, 9),0).date())
 l.append("-".join(strv.split("-")[::-1]))
 l.append(str(int(sheet.cell_value(row,6))))
 l.append(str(sheet.cell_value(row,7)))
 l.append(str(sheet.cell_value(row,10)))
 l.append(str(sheet.cell_value(row,12)))
 return l

def create_pdf(list_of_choices):
    
    # Give the location of the file
    loc =("C:\\Users\\asus\\Downloads\\Dummy Data (1).xls")

    # To open Workbook
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    choice=0
    count=0
    while count<len(list_of_choices):
        #event handling
        #choice=eval(input("enter your choice between 1 to 5\n"))
        choice=list_of_choices[count]
        count+=1

        # save FPDF() class into a
        # variable pdf
        pdf = FPDF()

        # Add a page
        pdf.add_page()

        # set style and size of font
        # that you want in the pdf
        pdf.set_font('Arial','B',28);

        # create a cell
        pdf.cell(230, 10, txt= "NATIONAL EDUCATION TEST",
                 ln = 1, align = 'C')
        pdf.cell(230,10,"NET 2021",
                 ln = 2, align = 'C')

        pdf.set_font('Arial','BIU',12);
        # add another cell
        pdf.cell(230, 10, txt = "(Education in our hands)",
                        ln = 3, align = 'C')
        #adding image
        pdf.image('C:/Users/asus/Downloads/cropped.jpg',6,6,35,35,'JPG')
        #line decoration
        pdf.set_line_width(1)
        pdf.set_draw_color(0, 20, 180)
        pdf.line(0, 50, 600, 50)

        #details
        pdf.set_font('Arial','BU',20);
        pdf.cell(190,35, txt = "Round-2  RESULTS",
                 ln = 2, align = 'C')
        pdf.set_font('Arial','B',14);
        pdf.cell(10,8, txt = "Date and Time: Aug 5-6 2021",
                 ln = 2, align = 'L')
        if(choice==1):
            pdf.image('C:/Users/asus/Downloads/ABC1 XYZ1.jpg',155,83,45,50,'JPG')
        elif(choice==2):
            pdf.image('C:/Users/asus/Downloads/ABC2 XYZ2.jpg',155,83,45,50,'JPG')
        elif(choice==3):
            pdf.image('C:/Users/asus/Downloads/ABC3 XYZ3.jpg',155,83,45,50,'JPG')
        elif(choice==4):
            pdf.image('C:/Users/asus/Downloads/ABC4 XYZ4.jpg',155,83,45,50,'JPG')
        elif(choice==5):
            pdf.image('C:/Users/asus/Downloads/ABC5 XYZ5.jpg',155,83,45,50,'JPG')
            
        epw = pdf.w - 2*pdf.l_margin
        col_width = epw/2.8
        th=pdf.font_size
        pdf.set_line_width(0.6)
        pdf.set_draw_color(0, 0, 0)
        l=f1(sheet,choice)

        data1= [['NET Registration No.',l[0]],
                    ["Student's Full Name:",l[1]],['Gender:',l[2]],['Date of Birth :',l[3]],
                    ['Grade:',l[4]],['Name of School:',l[5]],
                    ['City of Residence:',l[6]],['Country of Residence:',l[7]]]

        pdf.set_font('Times','B',14)
        for row in data1:
            for lstum in row:
                pdf.cell(col_width,2.3*th, str(lstum), border=1)
                
            pdf.ln(2.3*th)

        pdf.set_font('Times','BU',14)
        pdf.cell(30,20, txt = "MARKS DISTRIBUTION:- ",
                 ln = 13, align = 'L')

        pdf.set_font('Arial','',14.0);
        # table entries
        epw = pdf.w - 2*pdf.l_margin
        col_width = epw/4
        th=pdf.font_size
        pdf.set_line_width(0.6)
        pdf.set_draw_color(0, 0, 0)

        marks_obtained=0
        for x in range(2+25*(choice-1),27+25*(choice-1)):
               marks_obtained=marks_obtained+int(sheet.cell_value(x,18))
               
        if(marks_obtained>80):
            Grades='AA'
        elif(marks_obtained>70):
            Grades='BB'
        elif(marks_obtained>60):
            Grades='CC'
        elif(marks_obtained>50):
            Grades='DD'
        else:
            Grades='FF'

        data2= [['Paper(NET 2021)','Maximum Marks','Marks Obtained','Grades'],
                    [' ',100,marks_obtained,Grades]]
            
        pdf.set_font('Times','B',10.0)
        for row in data2:
            for datum in row:
                pdf.cell(col_width,2*th, str(datum), border=1)
                
            pdf.ln(2*th)
            
        pdf.cell(25,15,txt="Note**",
                 ln=15,align='L')
        pdf.cell(5,12,txt='1. In case candidate appeared in Round-1 also, must inform us via mail.',
                 ln=15, align='L')
        pdf.cell(5,12,txt='2. There was no negative marking in the Exam paper.',
                 ln=15,align='L')
        pdf.cell(5,12,txt='3. Obtained Marks are marks secured by individual in NET Exam.',
                 ln=15,align='L')
        # Add a page
        pdf.add_page()
        pdf.set_font('Times','BU',10)
        pdf.cell(8,10,txt="MARKS ANALYSIS :-",
                 ln=2,align='L')

        #SCORECARDS
        epw = pdf.w - 2*pdf.l_margin
        col_width = epw/6
        th=pdf.font_size
        data3 = [['Question No.','Student Ans','Correct Ans','Outcome','Score if Correct','Your Score']]

        l=f3(sheet,choice)#student no

        for e in range(0,25):
                data3.append([l[0][e] , l [1] [e] , l[2] [e] , l [3] [e] , l [4] [e] , l [5] [e] ] )

        pdf.set_font('Times','B',10.0)
        for row in data3:
            for datum in row:
                pdf.cell(col_width,2*th, str(datum), border=1)
                
            pdf.ln(2*th)
        #REMARKS
        a=str(sheet.cell_value(2,19))
        b=str(sheet.cell_value(27,19))
        c=str(sheet.cell_value(52,19))
        d=str(sheet.cell_value(77,19))
        e=(sheet.cell_value(102,19))
        if(choice==1):
            value=a
        elif(choice==2):
            value=b
        elif(choice==3):
            value=c
        elif(choice==4):
            value=d
        elif(choice==5):
            value=e
        pdf.set_font('Times','BU',20)
        pdf.cell(10,30,txt="Remarks:-",
                 ln=30,align='L')
        pdf.set_font('Times','B',10.0)
        epw = pdf.w - 2*pdf.l_margin
        col_width = epw/1.5
        th=pdf.font_size
        pdf.cell(col_width,2*th,value, border=1)

        # save the pdf with name .pdf
        var=str(int(sheet.cell_value(2+25*(choice-1),5)))+".pdf"
        pdf.output(var)
        print("pdf has been created for Student",choice )
#main program
tk = Tk()
tk.geometry("600x500")
tk.title("Results")
myTest = Test(tk)
tk.mainloop()
quit()
