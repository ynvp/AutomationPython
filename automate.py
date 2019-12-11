# Selenium Module(Automates Browser)
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium import webdriver
import csv
from tkinter import *
import tkinter.messagebox
import requests,os,time
from zipfile import ZipFile

window = Tk()
window.title('Gitam Univeristy Semester Result')
window.geometry('500x520')
window.resizable(0,0)

live_status = StringVar()
up_frame = Frame(window, bg='black', height=20, width=500)
up_frame.pack(side=TOP, fill=X)

bottom_frame = Frame(window,height=20,width=500)
bottom_frame.pack(side=BOTTOM,fill=X)

frame2 = Frame(window, height=480, width=500,bg="gray")
frame2.pack(fill=X)

status = Label(bottom_frame,bd=1,relief=SUNKEN,anchor=W,textvariable=live_status)
status.pack(fill=X)

path = 'GITAM_logo.png'
if os.path.exists(path):
    pass
else:
    r=requests.get('https://www.gitam.edu/assets/images/GITAM_logo.png')
    with open('GITAM_logo.png','wb') as f:
        f.write(r.content)

photo=PhotoImage(file=path)
photo_label=Label(up_frame,image=photo,width=500,height=150)
photo_label.pack()



semester_label = Label(frame2, text="Semester   :", bg="white", fg="black",font="consolas 15")
section_label = Label(frame2, text="Section    :", bg="white", fg="black",font="consolas 15")


semester_entry= IntVar()
section_entry= IntVar()

semester_entry = Entry(frame2)
section_entry = Entry(frame2)

semester_label.place(x=100, y=95)
section_label.place(x=100, y=145)

semester_entry.place(x=250, y=100)
section_entry.place(x=250, y=150)

gpa=[]
roll=[]
name=[]


def search():
    maxStudents = 0
    passPercent = 0
    avgCGPA = 0
    #Checking for empty fields
    if semester_entry.get() == '':
        tkinter.messagebox.showinfo('Error', "Semester value cannot be NULL.")
        return
    elif section_entry.get() == '':
        tkinter.messagebox.showinfo('Error', "Section value cannot be NULL.")
        return

    #else continue to conversion
    section_ip1 = int(section_entry.get())
    semester_ip1 =int(semester_entry.get())

    #check if any value exceeds it's limits
    if semester_ip1>8 :
        tkinter.messagebox.showinfo('Error',"Semester value cannot be greater than 8")
        semester_entry.delete(0, 'end')
        return
    elif section_ip1>19 :
        tkinter.messagebox.showinfo('Error',"Section value cannot be greater than 19")
        section_entry.delete(0,'end')
        return

    sec = section_ip1
    sem = semester_ip1

    tkinter.messagebox.showinfo('NOTE','Running Chrome to get Results.Click --OK-- to proceed.')
    tkinter.messagebox.showinfo('NOTE','Please wait while we fetch data and record it. \nIMP: Donot close CHROME during operation.')
    # Initializing webdriver with its origin
    path = 'chromedriver.exe'
    if os.path.exists(path):
        pass
    else:
        live_status.set('Downloading chromedriver for automation...')
        window.update()
        time.sleep(5)
        r = requests.get('https://chromedriver.storage.googleapis.com/75.0.3770.90/chromedriver_win32.zip')
        with open('chromedriver.zip','wb') as f:
            f.write(r.content)
        live_status.set('Extracting chromedriver')
        window.update()
        time.sleep(5)
        zipfile = ZipFile('chromedriver.zip','r')
        zipfile.extractall()
        live_status.set('Extracted chromedriver, Opening chrome...')
        window.update()
    browser = webdriver.Chrome('chromedriver')

    # Automate to go to http address
    browser.get("https://eweb.gitam.edu/mobile/Pages/NewGrdcrdInput1.aspx")

    for i in range(1,70):
        window.update()  # Roll loop
        if i == 69:
            browser.close()
            passPercent /=maxStudents
            passPercent *=100
            avgCGPA /= maxStudents
            tkinter.messagebox.showinfo('Results','Students strength : '+str(maxStudents)+'\nPassPercentage: '+str(round(passPercent,2))+'\nAverege CGPA: '+str(round(avgCGPA,2)))
            tkinter.messagebox.showinfo('NOTE',"The program has completed it's task, You can please press quit button to close the session")
            break
        semester = Select(browser.find_element_by_id("cbosem"))
        # Selecting semester option with Select class from Selenium
        semester.select_by_value(str(sem))
        # select_by_value selects semester with given value
        searchBar = browser.find_element_by_id("txtreg")
        # find_element_by_id finds element and equalises it to other object
        if i < 10 and sec< 10:
            # send_keys sends string to a form in HTML
            searchBar.send_keys("12171030"+str(sec)+"00"+str(i))
        elif i < 10 and sec>= 10:
            searchBar.send_keys("1217103"+str(sec)+"00"+str(i))
        elif i >= 10 and sec < 10:
            searchBar.send_keys("12171030"+str(sec)+"0"+str(i))
        elif i >= 10 and sec>= 10:
            searchBar.send_keys("1217103"+str(sec)+"0"+str(i))
        elem = browser.find_element_by_id('Button1')
        elem.click()  # click() fun automates button clicks
        try:
            WebDriverWait(browser, 0.1).until(EC.alert_is_present(),
                                                'Timed out waiting for PA creation ' +
                                                'confirmation popup to appear.')
            # wait 0.1 sec to catch an alert
            alert = browser.switch_to.alert
            alert.accept()
            print("Failed")

            # automates and accepts alert by clicking "OK" button
            gpa.append("NO_GPA_Recorded")
            roll.append("Page_Error")
            name.append("Page_Error")
            continue  # continues to next roll no.
        except TimeoutException:
            print("Successful!")
        Name = browser.find_element_by_id("lblname")
        marks = browser.find_element_by_id("lblgpa")
        regNo = browser.find_element_by_id("lblregdno")
        # takes text of the HTML element and assigns it to variable
        cgpa = marks.text
        gpa.append(cgpa)
        roll.append(regNo.text)
        name.append(Name.text)
        back_elem = browser.find_element_by_id('Button1')
        back_elem.click()
        live_status.set('Extracted Roll no '+str(i))
        maxStudents = i
        if  float(cgpa)> 0:
            passPercent +=1
        avgCGPA += float(cgpa)
        # creates data_all.csv and makes it read and write enable
        with open("data_"+str(sec)+".csv", "w+") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Name", "Roll No", "Marks"])
            for i in range(len(gpa)):
                # Writing fetched data into csv file
                writer.writerow([name[i], str(roll[i]), str(gpa[i])])


#Submit_Button
submit_button = Button(frame2, text="Submit", fg="green", bg="white", command=search,height=1,width=8,font="consolas 15 bold")
submit_button.place(x=200,y=210)
def thank():
    tkinter.messagebox.showinfo('Thank Note','Thank You For Using the Product')
    exit()
quit_button=Button(text="Quit",fg="white",bg="red",command=thank,height=1,width=8,font="consolas 12 italic")
quit_button.place(x=400,y=450)

window.mainloop()
