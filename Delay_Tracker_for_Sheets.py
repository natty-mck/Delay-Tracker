import datetime
from re import split
import openpyxl
import tkinter as tk
from tkinter import filedialog as fd
import customtkinter
from tkcalendar import Calendar
from datetime import date
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg,
    NavigationToolbar2Tk
)
import ctypes
import babel.numbers


matplotlib.use("TkAgg")
class info:
    current_time = datetime.now()
    year = current_time.year
    month = current_time.month
    day = current_time.day
    start_jobs = 19
    end_jobs = 339
    job_column = "G"



#open the spread sheet
def mainmenu():
    global spreadsheet_button, root, spread_upload, loading_caption

    bg_colour = '#ADADAD'
    root = customtkinter.CTk(fg_color= bg_colour)

    root.title('Delay Tracker')


    root.state("zoomed")

    
    customtkinter.set_window_scaling(1)
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    

    spreadsheet_path = ""

    spread_upload = customtkinter.CTkFrame(root, height= root.winfo_height() , width= root.winfo_width())
    spread_upload.pack()

    root.update()

    spreadsheet_caption = customtkinter.CTkLabel(text = "Spreadsheet Upload", master = spread_upload, font = ("Arial", 18))
    spreadsheet_caption.place(relx = 0.5, rely = 0.3, anchor = "center" )
    
    spreadsheet_button = customtkinter.CTkButton(master = spread_upload, text='Upload Spreadsheet', command=openspreadsheet)
    spreadsheet_button.place(relx = 0.5, rely = 0.4, anchor = "center")
    
    submit_button = customtkinter.CTkButton(master = spread_upload, text='Submit', command= submit_spreadsheet)
    submit_button.place(relx = 0.5, rely =0.6, anchor = "center" )


    root.mainloop()


def submit_spreadsheet():
    

    global spread_info, workbook, sheet, job_cell_cords, job_names, job_names_dict, job_choice_var, submit_button, start_cal, end_cal

    workbook = openpyxl.load_workbook(spreadsheet_path)
    sheet = workbook.active 

    spread_upload.destroy()

    spread_info = customtkinter.CTkFrame(root, height= 200 , width= root.winfo_width())
    spread_info.pack()

    job_cell_cords = []
    job_names = []
    for jobs in range(info.start_jobs, info.end_jobs, 5):
        job_cell_cords.append(jobs)
    
    for jobrow in job_cell_cords:
        job_names.append(sheet[info.job_column+str(jobrow)].value)

    job_names = [x for x in job_names if x != None]
    job_names_dict = {}
    for i in range(len(job_names)):
        job_names_dict[job_names[i]] = job_cell_cords[i]

    print(job_names_dict)
    
    spreadsheet_caption = customtkinter.CTkLabel(text = "Select Job", master = spread_info, font = ("Arial", 18))
    spreadsheet_caption.place(relx = 0.5, rely = 0.2, anchor = "center" )
    
    job_choice_var = customtkinter.StringVar(value=job_names[0])
    job_choice = customtkinter.CTkOptionMenu(spread_info,values=job_names, variable=job_choice_var)
    job_choice.place(relx = 0.5, rely =0.5, anchor = "center" )
    
    submit_button = customtkinter.CTkButton(master = spread_info, text='Analyse', command= submit_choice)
    submit_button.place(relx = 0.5, rely = 0.7, anchor = "center" )

    start_cal_caption = customtkinter.CTkLabel(text = "Select Start Date:", master = spread_info, font = ("Arial", 18))
    start_cal_caption.place(relx = 0.15, rely = 0.5, anchor = "center")
    start_cal = Calendar(spread_info, selectmode = 'day',
               year = info.year, month = 1,
               day = 1)
 
    start_cal.place(relx = 0.3,rely = 0.5, anchor = "center" )

    
    end_cal_caption = customtkinter.CTkLabel(text = "Select End Date:", master = spread_info, font = ("Arial", 18))
    end_cal_caption.place(relx = 0.65, rely = 0.5, anchor = "center" )
    end_cal = Calendar(spread_info, selectmode = 'day',
               year = info.year, month = info.month,
               day = info.day)
 
    end_cal.place(relx = 0.8,rely = 0.5, anchor = "center" )




def refresh():
    choice_details.destroy()
    graph_holder.destroy()
    submit_choice()



def submit_choice():


    
    start_date = start_cal.get_date()
    end_date = end_cal.get_date()

    start_date_obj = datetime.strptime(start_date, "%m/%d/%y")
    end_date_obj = datetime.strptime(end_date, "%m/%d/%y")

    date0 = datetime(info.year,1,1)
    date1 = datetime(info.year,start_date_obj.month,start_date_obj.day)
    date2 = datetime(info.year,end_date_obj.month,end_date_obj.day)


    start_difference = date1 - date0
    selected_difference = date2 - date1


    global choice_details

    submit_button.configure(command = refresh)

    choice_details = customtkinter.CTkFrame(root, height= root.winfo_height() - 200  , width= root.winfo_width()/2 -600 , fg_color="#7D9EC0")
    choice_details.pack(side = "left")
    
    total_delays = 0
    start_cell = sheet['B2']

    design_delays = 0
    fab_delays = 0
    mech_delays = 0
    elect_delays = 0
    loss_of_delays = 0
    loss_of_delays_design = 0
    loss_of_delays_mech = 0
    loss_of_delays_elect = 0
    loss_of_delays_fab = 0
    cost_of_delay = 465
    people_delays = 0
    for i in range(int(str(start_difference.days+1)), (int(str(selected_difference.days))+1+int(str(start_difference.days))),1):
        cell = sheet["S" + str(job_names_dict[job_choice_var.get()])]
        design_cell = cell.offset(row=1, column=i)
        fab_cell = cell.offset(row=2, column=i)
        mech_cell = cell.offset(row=3, column=i)
        elect_cell = cell.offset(row=4, column=i)

        design_value = list(str(design_cell.value))
        fab_value = list(str(fab_cell.value))
        mech_value = list(str(mech_cell.value))
        elect_value = list(str(elect_cell.value))


        numeric = []
        alpha = []
        for item in design_value:
            try:
                int(item)
                numeric.append(item)
            except:
                alpha.append(item)
         
        numeric = "".join(numeric)
        alpha= "".join(alpha)
        design_value = [numeric, alpha]


        numeric = []
        alpha = []
        for item in fab_value:
            try:
                int(item)
                numeric.append(item)
            except:
                alpha.append(item)
         
        numeric = "".join(numeric)
        alpha= "".join(alpha)
        fab_value = [numeric, alpha]


        numeric = []
        alpha = []
        for item in mech_value:
            try:
                int(item)
                numeric.append(item)
            except:
                alpha.append(item)
         
        numeric = "".join(numeric)
        alpha= "".join(alpha)
        mech_value = [numeric, alpha]


        numeric = []
        alpha = []
        for item in elect_value:
            try:
                int(item)
                numeric.append(item)
            except:
                alpha.append(item)
         
        numeric = "".join(numeric)
        alpha= "".join(alpha)
        elect_value = [numeric, alpha]




        for i in range(len(design_value)):
            if design_value[i] == "d" or design_value[i] == "D":
                design_delays = design_delays + 1
                try:
                    loss_of_delays = loss_of_delays + int(design_value[i-1]) * cost_of_delay
                    loss_of_delays_design = loss_of_delays_design + int(design_value[i-1]) * cost_of_delay
                    people_delays = people_delays + int(design_value[i-1])


                except:
                    loss_of_delays = loss_of_delays + cost_of_delay
                    loss_of_delays_design = loss_of_delays_design + cost_of_delay
                    people_delays = people_delays + 1 


        for i in range(len(fab_value)):
            if fab_value[i] == "d" or fab_value[i] == "D":
                fab_delays = fab_delays + 1
                try:
                    loss_of_delays = loss_of_delays + int(fab_value[i-1]) * cost_of_delay
                    loss_of_delays_fab = loss_of_delays_fab + + int(fab_value[i-1]) * cost_of_delay
                    people_delays = people_delays + int(fab_value[i-1])

                except:
                    loss_of_delays = loss_of_delays + cost_of_delay
                    loss_of_delays_fab = loss_of_delays_fab + cost_of_delay
                    people_delays = people_delays + 1


        for i in range(len(mech_value)):
            if mech_value[i] == "d" or mech_value[i] == "D":
                mech_delays = mech_delays + 1
                try:
                    loss_of_delays = loss_of_delays + int(mech_value[i-1]) * cost_of_delay
                    loss_of_delays_mech = loss_of_delays_mech + int(mech_value[i-1]) * cost_of_delay
                    people_delays = people_delays + int(mech_value[i-1])

                except:
                    loss_of_delays = loss_of_delays + cost_of_delay
                    loss_of_delays_mech = loss_of_delays_mech + cost_of_delay
                    people_delays = people_delays + 1


        for i in range(len(elect_value)):
            if elect_value[i] == "d" or elect_value[i] == "D":
                elect_delays = elect_delays + 1
                try:
                    loss_of_delays = loss_of_delays + int(elect_value[i-1]) * cost_of_delay
                    loss_of_delays_elect = loss_of_delays_elect + int(elect_value[i-1]) * cost_of_delay
                    people_delays = people_delays + int(elect_value[i-1])

                except:
                    loss_of_delays = loss_of_delays + cost_of_delay
                    loss_of_delays_elect = loss_of_delays_elect + cost_of_delay
                    people_delays = people_delays + 1



    total_delays = design_delays + fab_delays + mech_delays + elect_delays

    total_delays_caption = customtkinter.CTkLabel(text = str("Total delays for job= " + str(total_delays) + "(days)")  , master = choice_details, font = ("Arial", 18))
    total_delays_caption.place(relx = 0.96, rely = 0.1, anchor = "e" )

    total_people_delays_caption = customtkinter.CTkLabel(text = str("Delays per person= " + str(people_delays) + "(man days)")  , master = choice_details, font = ("Arial", 18))
    total_people_delays_caption.place(relx = 0.96, rely = 0.15, anchor = "e" )

    design_delays_caption = customtkinter.CTkLabel(text = str("Design delays for job= " + str(design_delays) + "(days)")  , master = choice_details, font = ("Arial", 18))
    design_delays_caption.place(relx = 0.96, rely = 0.25, anchor = "e" )

    fab_delays_caption = customtkinter.CTkLabel(text = str("Fabrication delays for job= " + str(fab_delays) + "(days)")  , master = choice_details, font = ("Arial", 18))
    fab_delays_caption.place(relx = 0.96, rely = 0.3, anchor = "e" )

    mech_delays_caption = customtkinter.CTkLabel(text = str("Mechanical delays for job= " + str(mech_delays) + "(days)")  , master = choice_details, font = ("Arial", 18))
    mech_delays_caption.place(relx = 0.96, rely = 0.35, anchor = "e" )

    elect_delays_caption = customtkinter.CTkLabel(text = str("Electrical delays for job= " + str(elect_delays) + "(days)")  , master = choice_details, font = ("Arial", 18))
    elect_delays_caption.place(relx = 0.96, rely = 0.4, anchor = "e" )


    delays_cost_caption = customtkinter.CTkLabel(text = "Total cost of delays of job= \u00A3" + str(loss_of_delays)  , master = choice_details, font = ("Arial", 18))
    delays_cost_caption.place(relx = 0.96, rely = 0.5, anchor = "e" )

    design_delays_cost_caption = customtkinter.CTkLabel(text = "Cost of Design delays of job= \u00A3" + str(loss_of_delays_design)  , master = choice_details, font = ("Arial", 18))
    design_delays_cost_caption.place(relx = 0.96, rely = 0.55, anchor = "e" )

    fab_delays_cost_caption = customtkinter.CTkLabel(text = "Cost of Fabrication delays of job= \u00A3" + str(loss_of_delays_fab)  , master = choice_details, font = ("Arial", 18))
    fab_delays_cost_caption.place(relx = 0.96, rely = 0.6, anchor = "e" )

    mech_delays_cost_caption = customtkinter.CTkLabel(text = "Cost of Mechanical delays of job= \u00A3" + str(loss_of_delays_mech)  , master = choice_details, font = ("Arial", 18))
    mech_delays_cost_caption.place(relx = 0.96, rely = 0.65, anchor = "e" )

    elect_delays_cost_caption = customtkinter.CTkLabel(text = "Cost of Electrical delays of job= \u00A3" + str(loss_of_delays_elect)  , master = choice_details, font = ("Arial", 18))
    elect_delays_cost_caption.place(relx = 0.96, rely = 0.7, anchor = "e" )

    note_caption = customtkinter.CTkLabel(text = "Note: Each person dealyed is \n equivilant to \u00A3" + str(cost_of_delay)  , master = choice_details, font = ("Arial", 18))
    note_caption.place(relx = 0.96, rely = 0.8, anchor = "e" )
    

    graph()


    root.mainloop()

def graph():
    global graph_holder
    weeks = []
    for i in range(1,53,1):
        weeks.append("week" + str(i))

    elect_total = []
    mech_total = []
    fab_total = []
    design_total = []
    weekly_total = []
    for x in range(0,364,7):
        design_delays = 0
        fab_delays = 0
        mech_delays = 0
        elect_delays = 0
        for y in range(x, x+7, 1):
            cell = sheet["S" + str(job_names_dict[job_choice_var.get()])]
            design_cell = cell.offset(row=1, column=y)
            fab_cell = cell.offset(row=2, column=y)
            mech_cell = cell.offset(row=3, column=y)
            elect_cell = cell.offset(row=4, column=y)

            print(design_cell)

            if "d" in list(str(design_cell.value)) or "D" in list(str(design_cell.value)):
                design_delays = design_delays + 1

            if "d" in list(str(fab_cell.value)) or "D" in list(str(fab_cell.value)):
                fab_delays = fab_delays + 1

            if "d" in list(str(mech_cell.value)) or "D" in list(str(mech_cell.value)):
                mech_delays = mech_delays + 1

            if "d" in list(str(elect_cell.value)) or "D" in list(str(elect_cell.value)):
                elect_delays = elect_delays + 1

        elect_total.append(elect_delays)
        mech_total.append(mech_delays)
        fab_total.append(fab_delays)
        design_total.append(design_delays)
        weekly_total.append(design_delays + fab_delays + mech_delays + elect_delays)

    print(weeks)
    print(weekly_total)
    print(len(weeks))
    print(len(weekly_total))
    print(elect_total)




    graph_holder = customtkinter.CTkFrame(root, height= root.winfo_height() - 200 ,  width = root.winfo_width()/2 + 600)
    graph_holder.pack(side = "right")

    figure = Figure(figsize=(12, 8), dpi=200)

    figure_canvas = FigureCanvasTkAgg(figure, graph_holder)
    axes = figure.add_subplot()
    axes.plot(weeks, weekly_total, linewidth = 4, color = "blue", label='Weekly Total')
    axes.plot(weeks, design_total, color = "green", label='Design Total')
    axes.plot(weeks, mech_total, color = "red", label='Mechanical Total')
    axes.plot(weeks, elect_total, color = "purple", label='Electrical Total')
    axes.plot(weeks, fab_total, color = "orange", label= 'Fabrication Total')
    axes.set_title('Weekly Delays Over Year')
    axes.set_ylabel('Delays')
    axes.set_xticklabels(weeks, fontsize=5)
    axes.tick_params(axis='x', rotation=45)
    axes.legend()

    figure_canvas.get_tk_widget().pack(side=tk.RIGHT, expand=1)

def openspreadsheet():

    global spreadsheet_path
    spreadsheet_path = fd.askopenfilename(title= "Choose File")


    spreadsheet_file_name = (list(str(spreadsheet_path))[::-1])
    for i in range(len(spreadsheet_file_name)):
        if spreadsheet_file_name[i] == "/":
            spreadsheet_file_name = ("".join(spreadsheet_file_name[0:i][::-1]))
            break
    
    spreadsheet_button.configure(text=spreadsheet_file_name)

    if str(spreadsheet_file_name) == "None" or str(spreadsheet_file_name) == "" or spreadsheet_file_name == None or str(spreadsheet_file_name) == "[]":
        spreadsheet_file_name.configure(text= "Upload Spreadsheet")



if __name__ == "__main__":

    mainmenu()