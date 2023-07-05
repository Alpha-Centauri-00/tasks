import datetime
import openpyxl
import tkinter as tk
from tkcalendar import *
from tkinter import ttk, messagebox






class Tasks:
    def __init__(self, root):
        self.root = root
        self.root.title("Tasks!")
        
        #self.root.geometry("1500x800")
        self.TKINTER_WIDGETS = {}
        self.TKINTER_DATA = {}
        self.APP_WIDTH = 1300
        self.APP_HEIGHT = 900
        self.DEFAULT_TIME = "11:00"

        self.sheet_path = r"D:\PYthon\Notifi_Tasks\Excel_app GUI\people.xlsx"

        self._width = 50
        self.pad_x = 20
        self.pad_y = 10
        self.stick = "ew"

        self.valuesList = []


        self.style = ttk.Style(self.root)
        
        
        self.root.tk.call("source","Excel_app GUI\\forest-dark.tcl")
        self.style.theme_use("forest-dark")
        self.style.configure('TLabelframe.Label', foreground ='#00936a')
        self.style.configure('TLabelframe.Label', font=('courier', 12, 'bold'))
        self.style.configure("Treeview.Heading", foreground="#00936a",font=('courier', 12, 'bold')) # foreground="white"

        ## Createing a frame
        self.frame = ttk.Frame(self.root)
        self.frame.pack(expand=0)

        self.frame_compo = ttk.LabelFrame(self.frame, text="Create a Task")
        self.frame_compo.grid(row=0,column=0,padx=25,pady=25)


        ###############
        # frame_title = ttk.LabelFrame(self.frame_compo,text="Title")
        # frame_title.grid(row=0,column=0,padx=self.pad_x,pady=self.pad_y)

        
        # self.Title_text = tk.Entry(frame_title,borderwidth=0,width=self._width)
        # self.Title_text.grid(row=0,column=0,sticky=self.stick,padx=self.pad_x,pady=self.pad_y)
            
        # if type == "text":
        #     entry_ = tk.Text(frame_,borderwidth=0,width=self._width,height=10)

        self.title = self.create_entry_element(self.frame_compo,"label","Title",0,0)     ## Title Entry



        self.message = self.create_entry_element(self.frame_compo,"text","Message",1,0)    ## Message TextBox

        self.title_  = tk.StringVar()
        self.title['textvariable'] = self.title_
        self.title_.trace_add('write', self._state)
        

        # self.message_ = tk.StringVar()
        # self.message['textvariable'] = self.message_
        # self.message_.trace_add('write', self._state)

        #self.link = self.create_entry_element(self.frame_compo,"label","Link",2,0)      ## Link Entry
        self.cmd = self.create_entry_element(self.frame_compo,"label","Command",3,0)   ## CMD Entry

        ## DATE / TIME
        self.frame_datetime_ = ttk.LabelFrame(self.frame_compo,text="Date & Time")
        self.frame_datetime_.grid(row=5,column=0,sticky=self.stick,padx=self.pad_x,pady=self.pad_y)

        ########################################################
        # Entry Date
        current_date_time = datetime.datetime.now().strftime('%d/%m/%Y')

        self.entry_date = tk.Entry(master=self.frame_datetime_,borderwidth=0,disabledbackground="#b5aeb3",justify="center")
        self.entry_date.insert(0, current_date_time)
        self.entry_date.configure(state=tk.DISABLED,width=10)
        self.entry_date.grid(row=0, column=1,sticky="ewns",padx=self.pad_x,pady=self.pad_y)

        
        # Entry Time
        self.entry_time = tk.Entry(master=self.frame_datetime_,borderwidth=0,disabledbackground="#b5aeb3",justify="center")
        self.entry_time.insert(0, self.DEFAULT_TIME)
        self.entry_time.configure(state=tk.DISABLED,width=10)
        self.entry_time.grid(row=0, column=2,sticky="ewns",padx=self.pad_x,pady=self.pad_y)

        # Button Select Date Time
        self.btn_select_date_time = tk.Button(master=self.frame_datetime_, text="📅",  command=self.select_date_time,borderwidth=0,justify="center")
        self.btn_select_date_time.grid(row=0, column=3, sticky="ewns",padx=self.pad_x,pady=5)
        ########################################################


        self.seperater = ttk.Separator(self.frame_compo)
        self.seperater.grid(row=6,column=0,padx=(20,10),pady=10,sticky="ew")


        self.btn_save_data = ttk.Button(master=self.frame_compo,text="Save",command=self.save_data,state=tk.DISABLED)
        self.btn_save_data.grid(row=8,column=0,sticky=self.stick,padx=self.pad_x,pady=self.pad_y)

        #☽
        # self.theme_mode = ttk.Checkbutton(self.frame,text="\u2600",style="Switch",command=self.change_theme)
        # self.theme_mode.grid(row=7,column=0,padx=5,pady=10,sticky="ew")

        self.frame_table = ttk.Frame(self.frame)
        self.frame_table.grid(row=0,column=1,padx=25,pady=25)

        self.scroll_bar = ttk.Scrollbar(self.frame_table)
        self.scroll_bar.pack(side="right",fill="y")

        self.cols = ("Title","Message","CMD","Date","Time")
        self.tree_view = ttk.Treeview(self.frame_table,show="headings",columns=self.cols,height=20,yscrollcommand=self.scroll_bar.set)
        self.tree_view.column("Title",width=100)
        self.tree_view.column("Message",width=170)
        self.tree_view.column("Date",width=70)
        self.tree_view.column("Time",width=50)
        self.tree_view.pack()
        self.scroll_bar.config(command=self.tree_view.yview)

        self.load_table()



    def _state(self,*_):
        if self.title_.get():
            self.btn_save_data['state'] = 'normal'
        else:
            self.btn_save_data['state'] = 'disabled'

        
    def save_data(self):
        title_ = self.title.get()
        message_ = self.message.get("1.0",tk.END)
        #link_ = self.link.get()
        cmd_ = self.cmd.get()
        date_ = self.entry_date.get()
        time_ = self.entry_time.get()
        ##  ##  ##  ##  ##  ##  ##  ##
        path = self.sheet_path
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        add_row_values = [title_,message_,cmd_,date_,time_]
        sheet.append(add_row_values)
        workbook.save(path)
        
        #self.load_table()
        self.tree_view.insert('',tk.END,values=add_row_values)
    
        self.title.delete(0,tk.END)
        self.message.delete("1.0", tk.END)
        #self.link.delete(0,tk.END)
        self.cmd.delete(0,tk.END)

        # self.TKINTER_WIDGETS["link_entry"].delete(0,customtkinter.END)
        # self.TKINTER_WIDGETS["cmd_entry"].delete(0,customtkinter.END)
        # self.TKINTER_WIDGETS["title_entry"].focus_set()

    def load_table(self):
        
        path = self.sheet_path
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        list_values = list(sheet.values)
        for col_name in list_values[0]:
            self.tree_view.heading(col_name,text=col_name)
        for value_tuple in list_values[1:]:
            self.tree_view.insert("",tk.END,values=value_tuple)
        self.tree_view.bind('<1>', self.select_single_row)
        self.tree_view.bind("<Delete>", self.delete_a_row)

        
        #self.tree_view.bind('<Double-Button-1>', self.double)

            
           

    def show_custom_messagebox(self,title):
        user_choice = None

        def on_yes_click():
            nonlocal user_choice
            user_choice = "Yes"
            dialog.destroy()

        def on_no_click():
            nonlocal user_choice
            user_choice = "No"
            dialog.destroy()

        # Create a custom dialog box
        dialog = tk.Toplevel()
        dialog.title("Confirmation")
        dialog.geometry("300x180")

        #root.tk.call("source","Excel_app GUI\\forest-dark.tcl")
        style = ttk.Style(root)
        style.theme_use("forest-dark")

        frame_ = ttk.LabelFrame(dialog,text="Delete")
        frame_.grid(row=0,column=0,padx=20,pady=20) #

        label = ttk.Label(frame_, text=f"Delete {title}?", font=("Arial", 14,"bold"),justify="center")
        label.grid(row=1, column=0, padx=10, pady=20)


        
        button_y = ttk.Button(frame_, text="Yes", command=on_yes_click)
        button_y.grid(row=2, column=0,padx=10,pady=10)

        button_n = ttk.Button(frame_, text="No", command=on_no_click)
        button_n.grid(row=2, column=1,padx=10,pady=10)

        # Make the dialog box modal (focus stays on the dialog)
        dialog.transient(master=root)
        dialog.grab_set()
        root.wait_window(dialog)

        # Process user_choice after the dialog is closed
        if user_choice == "Yes":
            return True
        return False
    

    def delete_a_row(self,event):

        '''Delete selected row'''
        # print('delete', len(self.tree_view.selection()))
        # if len(self.tree_view.selection()) != 0:
        #     row = self.tree_view.selection()[0]
        # try:
        #     print('deleterow', row)
        #     if messagebox.askokcancel():
        #         print("Delete now")
        #         self.tree_view.delete(row)
        #         deleted_row = row[2:]
        #         print(deleted_row)

        # except:
        #     print('no row selected')
        # self.load_table()


        '''Delete selected row'''
        # if len(self.tree_view.selection()) != 0:
        #     row = self.tree_view.selection()[0]
        #     item_values = self.tree_view.item(row)
        #     row_identifier = item_values['values'][0]  # Assuming the row identifier is in the first column
            

        #     if messagebox.askokcancel("Confirmation", f"Are you sure you want to Delete {(item_values)['values'][0]} ?"):
        #         self.tree_view.delete(row)

        #         # Delete row from Excel file
        #         workbook = openpyxl.load_workbook(self.sheet_path)
        #         worksheet = workbook.active  # Replace 'Sheet1' with the actual sheet name

        #         for excel_row in worksheet.iter_rows():
        #             if excel_row[0].value == row_identifier:  # Assuming the identifier is in the first column
        #                 worksheet.delete_rows(excel_row[0].row)
        #                 break  # Exit the loop after deleting the row

        #         workbook.save(self.sheet_path)

        # else:
        #     print('No row selected')
        
        
        #self.valuesList.append(self.tree_view.item(len(self.tree_view.get_children()),option='values'))
        if len(self.tree_view.selection()) != 0:
            row = self.tree_view.selection()[0]
            item_values = self.tree_view.item(row)
            row_identifier = item_values['values'][0]  # Assuming the row identifier is in the first column
            

            #if self.show_custom_messagebox(item_values["values"][0]):
            if messagebox.askyesno("Confirmation",f"You are going to delete {item_values['values'][0]}"):
                self.tree_view.delete(row)

                # Delete row from Excel file
                workbook = openpyxl.load_workbook(self.sheet_path)
                worksheet = workbook.active  # Replace 'Sheet1' with the actual sheet name

                for excel_row in worksheet.iter_rows():
                    if excel_row[0].value == row_identifier:  # Assuming the identifier is in the first column
                        worksheet.delete_rows(excel_row[0].row)
                        break  # Exit the loop after deleting the row

                workbook.save(self.sheet_path)

        else:
            print('No row selected')

        #self.load_table()





    def cell(self,event):
        '''Identify cell from mouse position'''
        row, col = self.tree_view.identify_row(event.y), self.tree_view.identify_column(event.x)
        pos = self.tree_view.bbox(row, col)       # Calculate positon of entry
        return row, col, pos
    
    def select_single_row(self,event=None):
        '''Single click to select row and column'''
        #global row, col, pos
        row, col, pos = self.cell(event)
        
        print('Select', row, col,pos)
        


    def create_entry_element(self,parent_frame,type,text_value,row_num,col_num):
        frame_ = ttk.LabelFrame(parent_frame,text=text_value)
        frame_.grid(row=row_num,column=col_num,padx=self.pad_x,pady=self.pad_y)

        if type == "label":
            entry_ = tk.Entry(frame_,borderwidth=0,width=self._width)
            
        if type == "text":
            entry_ = tk.Text(frame_,borderwidth=0,width=self._width,height=10)
        entry_.grid(row=row_num,column=col_num,sticky=self.stick,padx=self.pad_x,pady=self.pad_y)
        return entry_


    

    def select_date_time(self):
    
        
        # Disable Button Select EID
        self.btn_select_date_time.configure(state=tk.DISABLED)
        
        # Create Top Level Window
        self.top_level_date_time = tk.Toplevel(takefocus=True)

        # Properties
        self.top_level_date_time.title("Select Date / Time")
        # TKINTER_WIDGETS['top_level_date_time'].iconbitmap(os.path.join(IMAGES_DIRECTORY, CONFIG.get('tkinter', 'icon')))
        
        # Size
        top_level_date_time_width = 400
        top_level_date_time_height = 400

        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        app_center_coordinate_x = (screen_width / 3) - (top_level_date_time_width / 2.5)
        app_center_coordinate_y = (screen_height / 3) - (top_level_date_time_height / 2.5)

        # Position App to the Centre of the Screen
        self.top_level_date_time.geometry(f"{top_level_date_time_width}x{top_level_date_time_height}+{int(app_center_coordinate_x)}+{int(app_center_coordinate_y)}")

        # Prevent User from Resizing the Window
        self.top_level_date_time.resizable(width=False, height=False)

        # Close 'X' Button
        self.top_level_date_time.protocol("WM_DELETE_WINDOW", self.post_top_level_select_date_time)

        self.top_level_date_time.grid_rowconfigure(1, weight=1)
        self.top_level_date_time.grid_columnconfigure(0, weight=1)

        # - Top Level Select Date / Time Design - #

        # Frame Calendar
        frame_calendar = ttk.Frame(master=self.top_level_date_time)
        frame_calendar.grid(row=0, column=0, padx=15, pady=15, columnspan=3)

        # Label Date
        self.label_date =  ttk.Label(master=frame_calendar, text="- Select Date -")
        self.label_date.grid(row=0, column=0, padx=10, pady=5, columnspan=3, sticky='n')

        # Calendar
        self.CAL = Calendar(frame_calendar, selectmode='day', date_pattern='dd/mm/y', mindate=datetime.datetime.today())
        self.CAL.grid(row=0, column=0, padx=20, pady=30, columnspan=3, sticky='s')

        # Frame Time
        frame_time = ttk.Frame(master=self.top_level_date_time)
        frame_time.grid(row=1, column=0, padx=15, pady=5, columnspan=3)

        # Time
        # Label Time
        self.label_time =  ttk.Label(master=frame_time, text="Time")
        self.label_time.grid(row=0, column=0, padx=10, pady=10)
        

        # Hour Time
        self.spinbox_hours = ttk.Spinbox(frame_time, width=10, justify=tk.CENTER,from_=00,to=23, format="%02.0f")
        self.spinbox_hours.grid(row=0, column=1, padx=5)
        

        # Set Default Value
        self.string_var_hours = tk.StringVar()
        self.string_var_hours.set(self.DEFAULT_TIME.split(':')[0])
        self.spinbox_hours.config(textvariable=self.string_var_hours)

        
        # SpinBox Minutes
        self.spinbox_minutes = ttk.Spinbox(frame_time, width=10, justify=tk.CENTER,from_=00,to=59,format="%02.0f")
        self.spinbox_minutes.grid(row=0, column=2, padx=5)
        
        # TextVariable for Minutes
        self.string_var_minutes = tk.StringVar()
        self.string_var_minutes.set(self.DEFAULT_TIME.split(':')[1])
        self.spinbox_minutes.config(textvariable=self.string_var_minutes)

        # Button Select Date Time OK
        self.button_select_date_time_ok = ttk.Button(master=self.top_level_date_time, text="OK", command=self.update_date_time)
        self.button_select_date_time_ok.grid(row=2, column=0, padx=60, pady=15, sticky='sw')

        # Button Select Date Time Cancel
        self.button_select_date_time_cancel = ttk.Button(master=self.top_level_date_time, text="Cancel", command=self.post_top_level_select_date_time)
        self.button_select_date_time_cancel.grid(row=2, column=0, padx=60, pady=15, sticky='se')


    # Post Top Level Select Date Time
    def post_top_level_select_date_time(self):
        
        
        # Destroy Top Level Date Time
        self.top_level_date_time.destroy()

        # Activate Button Select Date Time
        self.btn_select_date_time.configure(state=tk.NORMAL)


    # Update Date Time
    def update_date_time(self):
        
        # Get New Date Time Details
        new_date = self.CAL.get_date()
        # new_time = f'{TKINTER_DATA["string_var_hours"].get()}:{TKINTER_DATA["string_var_minutes"].get()}'
        new_time = f'{self.spinbox_hours.get()}:{self.spinbox_minutes.get()}'

        # Destroy Top Level Widget
        self.post_top_level_select_date_time()

        # Update Date Entry
        self.entry_date.configure(state=tk.NORMAL)
        self.entry_date.delete(0, tk.END)
        self.entry_date.insert(0, new_date)
        self.entry_date.configure(state=tk.DISABLED)

        # Update Time Entry
        self.entry_time.configure(state=tk.NORMAL)
        self.entry_time.delete(0, tk.END)
        self.entry_time.insert(0, new_time)
        self.entry_time.configure(state=tk.DISABLED)

    
    



root = tk.Tk()

task = Tasks(root)




root.mainloop()