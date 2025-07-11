from tkinter import *
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk, ImageDraw
import openpyxl # For reading Excel files
from email.message import EmailMessage
import smtplib
import os
import io
import re # For email validation
from tkcalendar import Calendar, DateEntry
from datetime import datetime, timedelta
import time
import threading
import json
class ToolTip(object):
    def __init__(self, widget, text='widget info', wraplength=300):
        self.widget = widget
        self.text = text
        self.wraplength = wraplength
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert") or (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 30
        y += self.widget.winfo_rooty() + 30
        self.tip_window = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.geometry(f"+{x}+{y}")
        label = Label(tw, text=self.text, justify='left',
                      background="#ffffe0", relief='solid', borderwidth=1,
                      font=('tahoma', 9), wraplength=self.wraplength)
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None

class emailsender:
    check = False
    def __init__(self, root):
        self.root = root
        # Increased window height to accommodate new elements and better spacing
        self.root.geometry("750x680+80+0") # Adjusted height to 680
        self.root.title("Email Sender")
        self.root.resizable(False, False)
        self.root.config(bg='dodger blue2')
        self.root.wm_iconbitmap('assets\\email.ico')

        # ------------------ VARIABLES ------------------ #
        self.name_var = StringVar() # Still kept, but not used for Excel emails
        self.email_var = StringVar()
        self.subject_var = StringVar()
        self.attachments = []  # To store file paths
        self.email_list = []   # To store email IDs loaded from Excel
        # Email scheduling variables
        self.scheduled_time = None
        self.scheduled_email = None
        self.scheduled_emails_file = "scheduled_emails.json"
        self._start_schedule_monitor()
        self.remember_var = IntVar(value=0)  # 1 = remember, 0 = don't
        self.remembered_email_file = "remember.txt"
        # Variable for radio button selection: 1 for single, 2 for multiple
        self.email_mode_var = IntVar(value=1) # Default to single email mode

        # Status variables for multiple emails
        self.total_emails_var = StringVar(value="Total: 0")
        self.sent_emails_var = StringVar(value="Sent: 0")
        self.failed_emails_var = StringVar(value="Failed: 0")

        # ------------------ Title Section ------------------ #
        img = Image.open("assets\\Email.png")
        self.photoimg = ImageTk.PhotoImage(img)
        title_frame = Frame(self.root, bg='white')
        title_frame.grid(row=0, column=0, pady=5) # Adjusted pady from 10 to 5
        title_label = Label(title_frame, text=' Email Sender', image=self.photoimg, compound=LEFT,
                             font=('goudy old style', 28, 'bold'), bg='white', fg='dodger blue2')
        title_label.grid(row=0, column=0)

        img1 = Image.open("assets\\setting.png")
        self.photoimg1 = ImageTk.PhotoImage(img1)
        setting_button = Button(title_frame, image=self.photoimg1, bg='white', cursor='hand2',
                                activebackground='white', borderwidth=0, command=self.setting)
        setting_button.grid(row=0, column=1, padx=15)
        # Add tooltip to setting button
        ToolTip(setting_button, "Email Credentials Settings")


        # ------------------ To Email Section ------------------ #
        to_label = LabelFrame(root, text='To (Email Address(es))',
                              font=('times new roman', 16, 'bold'),
                              bd=5, fg='white', bg='dodger blue2')
        to_label.grid(row=1, column=0, padx=100, pady=5) # Adjusted pady from 10 to 5

        # Radio buttons for email mode selection
        single_email_radio = Radiobutton(to_label, text="Single Email", variable=self.email_mode_var, value=1,
                                         font=('times new roman', 12, 'bold'), bg='dodger blue2', fg='white',
                                         selectcolor='dodger blue4', command=self._on_email_mode_change)
        single_email_radio.grid(row=0, column=0, padx=5, pady=2, sticky=W) # Adjusted pady from 5 to 2

        multiple_emails_radio = Radiobutton(to_label, text="Multiple Emails (Excel)", variable=self.email_mode_var, value=2,
                                            font=('times new roman', 12, 'bold'), bg='dodger blue2', fg='white',
                                            selectcolor='dodger blue4', command=self._on_email_mode_change)
        multiple_emails_radio.grid(row=0, column=1, padx=5, pady=2, sticky=W) # Adjusted pady from 5 to 2


        # Entry for Email (now editable for manual input or display of loaded emails)
        self.to_entry = Entry(to_label, font=('times new roman', 16, 'bold'),
                              width=25, textvariable=self.email_var, state='normal') # Initial state set by _on_email_mode_change
        self.to_entry.grid(row=1, column=0, pady=2) # Adjusted pady from 5 to 2

        self.remember_check = Checkbutton(
            to_label, text="Remember", variable=self.remember_var,
            onvalue=1, offvalue=0, bg='dodger blue2', fg='white',
            font=('times new roman', 10), selectcolor='dodger blue4'
        )
        self.remember_check.grid(row=2, column=0, sticky=W, padx=5)
       
        # Browse Button to load emails from Excel
        # Ensure 'assets/browse.png' exists. If not, use a text-only button or a placeholder image.
        try:
            browse_img = Image.open("assets\\browse.png")
            browse_img = browse_img.resize((30, 30), Image.Resampling.LANCZOS)
            self.browse_photoimg = ImageTk.PhotoImage(browse_img)
            self.browse_button = Button(to_label, text='  Browse Excel', image=self.browse_photoimg, compound=LEFT,
                                   font=('times new roman', 12, 'bold'), cursor='hand2',
                                   command=self.browse_excel_emails, state='disabled') # Initial state set by _on_email_mode_change
        except FileNotFoundError:
            self.browse_button = Button(to_label, text='Browse Excel',
                                   font=('times new roman', 12, 'bold'), cursor='hand2',
                                   command=self.browse_excel_emails, state='disabled') # Fallback to text button
            messagebox.showwarning("Asset Warning", "assets\\browse.png not found. Using text button for 'Browse Excel'.", parent=self.root)

        self.browse_button.grid(row=1, column=1, padx=15, sticky=W, pady=2) # Adjusted pady from 5 to 2
        ToolTip(self.browse_button, "Browse Excel file for multiple email IDs")


        # ------------------ Subject Section ------------------ #
        subject_label = LabelFrame(root, text='Subject',
                                   font=('times new roman', 16, 'bold'),
                                   bd=5, fg='white', bg='dodger blue2')
        subject_label.grid(row=2, column=0, pady=2) # Adjusted pady from 5 to 2

        self.subject_entry = Entry(subject_label, font=('times new roman', 16, 'bold'),
                                   width=25, textvariable=self.subject_var)
        self.subject_entry.grid(row=0, column=0)


        # ------------------ Compose Email Section ------------------ #
        compose_label = LabelFrame(root, text='Compose Email',
                                   font=('times new roman', 16, 'bold'),
                                   bd=5, fg='white', bg='dodger blue2')
        compose_label.grid(row=3, column=0, pady=5, padx=20) # Adjusted pady from 10 to 5

        img2 = Image.open("assets\\mic.png")
        img2 = img2.resize((48, 48), Image.Resampling.LANCZOS)
        self.photoimg2 = ImageTk.PhotoImage(img2)

        speak_button = Button(compose_label, text='  Speak', image=self.photoimg2, compound=LEFT,
                              font=('arial', 12, 'bold'), cursor='hand2', bd=0, bg='dodger blue2',
                              activebackground='dodger blue2', command=self.speak)
        speak_button.grid(row=0, column=0)

        img3 = Image.open("assets\\attechment.png")
        img3 = img3.resize((48, 48), Image.Resampling.LANCZOS)
        self.photoimg3 = ImageTk.PhotoImage(img3)
        attech_button = Button(compose_label, text='  Attachments', image=self.photoimg3, compound=LEFT,
                               font=('arial', 12, 'bold'), cursor='hand2', bd=0, bg='dodger blue2',
                               activebackground='dodger blue2', command=self.attechment)
        attech_button.grid(row=0, column=1)

        # Label for image preview (initially empty)
        self.image_frame = Frame(compose_label)
        self.image_frame.grid(row=1, column=2, rowspan=2, padx=10, sticky='n')

        # List to keep references to thumbnails
        self.image_thumbnails = []

        # textarea
        textarea_frame = Frame(compose_label)
        textarea_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.textarea = Text(textarea_frame, font=('times new roman', 14), height=7, width=65, pady=5, wrap=WORD)
        self.textarea.grid(row=0, column=0, sticky="nsew")

        # Create Scrollbar widget
        scrollbar = Scrollbar(textarea_frame, orient=VERTICAL, command=self.textarea.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Connect Scrollbar to Text
        self.textarea.config(yscrollcommand=scrollbar.set)

        # ------------------ Action Buttons ------------------ #
        # Adjusted y-coordinates for buttons due to increased height and new status section
        img4 = Image.open("assets\\email_send.png")
        self.photoimg4 = ImageTk.PhotoImage(img4)
        send_button = Button(root, image=self.photoimg4, bg='dodger blue2', cursor='hand2',
                             activebackground='dodger blue2', borderwidth=0, command=self.send_mail)
        send_button.place(x=290, y=540) # Adjusted Y from 570 to 550

        ToolTip(send_button, "Send Email")

        img8 = Image.open("assets\\scheduled.png")
        self.photoimg8 = ImageTk.PhotoImage(img8)
        schedule_button = Button(self.root,image=self.photoimg8,bg='dodger blue2', cursor='hand2',activebackground='dodger blue2', borderwidth=0,command=self.open_schedule_window)
        schedule_button.place(x=390, y=540) # Adjusted Y from 570 to 550
        ToolTip(schedule_button, "Schedule E-mail Sending")

        img5 = Image.open("assets\\Clear.png")
        self.photoimg5 = ImageTk.PhotoImage(img5)

        clear_button = Button(root, image=self.photoimg5, bg='dodger blue2', cursor='hand2',
                              activebackground='dodger blue2', borderwidth=0, command=self.clear)
        clear_button.place(x=490, y=540) # Adjusted Y from 570 to 550
        ToolTip(clear_button, "Clear All Fields")

        img6 = Image.open("assets\\exit.png")
        self.photoimg6 = ImageTk.PhotoImage(img6)
        exit_button = Button(root, image=self.photoimg6, bg='dodger blue2', cursor='hand2',
                             activebackground='dodger blue2', borderwidth=0, command=self.iexit)
        exit_button.place(x=590, y=540) # Adjusted Y from 570 to 550
        ToolTip(exit_button, "Exit Application")

        # ------------------ Status Section (Moved to Bottom) ------------------ #
        status_frame = LabelFrame(root, text='Email Sending Status',
                                  font=('times new roman', 14, 'bold'), # Slightly larger font
                                  bd=5, fg='white', bg='dodger blue2')
        # Placed at a new row at the very bottom
        status_frame.grid(row=4, column=0, pady=(50,0), padx=20, sticky="ew") # Adjusted pady from (55,0) to (45,0)

        total_label = Label(status_frame, textvariable=self.total_emails_var,
                            font=('times new roman', 14), bg='dodger blue2', fg='white')
        total_label.grid(row=0, column=0, padx=10, pady=5, sticky=W)

        sent_label = Label(status_frame, textvariable=self.sent_emails_var,
                           font=('times new roman', 14), bg='dodger blue2', fg='white')
        sent_label.grid(row=0, column=1, padx=10, pady=5, sticky=W)

        failed_label = Label(status_frame, textvariable=self.failed_emails_var,
                             font=('times new roman', 14), bg='dodger blue2', fg='white')
        failed_label.grid(row=0, column=2, padx=10, pady=5, sticky=W)

        # Initialize the state of entry and button based on default radio button selection
        self._on_email_mode_change()
        self._load_remembered_email()
        
    def open_schedule_window(self):
        """Open a new window to schedule the email."""
        schedule_window = Toplevel(self.root)
        schedule_window.title("Schedule Email")
        schedule_window.geometry("400x320+200+100")
        schedule_window.config(bg='dodger blue2')
        schedule_window.resizable(False, False)

        try:
            schedule_window.wm_iconbitmap('assets\\email.ico')
        except:
            pass

        title_label = Label(schedule_window, text='Schedule Email',
                            font=('goudy old style', 20, 'bold'),
                            fg='white', bg='dodger blue2')
        title_label.pack(pady=15)

        schedule_label = Label(schedule_window, text="Select Date:",
                               font=('times new roman', 14, 'bold'),
                               bg='dodger blue2', fg='white')
        schedule_label.pack(pady=5)

        self.schedule_date = DateEntry(schedule_window, width=15,
                                       font=('times new roman', 12),
                                       date_pattern='yyyy-mm-dd',
                                       background='dodger blue4', foreground='white',
                                       borderwidth=2, relief="groove")
        self.schedule_date.pack(pady=5)

        time_label = Label(schedule_window, text="Select Time (HH:MM) in ISO format(24 hour format):",
                           font=('times new roman', 12, 'bold'),
                           bg='dodger blue2', fg='white')
        time_label.pack(pady=5)

        time_frame = Frame(schedule_window, bg='dodger blue2')
        time_frame.pack(pady=5)

        self.hour_spinbox = Spinbox(time_frame, from_=0, to=23, width=5,
                                    font=('times new roman', 12),
                                    format="%02.0f", bd=2, relief="groove",
                                    bg='white', fg='black', buttonbackground='dodger blue4')
        self.hour_spinbox.grid(row=0, column=0, padx=5)
        ToolTip(self.hour_spinbox, "Enter hour (00-24)")
        # Set minute spinbox initial value to 15
        self.hour_spinbox.delete(0, 'end')
        self.hour_spinbox.insert(0, f"{datetime.now().hour:02}")

        colon_label = Label(time_frame, text=":", font=('times new roman', 12, 'bold'), bg='dodger blue2', fg='white')
        colon_label.grid(row=0, column=1)

        self.minute_spinbox = Spinbox(time_frame, from_=0, to=59, width=5,
                                      font=('times new roman', 12),
                                      format="%02.0f", bd=2, relief="groove",
                                      bg='white', fg='black', buttonbackground='dodger blue4')
        self.minute_spinbox.grid(row=0, column=2, padx=5)

        # Set minute spinbox initial value to 15
        self.minute_spinbox.delete(0, 'end')
        self.minute_spinbox.insert(0, f"{datetime.now().minute:02}")

        # Add tooltip to the minute spinbox
        ToolTip(self.minute_spinbox, "Enter minutes (00-59)")
        schedule_button = Button(schedule_window, text="Schedule Email",
                                 font=('times new roman', 14, 'bold'),
                                 bg='gold2', fg='black', cursor='hand2',
                                 activebackground='dodger blue4', activeforeground='white',
                                 borderwidth=0, command=self.schedule_email)
        schedule_button.pack(pady=20)
        ToolTip(schedule_button, "Confirm and schedule the email")

    def schedule_email(self):
        """Schedule the email for a future time."""
        schedule_date = self.schedule_date.get()
        schedule_hour = self.hour_spinbox.get()
        schedule_minute = self.minute_spinbox.get()

        try:
            # Combine date and time into a single datetime object
            scheduled_datetime = datetime.strptime(f"{schedule_date} {schedule_hour}:{schedule_minute}", "%Y-%m-%d %H:%M")
            if scheduled_datetime < datetime.now():
                messagebox.showerror("Error", "Scheduled time cannot be in the past.")
                return

            self.scheduled_time = scheduled_datetime
            # messagebox.showinfo("Info", f"Email scheduled for {self.scheduled_time}")

            self.schedule_send(scheduled_datetime)

        except Exception as e:
            messagebox.showerror("Error", f"Invalid date/time format: {e}")

    def _on_email_mode_change(self):
        """Adjusts the state of the 'To' entry and 'Browse Excel' button based on radio button selection."""
        selected_mode = self.email_mode_var.get()
        if selected_mode == 1:  # Single Email
            self.to_entry.config(state='normal')
            self.browse_button.config(state='disabled')
            self.email_list = [] # Clear the list if switching to single mode
            self.total_emails_var.set("Total: 0")
            self.sent_emails_var.set("Sent: 0")
            self.failed_emails_var.set("Failed: 0")
            self.email_var.set("") # Clear displayed email if any
        else:  # Multiple Emails (Excel)
            self.to_entry.config(state='readonly') # Allow display of "X emails loaded"
            self.browse_button.config(state='normal')
            self.email_var.set("Browse for Excel file...") # Hint for user
            self.email_list = [] # Clear the list if switching to multiple mode
            self.total_emails_var.set("Total: 0")
            self.sent_emails_var.set("Sent: 0")
            self.failed_emails_var.set("Failed: 0")
        if selected_mode == 1:
            self.remember_check.config(state='normal')
        else:
            self.remember_var.set(0)
            self.remember_check.config(state='disabled')

    def iexit(self):
        exit_app = messagebox.askyesno('Notification', 'Do you want to exit the application', parent=self.root)
        if exit_app > 0:
            self.root.destroy()
        else:
            return

    def clear(self):
        self.to_entry.delete(0, END)
        self.subject_entry.delete(0, END)
        self.textarea.delete(1.0, END)
        self.attachments = []
        self.image_thumbnails = []
        self.email_list = [] # Clear the loaded email list
        self.total_emails_var.set("Total: 0")
        self.sent_emails_var.set("Sent: 0")
        self.failed_emails_var.set("Failed: 0")
        self.email_mode_var.set(1) # Reset to single email mode
        self._on_email_mode_change() # Apply state changes for single mode
        messagebox.showinfo("Information", "All fields and status cleared.", parent=self.root)

    def speak(self):
        # Ensure pygame mixer is initialized before use
        try:
            from pygame import mixer
            mixer.init()
            mixer.music.load('assets\\beep.mp3')
            mixer.music.play()
        except Exception as e:
            messagebox.showerror('Error', f'Could not initialize mixer or load beep.mp3: {e}', parent=self.root)
            return

        try:
            import speech_recognition
            sr = speech_recognition.Recognizer()
            with speech_recognition.Microphone() as m:
                sr.adjust_for_ambient_noise(m, duration=0.2)
                audio = sr.listen(m)
                text = sr.recognize_google(audio)
                self.textarea.insert(END, text + '.')
        except Exception as e:
            messagebox.showerror('Speech Recognition Error', f'Sorry your speech is not recognised due to {str(e)}',
                                 parent=self.root)

    def setting(self):
        # State variable to track password visibility
        self.password_visible = False # Initialize to hidden

        def clear1():
            from_entry.delete(0, END)
            pass_entry.delete(0, END)

        def save():
            if from_entry.get() == '' or pass_entry.get() == '':
                messagebox.showerror("Error", 'All fields are required', parent=root1)
            else:
                with open('credentials.txt', 'w') as f1:
                    f1.write(from_entry.get() + ',' + pass_entry.get())
                    messagebox.showinfo('Information', 'Credentials Save successfully', parent=root1)

        # Function to toggle password visibility using a button
        def toggle_password_visibility():
            self.password_visible = not self.password_visible # Toggle the state
            if self.password_visible: # If password is now visible
                pass_entry.config(show='')
                # Try to load open eye image, fallback to text
                try:
                    self.eye_toggle_button.config(image=self.eye_open_photo, text="", compound=LEFT)
                except AttributeError: # If images not loaded, use text
                    self.eye_toggle_button.config(text="Hide")
            else: # If password is now hidden
                pass_entry.config(show='*')
                # Try to load closed eye image, fallback to text
                try:
                    self.eye_toggle_button.config(image=self.eye_closed_photo, text="", compound=LEFT)
                except AttributeError: # If images not loaded, use text
                    self.eye_toggle_button.config(text="Show")


        root1 = Toplevel()
        root1.title('Setting')
        root1.geometry('620x350+350+70') # Increased height for toggle button
        root1.config(bg='dodger blue2')
        root1.resizable(False, False)
        img = Image.open("assets\\Email.png")
        root1.wm_iconbitmap('assets\\email.ico')
        self.photoimg = ImageTk.PhotoImage(img)
        title_label = Label(root1, text='Credential Settings', image=self.photoimg, compound=LEFT,
                            font=('goudy old style', 38, 'bold'), fg='white', bg='gray20')
        title_label.grid(row=0, column=0, padx=75)
        from_label = LabelFrame(root1, text='From (Email Address)',
                                font=('times new roman', 16, 'bold'),
                                bd=5, fg='white', bg='dodger blue2')
        from_label.grid(row=1, column=0, pady=15)
        from_entry = Entry(from_label, font=('times new roman', 16, 'bold'),
                           width=35)
        from_entry.grid(row=0, column=0)
        pass_label = LabelFrame(root1, text='Password',
                                font=('times new roman', 16, 'bold'),
                                bd=5, fg='white', bg='dodger blue2')
        pass_label.grid(row=2, column=0, pady=15)
        pass_entry = Entry(pass_label, font=('times new roman', 16, 'bold'),
                           width=35, show='*')
        pass_entry.grid(row=0, column=0)

        # Load eye images for toggle button, with fallback to text
        try:
            eye_open_img = Image.open("assets\\eye_open.png") # Assuming this asset exists
            eye_open_img = eye_open_img.resize((20, 20), Image.Resampling.LANCZOS)
            self.eye_open_photo = ImageTk.PhotoImage(eye_open_img)

            eye_closed_img = Image.open("assets\\eye_closed.png") # Assuming this asset exists
            eye_closed_img = eye_closed_img.resize((20, 20), Image.Resampling.LANCZOS)
            self.eye_closed_photo = ImageTk.PhotoImage(eye_closed_img)

            # Initial button configuration with closed eye image
            self.eye_toggle_button = Button(pass_label, image=self.eye_closed_photo,
                                            command=toggle_password_visibility,
                                            bd=0, bg='dodger blue2', activebackground='dodger blue2',
                                            cursor='hand2')
            self.eye_toggle_button.grid(row=0, column=1, padx=5, sticky=W)
            ToolTip(self.eye_toggle_button, "Show/Hide Password")

        except FileNotFoundError:
            # Fallback to text button if images are not found
            self.eye_toggle_button = Button(pass_label, text="Show",
                                            command=toggle_password_visibility,
                                            font=('times new roman', 10), bg='dodger blue2', fg='white',
                                            activebackground='dodger blue2', activeforeground='white',
                                            cursor='hand2')
            self.eye_toggle_button.grid(row=0, column=1, padx=5, sticky=W)
            ToolTip(self.eye_toggle_button, "Show/Hide Password")
            messagebox.showwarning("Asset Warning", "assets\\eye_open.png or assets\\eye_closed.png not found. Using text button for password toggle.", parent=root1)


        save_button = Button(root1, text='Save', bg='gold2', fg='black', cursor='hand2',
                             font=('times new roman', 18, 'bold'), activebackground='gray10',
                             activeforeground='white', borderwidth=0, command=save)
        save_button.place(x=200, y=290) # Adjusted Y for new height
        # Add tooltip to save button in settings
        ToolTip(save_button, "Save your email credentials")

        clear_button = Button(root1, text='Clear', bg='gold2', fg='black', cursor='hand2',
                            font=('times new roman', 18, 'bold'),activebackground='gray10',activeforeground='white', borderwidth=0,command=clear1)
        clear_button.place(x=320,y=290) # Adjusted Y for new height
        # Add tooltip to clear button in settings
        ToolTip(clear_button, "Clear credentials fields")

        # Load existing credentials if available
        try:
            with open('credentials.txt') as f:
                for i in f:
                    cr = i.strip().split(',')
            from_entry.insert(0, cr[0])
            pass_entry.insert(0, cr[1])
        except FileNotFoundError:
            pass # File doesn't exist yet, no credentials to load

        root1.mainloop()

    def attechment(self):
        files = filedialog.askopenfilenames(initialdir=os.getcwd(), title='Select Files', parent=self.root)
        for file_path in files:
            if file_path not in self.attachments:
                self.attachments.append(file_path)
                filename = os.path.basename(file_path)
                self.textarea.insert(END, f"\n")

                ext = filename.split('.')[-1].lower()
                if ext in ['png', 'jpg', 'jpeg', 'ico']:
                    img = Image.open(file_path)
                    img.thumbnail((50, 50))
                    thumb = ImageTk.PhotoImage(img)

                elif ext in ['mp3', 'wav']:
                    img = Image.new('RGB', (100, 50), color='lightgray')
                    d = ImageDraw.Draw(img)
                    d.text((0, 15), f'Audio: {filename}', fill='black')
                    thumb = ImageTk.PhotoImage(img)
                elif ext in ['mp4', 'avi']:
                    img = Image.new('RGB', (100, 50), color='lightgray')
                    d = ImageDraw.Draw(img)
                    d.text((5, 15), f'Video: {filename}', fill='black')
                    thumb = ImageTk.PhotoImage(img)

                else:
                    # Create a placeholder icon for non-image files
                    img = Image.new('RGB', (100, 50), color='lightgray')
                    d = ImageDraw.Draw(img)
                    d.text((0, 15), f'DOC : {filename}', fill='black')
                    thumb = ImageTk.PhotoImage(img)

                # Store reference to prevent garbage collection
                if not hasattr(self, 'thumb_refs'):
                    self.thumb_refs = []
                self.thumb_refs.append(thumb)

                self.textarea.image_create(END, image=thumb)
                self.textarea.insert(END, "\n")

    def browse_excel_emails(self):
        """Allows user to browse for an Excel file and load email IDs."""
        file_path = filedialog.askopenfilename(
            initialdir=os.getcwd(),
            title='Select Excel File',
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")),
            parent=self.root
        )

        if file_path:
            self.email_list = [] # Clear previous list
            self.sent_count = 0
            self.failed_count = 0
            self.total_emails_var.set("Total: 0")
            self.sent_emails_var.set("Sent: 0")
            self.failed_emails_var.set("Failed: 0")
            self.to_entry.delete(0, END) # Clear previous entry display

            try:
                workbook = openpyxl.load_workbook(file_path)
                sheet = workbook.active
                
                # Assuming emails are in the first column (column A)
                # You can modify this to search for a specific column header like 'Email'
                email_column_index = 0 # Default to first column (A)

                # Try to find a column named 'Email' (case-insensitive)
                header_row = [cell.value for cell in sheet[1]]
                for i, header in enumerate(header_row):
                    if header and str(header).strip().lower() == 'email':
                        email_column_index = i
                        break

                for row_index, row in enumerate(sheet.iter_rows(min_row=2)): # Start from second row to skip header
                    # Get cell value from the determined email column
                    email_cell_value = row[email_column_index].value
                    if email_cell_value:
                        email = str(email_cell_value).strip()
                        # Basic email validation using regex
                        if re.match(r"[^@]+@[^@]+\.[^@]+", email):
                            self.email_list.append(email)
                        else:
                            print(f"Skipping invalid email format: {email}") # For debugging
                            
                if self.email_list:
                    self.total_emails_var.set(f"Total: {len(self.email_list)}")
                    self.email_var.set(f"{len(self.email_list)} emails loaded from Excel")
                    messagebox.showinfo("Information", f"{len(self.email_list)} email(s) loaded successfully from {os.path.basename(file_path)}", parent=self.root)
                else:
                    self.email_var.set("No valid emails found")
                    messagebox.showwarning("Warning", "No valid email addresses found in the selected Excel file.", parent=self.root)

            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Excel file: {e}", parent=self.root)
                self.email_var.set("Error loading Excel")

    def send_mail(self):
        subject = self.subject_var.get()
        message_body = self.textarea.get(1.0, END).strip()

        if not subject or not message_body:
            messagebox.showerror('Error', 'Subject and message body are required', parent=self.root)
            return

        selected_mode = self.email_mode_var.get()
        recipients_to_send = []

        if selected_mode == 1: # Single Email Mode
            single_email = self.to_entry.get().strip()
            if not single_email or not re.match(r"[^@]+@[^@]+\.[^@]+", single_email):
                messagebox.showerror('Error', 'Please enter a valid single email address.', parent=self.root)
                return
            recipients_to_send = [single_email]
        else: # Multiple Emails (Excel) Mode
            if not self.email_list:
                messagebox.showerror('Error', 'No email addresses loaded from Excel. Please browse and select an Excel file.', parent=self.root)
                return
            recipients_to_send = self.email_list

        if not recipients_to_send:
            messagebox.showerror('Error', 'No email addresses to send to.', parent=self.root)
            return
        
        self._update_remembered_email()
        self.sent_count = 0
        self.failed_count = 0
        self.total_emails_var.set(f"Total: {len(recipients_to_send)}")
        self.sent_emails_var.set("Sent: 0")
        self.failed_emails_var.set("Failed: 0")

        messagebox.showwarning("Email Delivery Info",
                               "Our email may initially appear in receiver's spam or junk folder.\n\n"
                               "To ensure they receive future emails in their inbox, please tell them to mark the email as 'Not Spam' or move it to their inbox.\n\n"
                               "Thank you for your cooperation.", parent=self.root)

        try:
            with open('credentials.txt') as f1:
                cr = f1.readline().strip().split(',')
            sender_email = cr[0]
            sender_password = cr[1]
        except FileNotFoundError:
            messagebox.showerror("Error", "Credentials file (credentials.txt) not found. Please set your email and password in settings.", parent=self.root)
            return
        except Exception as e:
            messagebox.showerror("Error", f"Error reading credentials: {e}", parent=self.root)
            return

        # Initialize SMTP connection once for all emails
        s = None
        try:
            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.starttls()
            s.login(sender_email, sender_password)

            for i, recipient_address in enumerate(recipients_to_send):
                try:
                    message = EmailMessage()
                    message['Reply-To'] = sender_email
                    message['X-Mailer'] = 'Attendance System'
                    message['subject'] = subject
                    message['to'] = recipient_address
                    message['from'] = sender_email
                    message.set_content(message_body + "\n\nRegards,\nDev\nEmail sender")

                    # Add attachments
                    for path in self.attachments:
                        filename = os.path.basename(path)
                        ext = filename.split('.')[-1].lower()
                        with open(path, 'rb') as f:
                            file_data = f.read()
                        if ext in ['png', 'jpg', 'jpeg', 'ico']:
                            img = Image.open(io.BytesIO(file_data))
                            subtype = img.format.lower()
                            message.add_attachment(file_data, maintype='image', subtype=subtype, filename=filename)
                        else:
                            message.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=filename)

                    s.send_message(message)
                    self.sent_count += 1
                    # print(f"Sent to: {recipient_address}") # Debugging
                except Exception as e:
                    self.failed_count += 1
                    # print(f"Failed to send to {recipient_address}: {e}") # Debugging
                finally:
                    self.sent_emails_var.set(f"Sent: {self.sent_count}")
                    self.failed_emails_var.set(f"Failed: {self.failed_count}")
                    self.root.update_idletasks() # Update GUI immediately

            messagebox.showinfo("Information",
                                f"Email sending complete!\nTotal: {len(recipients_to_send)}\nSent: {self.sent_count}\nFailed: {self.failed_count}",
                                parent=self.root)

        except smtplib.SMTPAuthenticationError:
            messagebox.showerror("Error", "Failed to login to SMTP server. Check your email and password in settings.", parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during sending: {e}", parent=self.root)
        finally:
            if s:
                s.quit() # Ensure the SMTP connection is closed

    def schedule_send(self, scheduled_time):
        """Save scheduled email details to JSON and clear form."""
        subject = self.subject_var.get()
        message_body = self.textarea.get(1.0, END).strip()

        if not subject or not message_body:
            messagebox.showerror('Error', 'Subject and message body are required', parent=self.root)
            return

        # Determine recipients
        mode = self.email_mode_var.get()
        if mode == 1:
            recipients = [self.to_entry.get().strip()]
        else:
            recipients = self.email_list.copy()

        # Basic validation
        if not recipients or any(not re.match(r"[^@]+@[^@]+\.[^@]+", r) for r in recipients):
            messagebox.showerror("Error", "Valid recipient email(s) required.", parent=self.root)
            return

        # Prepare data
        scheduled_email = {
            "time": scheduled_time.strftime("%Y-%m-%d %H:%M:%S"),
            "recipients": recipients,
            "subject": subject,
            "body": message_body,
            "attachments": self.attachments.copy()
        }

        # Save to JSON
        if os.path.exists(self.scheduled_emails_file):
            with open(self.scheduled_emails_file, "r") as f:
                all_schedules = json.load(f)
        else:
            all_schedules = []

        all_schedules.append(scheduled_email)
        with open(self.scheduled_emails_file, "w") as f:
            json.dump(all_schedules, f, indent=2)

        messagebox.showinfo("Scheduled", f"Email scheduled for {scheduled_time}", parent=self.root)
        self._update_remembered_email()
        self.clear()  # Reset form

    def _start_schedule_monitor(self):
        threading.Thread(target=self._monitor_scheduled_emails, daemon=True).start()

    def _monitor_scheduled_emails(self):
        while True:
            if os.path.exists(self.scheduled_emails_file):
                with open(self.scheduled_emails_file, "r") as f:
                    try:
                        schedules = json.load(f)
                    except:
                        schedules = []

                now = datetime.now()
                remaining = []

                for email in schedules:
                    scheduled_time = datetime.strptime(email["time"], "%Y-%m-%d %H:%M:%S")
                    if now >= scheduled_time:
                        self._send_scheduled_email(email)
                    else:
                        remaining.append(email)

                # Save back remaining schedules
                with open(self.scheduled_emails_file, "w") as f:
                    json.dump(remaining, f, indent=2)

            time.sleep(30)  # check every 30 seconds

    def _send_scheduled_email(self, email):
        try:
            with open('credentials.txt') as f1:
                cr = f1.readline().strip().split(',')
            sender_email = cr[0]
            sender_password = cr[1]

            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.starttls()
            s.login(sender_email, sender_password)

            for recipient in email["recipients"]:
                msg = EmailMessage()
                msg['From'] = sender_email
                msg['To'] = recipient
                msg['Subject'] = email["subject"]
                msg.set_content(email["body"] + "\n\nRegards,\nDev\nEmail sender")

                for path in email["attachments"]:
                    try:
                        with open(path, 'rb') as f:
                            data = f.read()
                        filename = os.path.basename(path)
                        ext = filename.split('.')[-1].lower()
                        if ext in ['png', 'jpg', 'jpeg']:
                            img = Image.open(io.BytesIO(data))
                            subtype = img.format.lower()
                            msg.add_attachment(data, maintype='image', subtype=subtype, filename=filename)
                        else:
                            msg.add_attachment(data, maintype='application', subtype='octet-stream', filename=filename)
                    except Exception as e:
                        print(f"Attachment error: {e}")

                s.send_message(msg)

            s.quit()
            print(f"[Scheduled] Sent to: {', '.join(email['recipients'])}")
            msgbox_text = f"Scheduled email sent successfully!\n\nTo: {', '.join(email['recipients'])}\nSubject: {email['subject']}"
            self.root.after(0, lambda: messagebox.showinfo("Email Sent", msgbox_text, parent=self.root))

        except Exception as e:
            print(f"[Scheduled] Failed: {e}")

    def _load_remembered_email(self):
        if os.path.exists(self.remembered_email_file):
            with open(self.remembered_email_file) as f:
                email = f.read().strip()
            if email:
                self.email_var.set(email)
                self.remember_var.set(1)
                self.to_entry.delete(0, END)
                self.to_entry.insert(0, email)

    def _update_remembered_email(self):
        if self.email_mode_var.get() == 1 and self.remember_var.get() == 1:
            with open(self.remembered_email_file, "w") as f:
                f.write(self.to_entry.get().strip())
        elif os.path.exists(self.remembered_email_file):
            os.remove(self.remembered_email_file)

if __name__ == '__main__':
    root = Tk()
    obj = emailsender(root)
    root.mainloop()
