import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from threading import Thread
from time import sleep, time
from Expense_Reporting_Emails.Sent_Back_Email import main as sent_back_email_main
from Expense_Reporting_Emails.corrections import main as corrections_email_main
from Personal_Expense_Deduction_Email.Personal_Expense_Deduction_Email import main as personal_expense_main

class EmailProcessingApp:
    def __init__(self, root):
        self.root = root
        self.notebook = None
        self.stop_timer = False

        # Setup the UI
        self.setup_ui()

    def setup_ui(self):
        """ Set up the notebook and the two tabs. """
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")

        # Create the two tabs
        self.setup_tab("Sent Back Email", self.run_sent_back_email)
        self.setup_tab("Corrections Email", self.run_corrections_email)
        # self.setup_tab("Personal Expense Email", self.run_personal_expense_email)

    def prompt_sheet_name(self):
        ### it is should be only prompting when it needs a user input, WIP
        # """ Prompt the user for the sheet name using a popup dialog. """
        # self.sheet_name = simpledialog.askstring("Input", "Please enter the sheet name:")
        
        # if not self.sheet_name:
        #     self.root.after(0, lambda: messagebox.showwarning("Input Error", "Sheet name is required!"))
        #     return False
        # return True
        pass

    def setup_tab(self, tab_name, script_function):
        """ Create a tab with shared UI components and set the script to run. """
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text=tab_name)

        # Timer label and done button are specific to each tab
        timer_label = tk.Label(frame, text="Ready to process.", font=("Helvetica", 16))
        timer_label.pack(pady=20)

        done_button = ttk.Button(frame, text="Done", command=lambda: self.reset_ui(timer_label, done_button))
        done_button.pack_forget()  # Hide the done button initially

        # Start button to begin processing
        start_button = ttk.Button(frame, text="Start", command=lambda: self.start_processing(script_function, timer_label, done_button))
        start_button.pack(pady=20)

    def start_processing(self, script_function, timer_label, done_button):
        """ Start the script, show the stopwatch, and run the script. """
        # if not self.prompt_sheet_name():
        #     return  # Stop if no sheet name is provided
        
        self.stop_timer = False
        timer_label.config(text="Processing...")
        done_button.pack_forget()  # Hide the done button

        # Start the stopwatch and the script in separate threads
        Thread(target=self.start_timer, args=(timer_label,)).start()  # Start the timer
        Thread(target=script_function, args=(timer_label, done_button)).start()  # Run the script in a separate thread

    def start_timer(self, timer_label):
        """ Start a stopwatch timer that shows elapsed time. """
        start_time = time()
        while not self.stop_timer:
            elapsed_time = time() - start_time
            minutes, seconds = divmod(int(elapsed_time), 60)
            milliseconds = int((elapsed_time % 1) * 100)
            timer_label.config(text=f"Elapsed time: {minutes}m:{seconds:02}s:{milliseconds:02}ms")
            sleep(0.01)

    def run_sent_back_email(self, timer_label, done_button):
        try:
            sent_back_email_main()  # This is your existing Sent_Back_Email function
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.stop_timer = True  # Stop the stopwatch timer
            done_button.pack(pady=10)  # Show the done button

    def run_corrections_email(self, timer_label, done_button):
        try:
            corrections_email_main()
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.stop_timer = True  # Stop the stopwatch timer
            done_button.pack(pady=10)  # Show the done button

    ### Working on the kinks for this function
    def run_personal_expense_email(self, timer_label, done_button):
        # try:
        #     personal_expense_main()
        # except Exception as e:
        #     messagebox.showerror("Error", str(e))
        # finally:
        #     self.stop_timer = True  # Stop the stopwatch timer
        #     done_button.pack(pady=10)  # Show the done button
        pass

    def reset_ui(self, timer_label, done_button):
        """ Reset the UI components. """
        timer_label.config(text="Ready to process.")
        done_button.pack_forget()
        messagebox.showinfo("Reset", "Ready to run again.")

# Create the main window
root = tk.Tk()
root.title("Work Projects")
root.geometry("800x600")

# Initialize the application
app = EmailProcessingApp(root)

# Start the Tkinter main loop
root.mainloop()
