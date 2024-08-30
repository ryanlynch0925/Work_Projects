import tkinter as tk
from tkinter import ttk, messagebox
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
        self.current_batch = 0
        self.batch_size = 30
        self.unique_employees = []  # This will store the grouped employees data
        self.timer_running = False
        self.start_time = 0
        self.elapsed_time = 0
        self.tabs = {}  # Dictionary to store tab widget references

        # Setup the UI
        self.setup_ui()

    def setup_ui(self):
        """ Set up the notebook and the tabs. """
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")

        # Create tabs for different types of emails
        self.setup_tab("Sent Back Email", self.run_sent_back_email)
        self.setup_tab("Corrections Email", self.run_corrections_email)
        # self.setup_tab("Personal Expense Email", self.run_personal_expense_email)

    def setup_tab(self, tab_name, script_function):
        """ Create a tab with shared UI components and set the script to run. """
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text=tab_name)

        # Timer label and done button are specific to each tab
        timer_label = tk.Label(frame, text="Ready to process.", font=("Helvetica", 16))
        timer_label.pack(pady=20)

        done_button = ttk.Button(frame, text="Done", command=lambda: self.reset_ui(tab_name))
        done_button.pack_forget()  # Hide the done button initially

        # Start button to begin processing
        start_button = ttk.Button(frame, text="Start", command=lambda: self.start_processing(script_function, tab_name))
        start_button.pack(pady=20)

        # Next button for handling batch processing
        next_button = ttk.Button(frame, text="Next Batch", command=lambda: self.process_next_batch(tab_name))
        next_button.pack(pady=20)
        next_button.pack_forget()  # Hide the next button until needed

        # Store references in the tabs dictionary
        self.tabs[tab_name] = {
            'start_button': start_button,
            'done_button': done_button,
            'timer_label': timer_label,
            'next_button': next_button
        }

    def start_processing(self, script_function, tab_name):
        """ Start the script, show the stopwatch, and run the script. """
        self.stop_timer = False
        self.timer_running = True
        self.start_time = time()  # Initialize start time
        self.elapsed_time = 0  # Reset elapsed time
        tab_widgets = self.tabs[tab_name]
        tab_widgets['timer_label'].config(text="Processing...")
        tab_widgets['done_button'].pack_forget()  # Hide the done button

        # Start the stopwatch and the script in separate threads
        Thread(target=self.start_timer, args=(tab_name,)).start()  # Start the timer
        Thread(target=script_function, args=(tab_name,)).start()  # Run the script in a separate thread

    def start_timer(self, tab_name):
        """ Start a stopwatch timer that shows elapsed time. """
        while not self.stop_timer:
            if self.timer_running:
                current_time = time()
                self.elapsed_time = current_time - self.start_time
                minutes, seconds = divmod(int(self.elapsed_time), 60)
                milliseconds = int((self.elapsed_time % 1) * 100)
                # Use Tkinter's thread-safe method to update UI
                self.root.after(0, self.tabs[tab_name]['timer_label'].config, {
                    'text': f"Elapsed time: {minutes}m:{seconds:02}s:{milliseconds:02}ms"
                })
            sleep(0.01)

    def pause_timer(self):
        """ Pause the stopwatch timer. """
        self.timer_running = False

    def resume_timer(self):
        """ Resume the stopwatch timer. """
        self.start_time = time() - self.elapsed_time  # Adjust start time to account for the time already elapsed
        self.timer_running = True

    def run_sent_back_email(self, tab_name):
        """ Run the Sent Back Email script. """
        try:
            self.unique_employees = sent_back_email_main()  # Assume this returns grouped employee data
            if self.unique_employees is None:
                raise ValueError("Error: Sent Back Email script did not return any data.")
            self.current_batch = 0  # Reset batch index
            self.process_next_batch(tab_name)  # Start processing the first batch
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.stop_timer = True  # Stop the stopwatch timer

    def run_corrections_email(self, tab_name):
        """ Run the Corrections Email script. """
        try:
            self.unique_employees = corrections_email_main()  # Assume this returns grouped employee data
            self.current_batch = 0  # Reset batch index
            self.process_next_batch(tab_name)  # Start processing the first batch
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.stop_timer = True  # Stop the stopwatch timer

    def process_next_batch(self, tab_name):
        """ Process the next batch of emails (30 at a time). """
        start_idx = self.current_batch * self.batch_size
        end_idx = start_idx + self.batch_size
        batch = self.unique_employees[start_idx:end_idx]

        if not batch:
            messagebox.showinfo("Batch Complete", "All batches processed.")
            self.tabs[tab_name]['done_button'].pack(pady=10)
            self.stop_timer = True
            return

        # Simulate processing batch
        for employee, data in batch:
            print(f"Processing email for {employee}")  # Replace with actual email processing logic

        self.current_batch += 1

        # Pause the timer after the current batch is processed
        self.pause_timer()

        # Determine whether to display "Next Batch" or "Done" button
        if self.current_batch * self.batch_size >= len(self.unique_employees):
            self.tabs[tab_name]['done_button'].pack(pady=10)
        else:
            self.tabs[tab_name]['next_button'].pack(pady=10)

    def reset_ui(self, tab_name):
        """ Reset the UI components. """
        tab_widgets = self.tabs[tab_name]
        tab_widgets['timer_label'].config(text="Ready to process.")
        tab_widgets['done_button'].pack_forget()
        tab_widgets['next_button'].pack_forget()  # Hide the Next button
        messagebox.showinfo("Reset", "Ready to run again.")

# Create the main window
root = tk.Tk()
root.title("Work Projects")
root.geometry("800x600")

# Initialize the application
app = EmailProcessingApp(root)

# Start the Tkinter main loop
root.mainloop()
