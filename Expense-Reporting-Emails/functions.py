import traceback
import win32com.client as win32

### Resuable Functions ###
def initialize_outlook():
    try:
        outlook = win32.Dispatch('Outlook.Application')
        return outlook
    except Exception as e:
        print(f'Error occured while initializing Outlook: {e}')
        traceback.print_exc()

def gather_corrections_data(clean_filterd_df):
    unique_employees = clean_filterd_df.groupby(['Employee', 'Email', 'Expense Report'])
    return unique_employees