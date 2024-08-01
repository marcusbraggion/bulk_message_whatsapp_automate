# pty the necessary libraries
import keyboard as k    # for simulating keyboard keys
import pyautogui         # for automating mouse and keyboard actions
import time              # for adding delay in program execution
import pandas as pd      # for reading and manipulating Excel files
import webbrowser as web # for opening URLs in web browser
from urllib.parse import quote # for URL encoding special characters
import os

def send_whatsapp(data_file_excel, msg:str, x_cord=830, y_cord=954):
    df = pd.read_excel(data_file_excel, sheet_name="campanha")
    print(df)
    
    # Extracting the values in the "Name" and "Contact" columns of the DataFrame and assigning them to variables named name and contact, respectively
    name = df['NOME'].values
    contact = df['TELEFONE'].values

    print(f"{name}, {contact}")
    
    # Zipping together the values in name and contact into tuples and assigning them to a variable named zipped
    zipped = zip(name, contact)
    
    # Initializing a counter variable to keep track of the number of messages sent
    counter = 0
    
    # Looping over each tuple in zipped
    # Opening the WhatsApp Web URL for the corresponding contact and message text
    web.open(f"https://web.whatsapp.com/send?phone={contact}&text={quote(msg)}")
    
    # Adding a delay to allow the WhatsApp Web page to load
    time.sleep(15)
    
    # Simulating a mouse click at the specified coordinates to send the message
    pyautogui.click(x_cord, y_cord)
    
    # Adding a delay to allow time for the message to be sent
    time.sleep(2)
    
    # Simulating the pressing of the "Enter" key to dismiss the send confirmation pop-up
    k.press_and_release('enter')
    
    # Adding a delay to allow time for the pop-up to be dismissed
    time.sleep(2)
    
    # Simulating the pressing of the "Ctrl + W" keys to close the WhatsApp Web tab
    k.press_and_release('ctrl+w')
    
    # Adding a delay to allow time for the tab to be closed
    time.sleep(1)
    
    # Incrementing the counter variable and printing a message indicating that the message has been sent
    print("- Message sent..!!")

# Printing a message indicating that the function has completed execution
    print("Done!")

# Defining the paths to the Excel file and message text file to be used as inputs to the send_whatsapp function
if __name__ == "__main__":
    path = "C:/Users/vinic/OneDrive/Ambiente de Trabalho/Repos/bulk_message_whatsapp_automate/"
    excel_file:str = os.path.join(path, "contacts.xlsx")
    print(excel_file)
    send_whatsapp(excel_file,msg="Olá, Marcus")