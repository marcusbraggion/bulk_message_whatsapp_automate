# pty the necessary libraries
import keyboard as k    # for simulating keyboard keys
import pyautogui         # for automating mouse and keyboard actions
import time              # for adding delay in program execution
import pandas as pd      # for reading and manipulating Excel files
import webbrowser as web # for opening URLs in web browser
from urllib.parse import quote # for URL encoding special characters
import os

# Defining a function named send_whatsapp which takes three arguments:
# data_file_excel: the path to an Excel file containing contact information
# message_file_text: the path to a text file containing the message to be sent
# x_cord and y_cord: the coordinates of the mouse click to be performed (default values are set to the location of the send button in WhatsApp Web)

def send_whatsapp(data_file_excel, msg:str, x_cord=830, y_cord=954):
    """
    Sends WhatsApp messages to multiple contacts using data from an Excel file and a message template.
    
    Args:
        data_file_excel (str): The path to the Excel file containing contact information.
        message_file_text (str): The path to the text file containing the message template.
        x_cord (int, optional): The x-coordinate of the mouse click to send the message. Defaults to 830.
        y_cord (int, optional): The y-coordinate of the mouse click to send the message. Defaults to 954.
    
    Returns:
        None
    
    This function reads the contents of an Excel file and a message template file. It then extracts the names and 
    contact numbers from the Excel file and formats the message template with the name of each recipient. It opens 
    the WhatsApp Web URL for each contact and message, simulates a mouse click to send the message, and dismisses 
    the send confirmation pop-up. The function then closes the WhatsApp Web tab after a delay. The function prints a 
    message indicating the number of messages sent after each message is sent. Finally, the function prints a message 
    indicating that the function has completed execution.
    """

    # Reading the Excel file and assigning the contents to a DataFrame named df
    # Also specifying the "Contact" column as a string datatype to prevent Excel from automatically converting phone numbers to scientific notation
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