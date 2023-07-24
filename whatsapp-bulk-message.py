import tkinter as tk
from tkinter import ttk
import requests
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import pandas
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

new_data = []

def wpbot():

    excel_file = f"your_excel.xlsx"
    excel_data = pandas.read_excel(excel_file, sheet_name='Sheet1')
    count = 0
    message = entry_message.get('1.0', 'end-1c')


    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get('https://web.whatsapp.com')
    time.sleep(45)
    # input("Press ENTER after login into Whatsapp Web and your chats are visiable.")
    for column in excel_data['Phone'].tolist():
        try:
            #message ekstra eklenecek
            url = 'https://web.whatsapp.com/send?phone=' + str(excel_data['Phone'][count]) + '&text=' + message
            sent = True
            driver.get(url)
            try:
                time.sleep(10)

                mesaj = driver.find_element(By.XPATH ,"//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]")
                mesaj.click()
                time.sleep(10)
                butn = driver.find_element(By.XPATH ,"//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[2]/button")
                butn.click()
                sleep(10)
                actions = ActionChains(driver)
                actions.send_keys(Keys.ENTER)
                actions.perform()
            except Exception as e:
                print("Sorry message could not sent to " + str(excel_data['Phone'][count]))
                new_data.append([excel_data['Name'][count], excel_data['Email'][count], excel_data['Phone'][count], excel_data['Price'][count], excel_data['Date Created'][count]])

            else:
                sleep(3)
                print('Message sent to: ' + str(excel_data['Phone'][count]))
            count = count + 1
        except Exception as e:
            print('Failed to send message to ' + str(excel_data['Phone'][count]) + str(e))
    driver.quit()
    print("The script executed successfully.")
    new_df = pd.DataFrame(new_data, columns=["Name", "Email", "Phone", "Price", "Date Created"])
    excel_file = f"your_excel_fail.xlsx"
    new_df.to_excel(excel_file, index=False)






root = tk.Tk()
root.title("*")
root.configure(bg="#66347F")  
root.iconbitmap("*.ico")

root.geometry("800x800")

ttk.Separator(root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)

wp_label = tk.Label(root,bg="green",  text="Whatsapp  Massege Bot")
wp_label.pack()

label_message = ttk.Label(root, text="Message:")
label_message.pack()
entry_message = tk.Text(root, height=10)
entry_message.pack()

ttk.Separator(root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)

wp_button_meta = tk.Button(root,bg="#A4DE02", text="Start Sending * * ", command=wpbot)
wp_button_meta.pack()


ttk.Separator(root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)


root.mainloop()
