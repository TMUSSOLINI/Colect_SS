from selenium import webdriver
from selenium.webdriver.common.by import By
from tkinter import messagebox
from urllib.parse import quote
from dotenv import load_dotenv
import time
import webbrowser
import sqlite3
import datetime
import pyautogui
import tkinter as tk
import os


load_dotenv()
sql_queries = ["select count(*) from ticket where reportdate > dateadd(day,-1,getdate()) and ticketid not in(select relatedreckey from relatedrecord) and historyflag = 0 and source != 'ENEL';",
"select count(1) from ms_createsr where ms_source = 'GAIA' and ms_createdate > dateadd(day,-1,getdate()) and ms_ticketid is not null;",
"select count(1) from ms_createsr where ms_source = 'SP156' and ms_createdate > dateadd(hour,-1,getdate()) and ms_ticketid is not null;",
"select count(1) from ms_createsr where ms_source = 'SELIMP' and ms_createdate > dateadd(day,-1,getdate()) and ms_ticketid is not null;",
"select count(1) from ms_createsr where ms_source = 'SGZ' and ms_createdate > dateadd(day,-1,getdate()) and ms_ticketid is not null;",
"select count(1) from sr where source = 'URANO' and reportdate > dateadd(day,-20,getdate()) and ticketid is not null;"]


def login(driver, user, password):
    try:
        link_prod = os.getenv('SITE_PROD')
        driver.get(link_prod)
        time.sleep(2)

        username_field = driver.find_element(By.ID, "j_username")
        password_field = driver.find_element(By.ID, "j_password")

        username_field.send_keys(user)
        password_field.send_keys(password)

        time.sleep(2)

        login_button = driver.find_element(By.ID, "loginbutton")
        login_button.click()
    except Exception as e:
        error_message(e)


def access_sql(driver):
    try:
        system_configuration = driver.find_element(By.ID, "m7f8f3e49_ns_menu_UTIL_MODULE_a")
        system_configuration.click()
        time.sleep(2)

        sql_tool = driver.find_element(By.ID, "m7f8f3e49_ns_menu_UTIL_MODULE_sub_changeapp_MS_SQLT_a")
        sql_tool.click()
        time.sleep(5)

        standard_query = driver.find_element(By.ID,"m6a7dfd2f_tdrow_[C:1]_ttxt-lb[R:0]")
        standard_query.click()
        time.sleep(5)
    except Exception as e:
        error_message(e)


def execute_sql(driver, consults):
    try:
        text_area = driver.find_element(By.ID,"mb6ac736-ta")

        for query in sql_queries:
            text_area.send_keys(query)
            text_area.send_keys("\n")

        execute_button = driver.find_element(By.ID,"maa8ad01-pb")
        execute_button.click()
        time.sleep(80)
    except Exception as e:
        error_message(e)


def consult_result(driver):
    try:
        iframe = driver.find_element(By.ID, "me25d1b3-rtv")
        driver.switch_to.frame(iframe)
        time.sleep(80)
        query_result = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[2]/td/font")
        ss_gaia = int(query_result.text)
        return ss_gaia
    except Exception as e:
        error_message(e)
    
def send_message(value):
    try:
        day = datetime.date.today()
        day_form = day.strftime("%d/%m/%Y")
        if value != int(0):
            send = "No"
            execute_update("update SS set NOSS = ?, data_atualiza = ?, Enviado = ?", 0, day_form, send)
        else:
            days, date, send = execute_select_ss("select * from SS")
            new_day = days + 1
            send = 'Yes'
            execute_update("update SS set NOSS = ?, data_atualiza = ?, enviado = ?", new_day, day_form, send)
            execute_automation()
    except Exception as e:
        error_message(e)

def execute_automation():
    try:
        phone = os.getenv('TELEFONE')
        mensagem = message_text_wpp()
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={phone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        time.sleep(15)
        pyautogui.press('enter')
        time.sleep(5)
    except Exception as e:
        error_message(e)




def connect_db():
    try:
        conn = sqlite3.connect('Dias_sem_SS.db')
        cursor = conn.cursor()
        return cursor, conn
    except Exception as e:
        error_message(e)


def execute_update(comand, new_day, day_form, send):
    try:
        cursor, conn = connect_db()
        cursor.execute(comand, (new_day,day_form, send))
        conn.commit()
        conn.close()
    except Exception as e:
        error_message(e)


def execute_select_ss(comand):
    try:
        cursor, conn = connect_db()
        cursor.execute(comand)
        n_ss, date,send = cursor.fetchone()

        return n_ss, date,send
    except Exception as e:
        error_message(e)


def message_text_wpp():
    try:
        day, date, send = execute_select_ss("select * from SS")
        if day > 0:
            message_text_day = f"Bom dia Renan tudo bem. Estamos a {int(day)} dias sem SS's da GAIA"
            return message_text_day
        else:
            message_text = f"Bom dia Renan tudo bem. Não recebemos desde ontem SS's da GAIA"
            return message_text
    except Exception as e:
        error_message(e)


def display_message():
    root = tk.Tk()
    days, date, send = execute_select_ss("select * from SS")
    if send == 'Yes':
        messagebox.showinfo("Informação", "O programa foi executado e a mensagem foi enviada")
    else:
        messagebox.showinfo("Atenção", "O programa foi executado e a mensagem não foi enviada pois as SS's estavam normais.")
    root.destroy()


def error_message(e):
    root = tk.Tk()
    messagebox.showerror("Erro", f"A execução do aplicativo parou devido a um erro: {e}")
    root.destroy()


if __name__ == '__main__':
    user = os.getenv('LOGIN')
    password = os.getenv('SENHA')
    driver = webdriver.Chrome()
    login(driver,user,password)
    access_sql(driver)
    execute_sql(driver,sql_queries)
    ss_gaia = consult_result(driver)
    send_message(ss_gaia)
    display_message()
    driver.quit()
