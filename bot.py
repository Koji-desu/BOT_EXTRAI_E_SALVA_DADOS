# Import for the Desktop Bot
from botcity.core import DesktopBot

import pandas as pd
def not_found(label):
    print(f"Element not found: {label}")

dados = pd.read_excel(r"C:\Users\Administrator\PycharmProjects\PythonBOTS\ContosoFaturas\Contoso+Coffee+Shop+Invoices.xlsx")

bot = DesktopBot()
path_app = "C:\Program Files (x86)\Contoso, Inc\Contoso Invoicing\LegacyInvoicingApp.exe"

bot.execute(path_app)
bot.wait(2000)
bot.maximize_window()
bot.wait(1000)

if not bot.find("Invoices", matching=0.97, waiting_time=10000):
    not_found("Invoices")
bot.click()
def cadastraFaturas(data, conta, contato, valor, status):



    if not bot.find( "novo-registro", matching=0.97, waiting_time=10000):
        not_found("novo-registro")
    bot.click()

    if not bot.find( "date", matching=0.97, waiting_time=10000):
        not_found("date")
    bot.click_relative(59, 8)

    bot.type_keys(['home'])
    bot.type_keys(['shift' , 'end'])
    bot.paste(data)

    bot.tab()
    bot.paste(conta)

    bot.tab()
    bot.paste(contato)

    bot.tab()
    bot.paste(valor)
    bot.enter()

    coluna = status

    if not bot.find( "status-relative", matching=0.97, waiting_time=10000):
        not_found("status-relative")
    bot.click_relative(61, 7)

    if coluna == "Uninvoiced":
        if not bot.find( "status-relative-univoice", matching=0.97, waiting_time=10000):
            not_found("status-relative-univoice")
        bot.click_relative(75, 35)
    elif coluna == "Invoiced":
        if not bot.find( "click-relative-invoiced", matching=0.97, waiting_time=10000):
            not_found("click-relative-invoiced")
        bot.click_relative(73, 50)
        
    elif coluna == "Paid":
        if not bot.find( "click-relative-paid", matching=0.97, waiting_time=10000):
            not_found("click-relative-paid")
        bot.click_relative(66, 72)
        
        
    else:
        print("OPÇÃO INVALIDA")



    if not bot.find( "click-save", matching=0.97, waiting_time=10000):
        not_found("click-save")
    bot.click()
    


    

    



    
for coluna in dados.itertuples():
    cadastraFaturas(str(coluna[1]), str(coluna[2]), str(coluna[3]), str(coluna[4]), str(coluna[5]))

bot.alt_f4()


