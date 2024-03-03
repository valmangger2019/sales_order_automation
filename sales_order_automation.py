#Importing all libraries
import pandas as pd
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import win32com.client
import sys
import openpyxl

#Starting the program
def main():
    #1 - Initializin the SAP system
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    #2 - Defining a dataframe
    df_tabela = pd.read_excel('C:/Users/wande/OneDrive/Área de Trabalho/2024/Pessoal/Projectpython/Salesproject.xlsx', sheet_name= 'Pag1')

    current_row = 0

    #3 - Inicializin the iterrows
    for index, row in df_tabela.iterrows():

        #Declaration of variables
        order_type_value =  df_tabela.loc[current_row,'Order type']
        customer_reference_value = df_tabela.loc[current_row, 'Customer reference'] 
        order_issuer_value = df_tabela.loc[current_row,  'Order issuer'] 
        goods_recipient_value = df_tabela.loc[current_row, 'Goods recipient'] 
        material_value = df_tabela.loc[current_row,  'Material'] 
        order_quantity_value = df_tabela.loc[current_row,  'Order quantity'] 

        # Conditional if order_type.value NULL or " " end the application, otherwise continue.
        if order_type_value == 'END' or order_type_value == '  ' or order_type_value == '0':
            print('Finished aplication')
            break

        else:
            print('There are orders to be created')

        try: 
            print('Starting aplication')
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "VA01"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = order_type_value
            time.sleep(1)
            session.findById("wnd[0]/usr/ctxtVBAK-VKORG").setFocus()
            session.findById("wnd[0]/usr/ctxtVBAK-VKORG").caretPosition = 0
            session.findById("wnd[0]").sendVKey (4)
            session.findById("wnd[1]/usr/lbl[1,5]").setFocus()
            session.findById("wnd[1]/usr/lbl[1,5]").caretPosition = 2
            session.findById("wnd[1]").sendVKey (2)
            session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").setFocus()
            session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").caretPosition = 0
            session.findById("wnd[0]").sendVKey (4)
            session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 2
            session.findById("wnd[1]").sendVKey (2)
            session.findById("wnd[0]/usr/ctxtVBAK-SPART").setFocus()
            session.findById("wnd[0]/usr/ctxtVBAK-SPART").caretPosition = 0
            session.findById("wnd[0]").sendVKey (4)
            session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 3
            session.findById("wnd[1]").sendVKey (2)
            session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").setFocus()
            session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").caretPosition = 0
            session.findById("wnd[0]").sendVKey (4)
            session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 3
            session.findById("wnd[1]").sendVKey (2)
            session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").setFocus()
            session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").caretPosition = 0
            session.findById("wnd[0]").sendVKey (4)
            session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 2
            session.findById("wnd[1]").sendVKey (2)
            session.findById("wnd[0]/usr/ctxtVBAK-AUART").setFocus()
            session.findById("wnd[0]/usr/ctxtVBAK-AUART").caretPosition = 2
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = customer_reference_value
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = order_issuer_value
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").text = goods_recipient_value 
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").setFocus()
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 9
            session.findById("wnd[0]").sendVKey (0)

            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]").sendVKey (4)

            try:
               session.findById("wnd[1]/usr/lbl[1,3]").caretPosition = 2
               session.findById("wnd[1]").sendVKey (2)

               session.findById("wnd[0]").sendVKey (0)
            except:
               print('Payment condition not found')
           
            #Material information
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW").children[0].children[0].children[1].children[1].children[8].text = material_value
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW").children[0].children[0].children[1].children[1].children[24].text = order_quantity_value
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
            session.findById("wnd[0]/tbar[0]/btn[15]").press()

            current_row += 1
            time.sleep(1)
        except:
              print('Erro, please verify! Some information is missing or incorrect')   

if __name__ == "__main__":
    main()





















