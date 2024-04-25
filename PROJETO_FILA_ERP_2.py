import win32com.client
import time
import pandas as pd 
import subprocess
import win32gui
import win32con
import sqlite3
import shutil
from datetime import datetime
import os


def relatorio_va05():

    # DEFININDO O CAMINHO DO SAP
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

    # ABRIR O SAP
    subprocess.Popen(path)
    time.sleep(3)

    # FAZER LOGIN EM PRODUCAO
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.OpenConnection("PRD", True)
    session = connection.Children(0)

    # INSERIR CREDENCIAL
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "400"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "BOARETO"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Pgbml12032004#!"
    session.findById("wnd[0]").sendVKey(0)

    try:
        # Tentar selecionar a opção de logon múltiplo
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except Exception:
        # Se ocorrer um erro, continuar normalmente sem fazer nada
        pass

    # RELATORIO VA05
    time.sleep(1.5)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nva05"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 3
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"
    session.findById("wnd[1]/tbar[0]/btn[2]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[32]").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 10
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "0-10"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 84
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 80
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "84"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 0
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 0
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 2
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "2"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 80
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 76
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "80"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 11
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 6
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "11"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 80
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 69
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "80"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 16
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 13
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "16"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 65
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 64
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "65"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 78
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 74
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "78"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 17
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 16
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "17"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 40
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 34
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "40"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellColumn = "DO_SUM"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "40"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 77
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 67
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "77"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 70
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 63
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "70"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 12
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 11
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "12"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 13
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 2
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "13"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 19
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 12
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "19"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").setCurrentCell (17,"SELTEXT")
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").firstVisibleRow = 5
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "17"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").doubleClickCurrentCell
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "19"
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 30
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = 17
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").firstVisibleRow = 5
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectedRows = "17"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell (9,"BSTKD")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem ("&XXL")
    session.findById("wnd[1]/usr/chkCB_ALWAYS").setFocus()
    session.findById("wnd[1]/usr/chkCB_ALWAYS").selected = True
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"M:\FULFILLMENT\FILA_ERP_P"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "BASE.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    time.sleep(5)

    # INFORMANDO O NOME DA JANELA-EXCEL
    excel_window = win32gui.FindWindow(None, 'BASE - Excel')

    # SE A JANELA ESTIVER ABERTA, FECHAR
    if excel_window: 
        win32gui.PostMessage(excel_window, win32con.WM_CLOSE, 0, 0)

    # ENCERRAR O SAP
    time.sleep(3.5)
    subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])

def obter_ufs_cidades_cep():

    file_path = r"M:\FULFILLMENT\FILA_ERP_P\BASE.XLSX"
    txt_obter_ufs= r"M:\FULFILLMENT\FILA_ERP_P\obter_ufs.txt"

    # Carregue os dados do arquivo Excel para um DataFrame
    df = pd.read_excel(file_path)

    if 'Emissor da ordem' in df.columns:
        # Extrai os pedidos da coluna 'Documento de vendas' e junta-os em uma única string
        df_unique = df.drop_duplicates(subset=['Emissor da ordem'])
        
        # Cria uma string com os pedidos para escrever no arquivo de texto
        pedidos_str = '\n'.join(df_unique['Emissor da ordem'].astype(str))

        # Atualiza o arquivo de texto com os novos pedidos
        with open(txt_obter_ufs, 'w') as arquivo:
            arquivo.write(pedidos_str)

        print("Pedidos salvos e arquivo atualizado.")
    else:
        print("A coluna 'Emissor da ordem' não foi encontrada no DataFrame.")

    # DEFININDO O CAMINHO DO SAP
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

        # ABRIR O SAP
    subprocess.Popen(path)
    time.sleep(3)

        # FAZER LOGIN EM PRODUCAO
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.OpenConnection("PRD", True)
    session = connection.Children(0)

        # INSERIR CREDENCIAL
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "400"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "BOARETO"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Pgbml12032004#!"
    session.findById("wnd[0]").sendVKey(0)

    try:
        # Tentar selecionar a opção de logon múltiplo
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except Exception:
        # Se ocorrer um erro, continuar normalmente sem fazer nada
        pass

    #logijn tabela

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NSE16N"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = "KNA1"
    session.findById("wnd[0]/usr/ctxtGD-TAB").setFocus()
    session.findById("wnd[0]/usr/ctxtGD-TAB").caretPosition = 4
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").setFocus()
    session.findById("wnd[0]/usr/txtGD-MAX_LINES").caretPosition = 0
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").setFocus()
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").caretPosition = 0
    session.findById("wnd[0]/tbar[1]/btn[18]").press()
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").text = "KUNNR"
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").setFocus()
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").caretPosition = 5
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").text = "NAME1"
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").setFocus()
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").caretPosition = 5
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").text = "ORT01"
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").setFocus()
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").caretPosition = 5
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").text = "REGIO"
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").setFocus()
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").caretPosition = 5
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").text = "PSTLZ"
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").setFocus()
    session.findById("wnd[0]/usr/ctxtGD_ADD_COLUMN").caretPosition = 5
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,1]").selected = True
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,2]").selected = True
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,2]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 3
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").selected = True
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,1]").selected = True
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,2]").selected = True
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC").verticalScrollbar.position = 0
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus()
    session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press()
    session.findById("wnd[1]/tbar[0]/btn[21]").press()
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"M:\FULFILLMENT\FILA_ERP_P"
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "obter_ufs.txt"
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 13
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    session.findById("wnd[0]").sendVKey (8)
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").setCurrentCell (3,"NAME1")
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectAll()
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem ("&XXL")
    session.findById("wnd[1]/usr/chkCB_ALWAYS").setFocus()
    session.findById("wnd[1]/usr/chkCB_ALWAYS").selected = True
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"M:\FULFILLMENT\FILA_ERP_P"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    time.sleep(5)

    # INFORMANDO O NOME DA JANELA-EXCEL
    excel_window = win32gui.FindWindow(None, 'EXPORT - Excel')

    # SE A JANELA ESTIVER ABERTA, FECHAR
    if excel_window: 
        win32gui.PostMessage(excel_window, win32con.WM_CLOSE, 0, 0)

    # ENCERRAR O SAP
    time.sleep(3)
    subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])

def analise_das_ovs():

    # Passando o caminho da planilha base
    file_path = r"M:\FULFILLMENT\FILA_ERP_P\BASE.XLSX"
    
    # Lendo a planilha
    base_pedro = pd.read_excel(file_path)
    
    # Adicionando a coluna Status
    base_pedro['Status'] = ''

    # Adicionando a coluna UF após a coluna Status
    base_pedro.insert(base_pedro.columns.get_loc('Status') + 1, 'UF', '')

    # Adicionando a coluna CIDADE após a coluna UF
    base_pedro.insert(base_pedro.columns.get_loc('UF') + 1, 'CIDADE', '')

    base_pedro.insert(base_pedro.columns.get_loc('CIDADE') + 1, 'CEP', '')

    # Planilha com os dados para preencher a coluna "UF"
    file_path_dados_uf = r"M:\FULFILLMENT\FILA_ERP_P\EXPORT.XLSX"

    # Lendo a planilha para preencher a coluna "UF"
    dados_uf = pd.read_excel(file_path_dados_uf)




    # Realizando a operação semelhante ao "procv" para a UF
    grupo_dados_uf = dados_uf.groupby('Cliente')['Região'].first().reset_index()

    # Realizando a operação semelhante ao "procv" para CIDADE
    grupo_dados_uf_cidade = dados_uf.groupby('Cliente')['Local'].first().reset_index()

    # Realizando a operação semelhante ao "procv" para o CEP
    grupo_dados_uf_cep = dados_uf.groupby('Cliente')['Código postal'].first().reset_index()    


    # Mapeando os valores de UF na base_pedro usando o 'Documento de vendas' como chave
    base_pedro['UF'] = base_pedro['Emissor da ordem'].map(grupo_dados_uf.set_index('Cliente')['Região'])

    # Mapeando os valores de UF na base_pedro usando o 'Documento de vendas' como chave
    base_pedro['CIDADE'] = base_pedro['Emissor da ordem'].map(grupo_dados_uf_cidade.set_index('Cliente')['Local'])

    #mapear cep

    base_pedro['CEP'] = base_pedro['Emissor da ordem'].map(grupo_dados_uf_cep.set_index('Cliente')['Código postal'])

    # Informando os clientes que não devem ser liberados em massa
    clientes_com_particuliaridades = [
        'IRMAOS GONCALVES COMERCIO E INDUSTRIA LT', 'DRI COMERCIO DE BIJUTERIAS E ACESSORIOS LTDA', 'KOLKE DO BRASIL ', 'KOLKE DO BRASIL IMPORTACAO E EXPORTACAO LTDA', 'ALUTECH TECNOLOGIA E LOCACOES SA', 
        'SUPERMERCADOS BH COMERCIO DE ALIMENTOS S', 'SUPERMERCADOS BH COMERCIO DE ALIMENTOS SA', 'NEW YORK COM. E IMPORTACAO EXPORTAO LTD', 
        'NEW YORK COM. E IMPORTACAO E EXPORTAO LTDA', 'VANDIR RECH FILHO E CIA LTDA', 'BADU GAMES LTDA', 'AMERICANAS S.A.', 'AMERICANAS S.A', 
        'WIBBI FIT LTDA', 'WMS SUPERMERCADOS DO BRASIL SA', 'LEROY MERLIN CIA BRASILEIRA DE BRICOLAGE M', 'SISTECNICA INFORMATICA E SERVICOS EIREL', 
        'LOJAS DE DEPARTAMENTOS MILIUM LTDA', 'O. MIRANDA DA ROCHA COMERCIO DE MOVEIS E IRELI', 'PET SHOP INTERLAGOS COMERCIAL LTDA', 
        'LOGIN INFORMATICA COM E REPRESENTACAO LT','APOLLO MATERIAIS MEDICO HOSPITALARES LTDA', 'BAMBINO FREE SHOP EIRELI', 'ELGIN INDUSTRIAL DA AMAZONIA LTDA', 
        'AFC COMERCIO VAREJISTA DE COMPONENTES ELETRONICOS LTDA', 'ANALISE INFORMATICA LTDA.', 'BAH FREE SHOP LTDA', 'COMERCIAL DE MOVEIS BRASILIA LTDA', 
        'CONDOR SUPER CENTER LTDA', 'DISTRIBUIDORA DE MEDICAMENTOS SANTA CRUZ LTDA', 'DUFRY DO BRASIL DUTY FREE SHOP LTDA.', 'DUTY FREE SHOP TM IMPORTACAO E EXPORTACAO LTDA', 
        'DUTY FREE SHOP TM IMPORTACAO E EXPORTACA', 'ELETROMAR MOVEIS E ELETRODOMESTICOS LTDA', 'FORTLUX ATACADISTA LTDA', 'UL TESTTECH LABORATORIOS DE AVALIACAO DA CONFORMIDADE LTDA'
    ]
    clientes_com_particuliaridades_bp = [
        'A000805', '230947', 'A099652', '1326653', 'A023987', 'A051601', '560271', 'A073929', '933999', 'A054513', '400624', 'CEN1100'
        'A048751', 'A099652'
    ]
    base_pedro.loc[base_pedro['Nome emissor ordem'].isin(clientes_com_particuliaridades), 'Status'] = 'Tratar'
    base_pedro.loc[base_pedro['Emissor da ordem'].isin(clientes_com_particuliaridades_bp), 'Status'] = 'Tratar'
    
    # Informando os canais que não devem ser liberados em massa
    canais_especificos = [
        'PROV', 'GOV1', 'OPER', 'CORP', 'VARN', 'PRO2', 'OEM1', 'OEM', 'WATT', 'GOV2', 'M2C', 'GOV', 'EXP1'
    ]
    base_pedro.loc[base_pedro['Escritório de vendas'].isin(canais_especificos), 'Status'] = 'Out. responsáveis'
    
    # Analisando bloqueio Ag. Saldo 
    bloqueio_saldo = ['Ag. Saldo', 'Ag. Saldo Taberna']
    base_pedro.loc[(base_pedro['Descrição do bloqueio da remessa'].isin(bloqueio_saldo)) & (base_pedro['Organização de vendas'] == '1000'), 'Status'] = 'Ag. saldo'
    
    # Analisando bloqueio Ag. Analise Fiscal
    base_pedro.loc[base_pedro['Descrição do bloqueio da remessa'] == 'Anál.Fiscal Pendente', 'Status'] = 'Ag. fiscal'
    
    # Somando o valor total por ordem de venda
    soma_valores = base_pedro.groupby('Documento de vendas')['Valor líquido (item)'].sum().reset_index()
    
    # Definindo pedidos abaixo de R$ 400,00
    documentos_abaixo_400 = soma_valores[soma_valores['Valor líquido (item)'] < 330]['Documento de vendas']
    
    # Analisando pedidos abaixo de R$ 400,00
    base_pedro.loc[base_pedro['Documento de vendas'].isin(documentos_abaixo_400) & (base_pedro['Status global'] == 'B'), 'Status'] = 'Abaixo de 400'
    
    # Analisando se o status é "Abaixo de 400" e o tipo de pedido for bonificado
    base_pedro.loc[(base_pedro['Status'] == 'Abaixo de 400') & (base_pedro['Tipo doc.vendas'] == 'YBON'), 'Status'] = 'Remessa Liberada'
    
    # Analisando se o status é "Abaixo de 4000" e o 'Escritório de vendas' está em canais_especificos
    base_pedro.loc[(base_pedro['Status'] == 'Abaixo de 400') & (base_pedro['Escritório de vendas'].isin(canais_especificos)), 'Status'] = 'Out. responsáveis'
    
    # Definindo quais OVS vamos liberar
    base_pedro.loc[base_pedro['Status'] == '', 'Status'] = 'Remessa Liberada'
    
    # Definindo clientes que vamos travar em lotação
    clientes_lotacao = [
        'DROGARIA ARAUJO S A', 'LFB SOCIEDADE LTDA', 'WMS SUPERMERCADOS DO BRASIL SA', 'BANDEIRANTES COMERCIO DE RACOES LTDA', 'ARMARINHOS FERNANDO LTDA', 'CARREFOUR COMERCIO E INDUSTRIA LTDA', 'GRUPO CASAS BAHIA', 
        'CENTRAL PET MARINGA DISTRIBUIDORA DE PRODUTOS PARA', 'COBASI O SHOPPING DO SEU ANIMAL', 'INNOVAT DISTRIBUIDORA E COMERCIO LTDA', 'LOJAS CEM SA', 'MAGAZINE LUIZA SA', 
        'ARTHUR LUNDGREN TECIDOS S A CASAS PERNAMBUCANAS', 'PET CENTER COMERCIO E PARTICIPACOES S.A.', 'PETSUPERMARKET COMERCIO DE PRODUTOS PARA ANIMAIS LTDA', 'REVAL ATACADO DE PAPELARIA LTDA.'
    ]
    clientes_lotacao_bp = [
        'A026384', 'A025089', 'A001024', '306118', 'A010484', '150559', 'A047560', 
        '4645260', 'A050916', 'A017510', '290004', 'A059775', '4529275', 'A050533', 
        'A049975', 'A049978', 'A049971', 'A049972', 'A049967', 'A040239', '2611333', 
        'A054513', 'A510682'
    ]
    base_pedro.loc[base_pedro['Nome emissor ordem'].isin(clientes_lotacao), 'Status'] = 'Lotaçao - clientes Mac'
    base_pedro.loc[base_pedro['Emissor da ordem'].isin(clientes_lotacao_bp), 'Status'] = 'Lotaçao - clientes Mac'
    base_pedro.loc[base_pedro['Nome emissor ordem'].isin(clientes_lotacao) & (base_pedro['Escritório de vendas'].isin(canais_especificos)), 'Status'] = 'Out. responsáveis'
    base_pedro.loc[base_pedro['Emissor da ordem'].isin(clientes_lotacao_bp) & (base_pedro['Escritório de vendas'].isin(canais_especificos)), 'Status'] = 'Out. responsáveis'
    
    # Definindo clientes para barrar o faturamento
    clientes_travados = ['SNG SUPLEMENTOS LTDA', 'AMAZON SERVICOS DE VAREJO DO BRASIL LTDA']
    clientes_travados_bp = ['A075077']
    base_pedro.loc[base_pedro['Nome emissor ordem'].isin(clientes_travados), 'Status'] = 'BARRAR FATURAMENTO'
    base_pedro.loc[base_pedro['Emissor da ordem'].isin(clientes_travados_bp), 'Status'] = 'BARRAR FATURAMENTO'
    
    # Analisando pedidos da GIGA-MANAUS
    base_pedro.loc[base_pedro['Organização de vendas'] == 3000, 'Status'] = 'GIGA MANAUS'
    base_pedro.loc[(base_pedro['Descrição do bloqueio da remessa'].isin(bloqueio_saldo)) & (base_pedro['Organização de vendas'] == '3000'), 'Status'] = 'GIGA MANAUS'
    
    # Definir WELLNESS acima de R$80K para travar em lotação
    soma_valores_well = base_pedro[base_pedro['Escritório de vendas'] == 'WELL'].groupby('Documento de vendas')['Valor líquido (item)'].sum()
    well_acima_80000 = soma_valores_well[soma_valores_well >= 80000].index
    
    # Inserindo o status para lotação WELLNESS
    base_pedro.loc[base_pedro['Documento de vendas'].isin(well_acima_80000), 'Status'] = 'Lotacao Well'
    
    # Analisando bloqueio Ag. Pagamento
    base_pedro.loc[base_pedro['Descrição do bloqueio da remessa'] == 'Aguardando PgtBoleto', 'Status'] = 'Ag. Pagamento'
    
    # Analisando se é um venda ordem e a venda já foi faturada
    base_pedro.loc[(base_pedro['Tipo doc.vendas'] == 'YRCS') & (base_pedro['Status global'] == 'B'), 'Status'] = 'Venda OP. trg Faturada'
    
    # Lista de valores mínimos por UF
    valores_minimos = {
        'AC': 100000, 'AL': 100000, 'AP': 100000, 'AM': 100000, 'BA': 70000, 'CE': 100000, 'ES': 30000, 'GO': 50000,
        'MA': 100000, 'MT': 100000, 'MS': 100000, 'MG': 40000, 'PA': 100000, 'PB': 100000, 'PR': 30000, 'PE': 100000,
        'PI': 100000, 'RJ': 20000, 'RN': 100000, 'RS': 30000, 'RO': 100000, 'RR': 100000, 'SC': 40000, 'SP': 20000,
        'SE': 100000, 'TO': 100000, 'DF': 50000
    }

    # Criar uma máscara para filtrar os pedidos cujo status está em branco ou já foi definido como "Tratar" ou "Remessa Liberada"
    #mascara = (base_pedro['Status'] == '') | (base_pedro['Status'] == 'Tratar') | (base_pedro['Status'] == 'Remessa Liberada')

    # Calcular o valor total dos pedidos para cada UF com base na máscara de status
    #total_por_cidade = base_pedro[mascara].groupby(['UF', 'CIDADE'])['Valor líquido (item)'].sum()

    # Realizar a análise e atribuir os status adequados antes de "Lotação TESTE"
    #for (uf, cidade), valor_total in total_por_cidade.items():
    #    valor_minimo_uf = valores_minimos.get(uf, 0)  # Obtém o valor mínimo para o estado
    #    if valor_total >= valor_minimo_uf:
            # Filtrar os pedidos dessa cidade que atendem aos critérios
    #        pedidos_cidade = base_pedro[mascara & (base_pedro['UF'] == uf) & (base_pedro['CIDADE'] == cidade)]
            
            # Analisar cada pedido da cidade
    #        for index, pedido in pedidos_cidade.iterrows():
                # Verificar se o status anterior era "Tratar"
    #            if pedido['Status'] == 'Tratar':
    #                base_pedro.at[index, 'Status'] = 'Lotação - Tratar cliente'
                # Verificar se o status anterior era "Remessa Liberada"
    #            elif pedido['Status'] == 'Remessa Liberada':
    #                base_pedro.at[index, 'Status'] = 'Lotação - Liberar remessa'
                # Se não era nenhum dos anteriores, atribuir "Lotação TESTE"
    #            else:
    #                base_pedro.at[index, 'Status'] = 'Lotação teste'


    # Salvando a base .XLSX
    base_pedro.to_excel(r"M:\FULFILLMENT\FILA_ERP_P\LIBERAÇÃO_OV.xlsx", index=False, engine='openpyxl') 
    
    # Caminho da base 
    caminho_relatorio = r"M:\FULFILLMENT\FILA_ERP_P\LIBERAÇÃO_OV.xlsx"
    
    # Caminho da cópia da base
    caminho_destino = r"M:\FULFILLMENT\REM_PEDRO"
    
    # Pegando a data e hora atuais
    agora = datetime.now()
    data_hora_atual = agora.strftime("%Y-%m-%d_%H-%M-%S")
    
    # Definindo o caminho completo
    nome_arquivo_copia = f"LIBERAÇÃO_OV_{data_hora_atual}.xlsx"
    caminho_copia = os.path.join(caminho_destino, nome_arquivo_copia)
    
    
    # Salvando cópia
    if shutil.copyfile(caminho_relatorio, caminho_copia):
        print('relatorio atualizado')
    else:
        print('atualizacao falhou')    

        print('atualizacao falhou')    

def reprocessar_saldo_zerado():
    
    # Caminho da base
    excel_path = r"M:\FULFILLMENT\FILA_ERP_P\LIBERAÇÃO_OV.xlsx"
    df = pd.read_excel(excel_path)
    
    # Selecionando itens com a quantidade confirmada 0
    filtered_df = df[(df['Quantidade confirmada (divisão remessa)'] == 0) & (df['Status'] == 'Remessa Liberada')]
    
    # Verificar se há alguma linha após o filtro
    if not filtered_df.empty:
        # Selecionando o código dos itens
        material_list = filtered_df['Material'].tolist()

        # Tirando itens duplicados
        unique_materials = list(set(material_list))

        # Definindo caminho do arquivo .txt dos itens
        output_path = r"M:\FULFILLMENT\FILA_ERP_P\ITENS_PROCESSAR_SALDO.txt"

        # Salvando 
        with open(output_path, 'w') as file:
            for material in unique_materials:
                file.write(f"{material}\n")

        print(f"Os itens foram salvos em: {output_path}")
        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

        # Abrir SAP
        subprocess.Popen(path)
        time.sleep(3)

        # Conexão SAP
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.OpenConnection("PRD", True)
        session = connection.Children(0)

        # Inserir login
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "400"
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "BOARETO"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Pgbml12032004#!"
        session.findById("wnd[0]").sendVKey(0)

        try:
            # Tentar selecionar a opção de logon múltiplo
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            # Se ocorrer um erro, continuar normalmente sem fazer nada
            pass

        # Executar V_v02
        time.sleep(3)

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nv_v2"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/chkP_SIMUL").selected = False
        session.findById("wnd[0]/usr/radP_SCHLV").select()
        session.findById("wnd[0]/usr/radP_SCHLV").setFocus()
        session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"M:\FULFILLMENT\FILA_ERP_P\ITENS_PROCESSAR_SALDO.txt"
        session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
        session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
        session.findById("wnd[2]").sendVKey (0)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

    else:
        print("Não tem itens com saldo zerado")

    time.sleep(5)
    subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])

def cancelar_abaixo_400():
    
    def cancelar_ov(pedido, session, base_teste_filtered):
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = pedido
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 7
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        #tentar entrar na ov, se tiver bloqueada pular para a proxima
        try:
            session.findById("wnd[0]/tbar[1]/btn[34]").press()
        
            time.sleep(2)
            session.findById("wnd[1]/usr/cmbRV45A-S_ABGRU").key = "YW"
            time.sleep(1)

            # Cancelar os itens para cada linha com o mesmo pedido
            for index, row in base_teste_filtered.iterrows():
                try:
                    session.findById("wnd[1]/tbar[0]/btn[7]").press()
                    time.sleep(1)
                    try:
                        session.findById("wnd[2]").sendVKey(0)
                    except Exception:
                        pass 
                    try:
                        session.findById("wnd[2]/tbar[0]/btn[0]").press()
                    except Exception as e:
                        print(f"Erro ao executar próxima ação: {str(e)}")
                except Exception as e:
                    print(f"Erro ao cancelar item do pedido {pedido}: {str(e)}")
                
                session.findById("wnd[0]/tbar[0]/btn[11]").press()    
            # Gravar o pedido
            time.sleep(2)
            session.findById("wnd[0]/tbar[0]/btn[11]").press()

        except Exception:
            pass

    # CAMINHO SAP
    path_sap = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

    # LENDO A BASE TESTE
    base_liberacao = r"M:\FULFILLMENT\FILA_ERP_P\LIBERAÇÃO_OV.xlsx"
    base_teste = pd.read_excel(base_liberacao)

    # FILTRANDO O STATUS 'Abaixo de 400'
    base_teste_filtered = base_teste[base_teste['Status'] == 'Abaixo de 400']

    # Verificar se há pedidos com status 'Abaixo de 400'
    if not base_teste_filtered.empty:
        # Abrir o SAP
        subprocess.Popen(path_sap)
        time.sleep(3)

        # Fazer login no SAP
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.OpenConnection("PRD", True)
        session = connection.Children(0)

        # LOGIN
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "400"
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "BOARETO"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Pgbml12032004#!"
        session.findById("wnd[0]").sendVKey(0)

        try:
            # Tentar selecionar a opção de logon múltiplo
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            # Se ocorrer um erro, continuar normalmente sem fazer nada
            pass
        

        # PARA CADA OV COM O STATUS 'Abaixo de 400', CHAMAR A FUNÇÃO
        for pedido in base_teste_filtered['Documento de vendas'].unique():
            cancelar_ov(pedido, session, base_teste_filtered[base_teste_filtered['Documento de vendas'] == pedido])

        # FECHAR
        time.sleep(4)
        subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])
    else:
        print("Não temos pedidos 'Abaixo de 400'")

def travar_lotacao():

        def alterar_transp(pedido):
    # ENTRAR NA ORDEM DE VENDA E EXIBIR FLUXO
        
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = pedido
            session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 7
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[9]").select()
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER_EXT[1,3]").text = "680874"
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER_EXT[1,3]").setFocus()
            session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER_EXT[1,3]").caretPosition = 6
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
        # CAMINHO SAP
        path_sap = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

        # LENDO A BASE TESTE
        base_liberacao = r"M:\FULFILLMENT\FILA_ERP_P\LIBERAÇÃO_OV.xlsx"
        base_teste = pd.read_excel(base_liberacao)

        # FILTRANDO O STATUS LOTACAO
        pedidos_lotacao = base_teste.loc[base_teste['Status'] == 'Lotaçao', 'Documento de vendas'].unique()

        # Verificar se há pedidos com status 'Lotacao'
        if len(pedidos_lotacao) > 0:
            # Abrir o SAP
            subprocess.Popen(path_sap)
            time.sleep(3)

            # Fazer login no SAP
            sapguiauto = win32com.client.GetObject("SAPGUI")
            application = sapguiauto.GetScriptingEngine
            connection = application.OpenConnection("PRD", True)
            session = connection.Children(0)

            # LOGIN
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "400"
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "BOARETO"
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Pgbml12032004#!"
            session.findById("wnd[0]").sendVKey(0)

            try:
                # Tentar selecionar a opção de logon múltiplo
                session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
                session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except Exception:
                # Se ocorrer um erro, continuar normalmente sem fazer nada
                pass


            # PARA CADA OV COM O STATUS LOTACAO, CHAMAR A FUNÇÃO
            for pedido in pedidos_lotacao:
                alterar_transp(pedido)

            # FECHAR
            time.sleep(4)
            subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])
        else:
            print("Não temos pedidos para travar em lotacao")

def salvar_ovs_ag_liberacao():

    # Definir caminho da base
    caminho_base = r"M:\FULFILLMENT\FILA_ERP_P\LIBERAÇÃO_OV.xlsx"

    # Ler a base
    base = pd.read_excel(caminho_base)

    # Remover OVs duplicadas
    coluna_ov = "Documento de vendas"
    base_ = base.drop_duplicates(coluna_ov)

    # Selecionar as OVs com o status Remessa Liberada
    coluna_status = "Status"    
    criterio = "Remessa Liberada"
    liberar = base_.loc[base_[coluna_status] == criterio, coluna_ov]

    # Excluir nome da coluna
    liberar_ = liberar.to_string(index=False)

    # Informar o caminho que vai ficar o arquivo .txt
    liberarov = r"M:\FULFILLMENT\FILA_ERP_P\gerar_remessa.txt"

    # Atualizar o .txt com as OVs aguardando liberação
    with open(liberarov, 'w+') as arquivo:
        arquivo.write(liberar_)

def salvar_ovs_ag_faturamento():
    # Definir caminho da base
    caminho_base = r"M:\FULFILLMENT\FILA_ERP_P\LIBERAÇÃO_OV.xlsx"

    # Ler a base
    base = pd.read_excel(caminho_base)

    # Remover OVs duplicadas
    coluna_ov = "Documento de vendas"
    base_ = base.drop_duplicates(coluna_ov)

    # Selecionar as OVs com o status Remessa Liberada e tipo de documento "YRCS"
    coluna_status = "Status"
    coluna_tipo_doc = "Tipo doc.vendas"
    criterio_status = "Remessa Liberada"
    criterio_tipo_doc = "YRCS"
    liberar = base_.loc[(base_[coluna_status] == criterio_status) & (base_[coluna_tipo_doc] == criterio_tipo_doc), coluna_ov]

    # Excluir nome da coluna
    liberar_ = liberar.to_string(index=False)

    # Informar o caminho que vai ficar o arquivo .txt
    liberar_txt = r"M:\FULFILLMENT\FILA_ERP_P\VENDA.OP.TG.OV.txt"

    # Atualizar o .txt com as OVs aguardando liberação
    with open(liberar_txt, 'w+') as arquivo:
        arquivo.write(liberar_)

def gerar_remessas_vl10c():

    # Informar o caminho do SAP
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    
    # Abrir o SAP
    subprocess.Popen(path)
    time.sleep(3)
    
    # Conexão com o SAP
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.OpenConnection("PRD", True)
    session = connection.Children(0)
    
    # Inserir login
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "400"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "BOARETO"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Pgbml12032004#!"
    session.findById("wnd[0]").sendVKey(0)
    
    try:
        # Tentar selecionar a opção de logon múltiplo
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
        session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except Exception:
        # Se ocorrer um erro, continuar normalmente sem fazer nada
        pass



    # Executar VL10C
    time.sleep(3)
    
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl10c"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 4
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "4"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB2").select()
    session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB2/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1020/btn%_ST_VBELN_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[23]").press()
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"M:\FULFILLMENT\FILA_ERP_P\gerar_remessa.txt"
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[2]").sendVKey (0)
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    session.findById("wnd[0]/tbar[1]/btn[19]").press()
    
    # Encerrar o SAP
    time.sleep(5) 
    subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])

def faturar_triangulacao_vf04():
    # Caminho da base
    excel_path = r"M:\FULFILLMENT\FILA_ERP_P\LIBERAÇÃO_OV.xlsx"
    df = pd.read_excel(excel_path)

    # Selecionando as ordens de venda da operação triangular que vamos faturar
    filtered_df = df[(df['Status'] == 'Remessa Liberada') & (df['Tipo doc.vendas'] == 'YRCS')]

    if not filtered_df.empty:

        path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(3)

        # Fazer login em produção
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.OpenConnection("PRD", True)
        session = connection.Children(0)

        # Inserir credencial    
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "400"
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "BOARETO"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Pgbml12032004#!"
        session.findById("wnd[0]").sendVKey(0)

        try:
            # Tentar selecionar a opção de logon múltiplo
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            # Se ocorrer um erro, continuar normalmente sem fazer nada
            pass        


        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf04"
        session.findById("wnd[0]").sendVKey (0)
        session.findById("wnd[0]/usr/ctxtP_FKDAB").text = ""
        session.findById("wnd[0]/usr/txtS_VBELN-HIGH").setFocus()
        session.findById("wnd[0]/usr/txtS_VBELN-HIGH").caretPosition = 0
        session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[23]").press()
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = r"M:\FULFILLMENT\FILA_ERP_P\VENDA.OP.TG.OV.txt"
        session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
        session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
        session.findById("wnd[2]").sendVKey (0)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/usr/tabsTABSTRIP_A_TS/tabpTSSE/ssub%_SUBSCREEN_A_TS:SDBILLDL:8001/chkP_ALLEA").selected = True
        session.findById("wnd[0]/usr/tabsTABSTRIP_A_TS/tabpTSSE/ssub%_SUBSCREEN_A_TS:SDBILLDL:8001/chkP_ALLEL").selected = True
        session.findById("wnd[0]/usr/tabsTABSTRIP_A_TS/tabpTSSE/ssub%_SUBSCREEN_A_TS:SDBILLDL:8001/chkP_ALLEL").setFocus()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell (-1,"")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("SELKZ")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("FKTYP")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("VKORG")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("FKDAT")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("KUNNR")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("FKART")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("LLAND")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("VBELN")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("VTWEG")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("SPART")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("VBTYP")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("ADRNR")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("NAME1")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("ORT01")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("SORTKRI")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("KZPOS")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("ZAEHL")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("VSTEL")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("PDSTK")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("NETWR")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("WAERK")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("VBART")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("VBARTBEZ")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("BDR_REF")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("BDR_REF_LOGSYS")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("DRAFT_TYPE")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("DBD_REF")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("SOLUTION_ORDER_ID")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("DUMMY")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("V_FKDAT")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("V_FKART")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("STATF")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn ("SAMMG")
        session.findById("wnd[0]/tbar[1]/btn[9]").press()


        # Encerrar o SAP
        time.sleep(5)
        subprocess.call(['taskkill', '/F', '/IM', 'saplogon.exe'])

    else:
        print("NÃO TEMOS VENDA OP TRIANGULAR")

#EXPORTAR LISTA DAS ORDENS DE VENDA
#relatorio_va05()

#OBTER OS ESTADOS DOS CLIENTES
#obter_ufs_cidades_cep()

#ANALISAR AS ORDENS DE VENDA, DEFININDO STATUS
#analise_das_ovs()

#CONFIRMAR QTD. ITENS COM SALDO ZERADO 
#reprocessar_saldo_zerado()

#CANCELAR OS PEDIDOS ABAIXO DE R$400,00
#cancelar_abaixo_400()

#TRAVAR OS PEDIDOS CLIENTES MAC EM LOTAÇÃOA
#travar_lotacao()

#SALVAR O .TXT DAS OVS AGUARDANDO LIBERAÇAO
#salvar_ovs_ag_liberacao()

#SALVAR ORDENS DO PROCESSO DE TRIANGULGULAÇAO
#salvar_ovs_ag_faturamento()

#GERAR AS REMESSAS DAS ORDENS DE VENDA AGUARDDANDO LIBERAÇÃO
#gerar_remessas_vl10c()

#FATURAR OS PEDIDOS VENDA ORDEM, PROCESSO DE TRIANGULAÇAO
#faturar_triangulacao_vf04()

#lista das funcoes se for executar em looping
fila_erp = [
    relatorio_va05, obter_ufs_cidades_cep, analise_das_ovs,
    cancelar_abaixo_400, salvar_ovs_ag_liberacao,
    salvar_ovs_ag_faturamento, gerar_remessas_vl10c,
    faturar_triangulacao_vf04
]

def avisar_que_parou():

    def enviar_email(data_hoje):
        try:
            import win32com.client as win32

            # Inicializar o Outlook
            outlook = win32.Dispatch('outlook.application')
            
            # Criar um novo e-mail
            mail = outlook.CreateItem(0)
            mail.Subject = f"Erro ao executar Robô fila ERP - ({data_hoje})"
            
            # Corpo do e-mail em HTML
            email_body = f"<html><body>{mensagem_inicial()}</body></html>"
            
            # Adicionar o corpo do e-mail formatado em HTML
            mail.HTMLBody = email_body
            
            # Lista de destinatários de e-mail
            destinatarios = [
                "ana.morbidelli@grupomulti.com.br",
                "natalia.silva@grupomulti.com.br",
                "mel.pinho@grupomulti.com.br",
                "geferson.pinheiro@grupomulti.com.br",
                "juliana.tavares@grupomulti.com.br",
                "jhosefh.moreira@grupomulti.com.br",          
                "ruan.pereira@grupomulti.com.br",      
                "tamyres.oliveira@grupomulti.com.br",      
                "elleyn.sampaio@grupomulti.com.br",      
                "mateus.alvarenga@grupomulti.com.br",
                "rafaela.batista@grupomulti.com.br",
                "jenifer.oliveira@grupomulti.com.br",
                "camila.costa@grupomulti.com.br",
                "fernanda.ribeiro@grupomulti.com.br",
            ]

            # Converter a lista de destinatários em uma única string separada por ponto e vírgula
            destinatarios_str = "; ".join(destinatarios)

            # Atribuindo os destinatários ao campo "To" do e-mail
            mail.To = destinatarios_str

            # Enviar o e-mail
            mail.Send()
            
            # Retornar True se o e-mail foi enviado com sucesso
            return True
        except Exception as e:
            # Imprimir o erro
            print(f"Erro ao enviar e-mail: {e}")
            # Retornar False se houver um erro
            return False

    # Função para criar a mensagem inicial do e-mail
    def mensagem_inicial():
        return "Olá, tudo bem?<br><br>Informo que houve um erro ao tentar executar o robô da fila ERP, por gentileza realizar as tratativas manualmente.<br><br>"

    # Função para criar a mensagem final do e-mail
    def mensagem_final():
        return "<br><br>Atenciosamente,<br>Pedro Gudryan"

    # Verificar se o e-mail foi enviado com sucesso
    data_hoje = datetime.now().strftime("%d-%m-%Y")
    if enviar_email(data_hoje):
        print(f"E-mail enviado com sucesso para os destinatários.")
    else:
        print('Não foi possível enviar o e-mail.')

def rodar_job():
    max_tentativas = 2  # Número máximo de tentativas antes de considerar um erro crítico
    tentativas_erro = {func.__name__: 0 for func in fila_erp}  # Dicionário para contar os erros por função

    while True:
        for funcao in fila_erp:
            try:
                print(f'Executando {funcao.__name__}...')
                funcao()
                print(f'{funcao.__name__} executado com sucesso')
                tentativas_erro[funcao.__name__] = 0  # Reseta a contagem de erros após sucesso
            except Exception as e:
                tentativas_erro[funcao.__name__] += 1
                print(f'Erro ao executar {funcao.__name__}: {str(e)}')
                # Verifica se a função é 'gerar_remessas_vl10c' e se o número de tentativas falhas atingiu o limite
                if funcao.__name__ == 'gerar_remessas_vl10c' and tentativas_erro[funcao.__name__] >= max_tentativas:
                    print(f'Erro crítico ao executar gerar_remessas_vl10c, alertando o time...')
                    avisar_que_parou()  # Envia e-mail apenas para erros na função específica e após tentativas repetidas
                continue

        print('Aguardando leadtime para rodar job novamente')
        time.sleep(900)  # Aguarda 15 minutos antes de tentar novamente

rodar_job() 

#relatorio_va05()

#obter_ufs_cidades_cep()

#analise_das_ovs()
