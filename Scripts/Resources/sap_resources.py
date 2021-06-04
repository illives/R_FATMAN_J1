import sys, win32com.client, subprocess, time, os
import pandas as pd
import numpy as np
from time import sleep
from .gui_resources import gui

class sap_tools:
    
    def __init__(self, user, password, data, sap_path, file_path):
        self.user = user 
        self.password = password
        self.data = data
        self.sap_path = sap_path
        self.file_path = file_path
    
    def J1B1N(user, password, sap_path, file_path):
        """
        Acessar a J1B1N do SAP para criar NFs de acordo com planilha.\n
        (User = Usuario SAP,\n
        password = senha sap,\n
        sap_path = Diretório SAP,\n
        file_path = Diretório dos arquivos)
        """
        try:
        
            subprocess.Popen(sap_path)
            sleep(2)

            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return

            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return
            connection = application.OpenConnection("001 - SAP PRODUÇÃO CLARO BRASIL – CLICAR AQUI")

            if not type(connection) == win32com.client.CDispatch:
                application = None
                SapGuiAuto = None
                return

            session = connection.Children(0)
            if not type(session) == win32com.client.CDispatch:
                connection = None
                application = None
                SapGuiAuto = None
                return

            try:
                df = pd.read_excel(file_path + 'Base_baixa.xlsx')
                ultima = len(df)
                session.findById("wnd[0]").maximize
                session.findById("wnd[0]/usr/txtRSYST-BNAME").text = (user)
                session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = (password)
                session.findById("wnd[0]").sendVKey(0)
                lista = []
                texto = ('ctg_nf ; empresa ; loc_neg ; func_parceiro ; id_parceiro ; t_item_nf ; material ; centro ; qtd ; preco ; cfop ; dir_fisc_icm ; dir_fisc_ipi ; confins ; pis ; tipo_imposto_01 ; tipo_imposto_02 ; tipo Imposto_03 ; mont_basico_ICM3 ; mont_basico_ICS3 ; taxa_imposto_0102 ; taxa_imposto_03 ; outra_base ; msg ; doc_num')
                lista.append(texto)
                for k in range(0,ultima):
                    ctg_nf = df.iloc[k,0]
                    empresa = df.iloc[k,1]
                    if len(str(empresa)) <= 1:
                        empresa = str(f'00{empresa}')
                    loc_neg = df.iloc[k,2]
                    func_parceiro = df.iloc[k,3]
                    id_parceiro = df.iloc[k,4]
                    t_item_nf = df.iloc[k,5]
                    material = df.iloc[k,6]
                    centro = df.iloc[k,7]
                    qtd = df.iloc[k,8]
                    preco = df.iloc[k,9]
                    cfop = df.iloc[k,10]
                    dir_fisc_icm = df.iloc[k,11]
                    dir_fisc_ipi = df.iloc[k,12]
                    confins = df.iloc[k,13]
                    pis = df.iloc[k,14]
                    tipo_imposto_01 = df.iloc[k,15]
                    tipo_imposto_02 = df.iloc[k,16]
                    try:
                        tipo_imposto_03 = df.iloc[k,17]
                    except:
                        tipo_imposto_03 = ''
                    mont_basico_ICM3 = df.iloc[k,18]
                    mont_basico_ICS3 = df.iloc[k,19]
                    taxa_imposto_0102 = df.iloc[k,20]
                    taxa_imposto_03 = df.iloc[k,21]
                    outra_base = df.iloc[k,22]
                    msg = df.iloc[k,23]
                    Calculo_ICM3 = df.iloc[k,26]
                    Calculo_ICS3 = df.iloc[k,27]
                    session.findById("wnd[0]/tbar[0]/okcd").text = "J1B1N"
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-NFTYPE").text = ctg_nf
                    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-BUKRS").text = empresa
                    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-BRANCH").text = loc_neg
                    session.findById("wnd[0]/usr/cmbJ_1BDYDOC-PARVW").key = "WE"
                    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-PARID").text = id_parceiro
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4").select()
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB4/ssubHEADER_TAB:SAPLJ1BB2:2400/tblSAPLJ1BB2MESSAGE_CONTROL/txtJ_1BDYFTX-MESSAGE[0,0]").text = msg
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5").select()
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/ctxtJ_1BDYDOC-INCO1").text = "CIF"
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/txtJ_1BDYDOC-ANZPK").text = "1"
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/txtJ_1BDYDOC-NTGEW").text = "100"
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB5/ssubHEADER_TAB:SAPLJ1BB2:2500/txtJ_1BDYDOC-BRGEW").text = "100"
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1").select()
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-ITMTYP[1,0]").text = t_item_nf
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-MATNR[2,0]").text = material
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-WERKS[3,0]").text = centro
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/txtJ_1BDYLIN-MENGE[6,0]").text = qtd
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/txtJ_1BDYLIN-NETPR[9,0]").text = preco
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-CFOP[13,0]").text = cfop
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW1[14,0]").text = dir_fisc_icm
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW2[15,0]").text = dir_fisc_ipi
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW4[16,0]").text = confins
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-TAXLW5[17,0]").text = pis
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]").sendVKey(0)
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL").getAbsoluteRow(0).selected = True
                    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/btn%#AUTOTEXT002").press()
                    session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/ctxtJ_1BDYSTX-TAXTYP[0,0]").text = tipo_imposto_01
                    try:
                        session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/ctxtJ_1BDYSTX-TAXTYP[0,2]").text = tipo_imposto_02
                    except:
                        pass
                    try:
                        session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/ctxtJ_1BDYSTX-TAXTYP[0,1]").text = tipo_imposto_03
                        session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-RATE[4,1]").text = taxa_imposto_03
                        session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-TAXVAL[5,1]").text = Calculo_ICS3
                    except:
                        pass
                    try:
                        session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-BASE[3,1]").text = mont_basico_ICS3
                    except:
                        pass
                    session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-BASE[3,0]").text = mont_basico_ICM3
                    session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-TAXVAL[5,0]").text = Calculo_ICM3
                    session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-RATE[4,0]").text = taxa_imposto_0102
                    session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-OTHBAS[7,0]").text = ""
                    session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-OTHBAS[7,2]").text = outra_base
                    #session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL/txtJ_1BDYSTX-RATE[4,2]").text = taxa_imposto_03
                    session.findById("wnd[0]").sendVKey (0)
                    session.findById("wnd[0]").sendVKey (0)
                    session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL").getAbsoluteRow(0).selected = True
                    session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL").getAbsoluteRow(1).selected = True
                    session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/tblSAPLJ1BB2TAX_CONTROL").getAbsoluteRow(2).selected = True
                    #session.findById("wnd[0]/usr/tabsITEM_TAB/tabpTAX/ssubITEM_TABS:SAPLJ1BB2:3200/btnPB_CALCULATOR").press()
                    session.findById("wnd[0]/tbar[0]/btn[11]").press()
                    doc_num = session.findById("wnd[0]/sbar").text
                    doc_num = doc_num.split()
                    doc_num = str(doc_num[2])
                    #taxa_imposto = df.iloc[k,18]
                    texto = (f'{ctg_nf} ; {empresa} ; {loc_neg} ; {func_parceiro} ; {id_parceiro} ; {t_item_nf} ; {material} ;{centro} ; {qtd} ; {preco} ; {cfop} ; {dir_fisc_icm} ; {dir_fisc_ipi} ; {confins} ; {pis} ; {tipo_imposto_01} ; {tipo_imposto_02} ; {tipo_imposto_03} ; {mont_basico_ICM3}  ; {mont_basico_ICS3} ; {taxa_imposto_0102} ; {taxa_imposto_03} ; {outra_base} ; {msg} ; {doc_num}')
                    lista.append(texto)
                    session.findById("wnd[0]").sendVKey (3)
                df = pd.DataFrame(lista)
                df.to_csv(file_path + 'Base_baixa.csv', sep = ';', encoding= 'UTF-8', index = False, header = 0)
            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print(exc_type, fname, exc_tb.tb_lineno)
                df = pd.DataFrame(lista)
                df.to_csv(file_path + 'Base_baixa.csv', sep = ' ; ', encoding= 'UTF-8', index = False, header = 0)
            finally:
                os.system("taskkill /f /im saplogon.exe")
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
    
    def J1BNFE(user, password, sap_path, file_path):
        """
        Acessar a J1BNFE do SAP para consultar NFs de acordo com planilha.\n
        (User = Usuario SAP,\n
        password = senha sap,\n
        sap_path = Diretório SAP,\n
        file_path = Diretório dos arquivos)
        """
        try:
        
            subprocess.Popen(sap_path)
            sleep(2)

            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return

            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return
            connection = application.OpenConnection("001 - SAP PRODUÇÃO CLARO BRASIL – CLICAR AQUI")

            if not type(connection) == win32com.client.CDispatch:
                application = None
                SapGuiAuto = None
                return

            session = connection.Children(0)
            if not type(session) == win32com.client.CDispatch:
                connection = None
                application = None
                SapGuiAuto = None
                return

            try:
                df = pd.read_csv(file_path + 'Base_baixa.csv', encoding = 'UTF-8', sep = ' ; ', header = None, engine = 'python', skiprows = range(1))
                df.rename(columns = {23:"doc_num"}, inplace = True)
                df.rename(columns = {0:"ctg_nf"}, inplace = True)
                df["doc_num"] = df["doc_num"].str.replace(r'"','')
                df["ctg_nf"] = df["ctg_nf"].str.replace(r'"','')
                df["doc_num"] = df["doc_num"].astype(int)
            except:
                df = pd.read_csv(file_path + 'Base_baixa.csv', encoding = 'UTF-8', sep = ' ; ', engine = 'python', header = 0)
            lista = df["doc_num"].unique()
            lista = lista.tolist()
            df_list = pd.DataFrame(lista)
            df_list.to_clipboard(index = False, header = None)
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = (user)
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = (password)
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/okcd").text = "j1bnfe"
            session.findById("wnd[0]").sendVKey (0)
            session.findById("wnd[0]/usr/ctxtBUKRS-LOW").text = "001"
            session.findById("wnd[0]/usr/btn%_DOCNUM_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[16]").press()
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]").sendVKey (8) 
            #gui.waiting_frame()
            session.findById("wnd[0]").sendVKey (5)
            cont = 0
            while True:
                try:
                    session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").currentCellRow = cont
                    cont +=1
                except:
                    break   
            session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").selectColumn ("DOCNUM")
            session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").selectColumn ("CODE")
            Tabla = session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell")
            Tabla.contextMenu()
            sleep(3)
            Tabla.selectContextMenuItemBytext ("Copiar texto")
            sleep(1)
            Tabla.selectContextMenuItemBytext ("Copiar texto")
            df2 = pd.read_clipboard(header = None)
            df2.rename(columns = {0:"NF", 1:"status"}, inplace = True)
            df2 = df2[df2.status == 100]
            lista = df2["NF"].unique()
            lista = lista.tolist()
            df['status'] = np.where((df['doc_num'].isin(lista)), '100', '')
            cab = ["ctg_nf" , "empresa" , 'loc_neg' , 'func_parceiro' , 'id_parceiro' , 't_item_nf' , 'material__centro' , 'qtd' , 'preco' , 'cfop' , 'dir_fisc_icm' , 'dir_fisc_ipi' , 'confins' , 'pis' , 'tipo_imposto_01' , 'tipo_imposto_02' , 'tipo_imposto_03', 'mont_basico_ICM3' , 'mont_basico_ICS3' , 'taxa_imposto_0102' , 'taxa_imposto_03' , 'outra_base' , 'msg', 'doc_num', 'status']
            try:
                df.to_csv (file_path + 'Base_baixa.csv', sep = ';', encoding= 'UTF-8', index = False, header = cab)
            except:
                df.to_csv (file_path + 'Base_baixa.csv', sep = ';', encoding= 'UTF-8', index = False)
            session.findById("wnd[0]").sendVKey (3)
            session.findById("wnd[0]").sendVKey (3)
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
            os.system("taskkill /f /im saplogon.exe")
        finally:
            os.system("taskkill /f /im saplogon.exe")
            os.system("taskkill /f /im excel.exe")
        
    def J1B3N(user, password, sap_path, file_path):
        """
        Acessar a J1B3N do SAP para imprimir NFs de acordo com planilha.\n
        (User = Usuario SAP,\n
        password = senha sap,\n
        sap_path = Diretório SAP,\n
        file_path = Diretório dos arquivos)
        """
        try:
        
            subprocess.Popen(sap_path)
            sleep(2)

            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return

            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return
            connection = application.OpenConnection("001 - SAP PRODUÇÃO CLARO BRASIL – CLICAR AQUI")

            if not type(connection) == win32com.client.CDispatch:
                application = None
                SapGuiAuto = None
                return

            session = connection.Children(0)
            if not type(session) == win32com.client.CDispatch:
                connection = None
                application = None
                SapGuiAuto = None
                return

            df = pd.read_csv(file_path + 'Base_baixa.csv', encoding = 'UTF-8', sep = ';', header = 0)
            ultima = len(df)
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = (user)
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = (password)
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").text = "J1B3N"
            session.findById("wnd[0]").sendVKey (0)
            for k in range(0, ultima):
                status = str(df.iloc[k,24])
                doc_num = df.iloc[k,23]
                if status == '100':
                    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-DOCNUM").text = ''
                    session.findById("wnd[0]/usr/ctxtJ_1BDYDOC-DOCNUM").text = doc_num
                    session.findById("wnd[0]/mbar/menu[0]/menu[8]").select()
                    session.findById("wnd[0]").sendVKey (0)
                    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
        finally:
            os.system('taskkill /f /im excel.exe')
            os.system('taskkill /f /im saplogon.exe')