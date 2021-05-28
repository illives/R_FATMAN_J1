import sys, os
from Resources.gui_resources import gui
from Resources.json_resource import json_tool
from Resources.sap_resources import sap_tools



def Main():
  try:
    gui.main_frame()
    pwd = os.getcwd()
    path = f'{pwd[:-7]}My_docs\\status.json'
    data = json_tool.read_json(path)
    status = data["STATUS"]
    if status == "ATIVO":
      os.system('taskkill /f /im saplogon.exe')
      os.system('taskkill /f /im excel.exe')
      path = f'{pwd[:-7]}My_docs\\index.json'
      data = json_tool.read_json(path)
      sap_tools.J1B1N(data["USER_SAP"], data["PASSWORD_SAP"], data["SAP_PATH"], data["DESTIN_PATH"])
      gui.waiting_frame()
      sap_tools.J1BNFE(data["USER_SAP"], data["PASSWORD_SAP"], data["SAP_PATH"], data["DESTIN_PATH"])
      sap_tools.J1B3N(data["USER_SAP"], data["PASSWORD_SAP"], data["SAP_PATH"], data["DESTIN_PATH"])
    elif status == "INATIVO":
      pass
  except Exception as e:
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    print(exc_type, fname, exc_tb.tb_lineno)
  finally:
    quit()

if __name__ == "__main__":
  Main()


