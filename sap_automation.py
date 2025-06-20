
import win32com.client
import time
import os
import pandas as pd
import pyautogui

# CONFIGURAÇÕES (ajuste conforme seu ambiente local)
saplogon_path = r"C:\Path\To\SAP\SAPLogon.exe"
credentials_path = r"C:\Path\To\Credentials\SAP_Credentials.xlsx"
export_folder = r"C:\Path\To\Export"
sap_connection_name = "SAP_CONNECTION_NAME"

# Carregar credenciais de usuário a partir do Excel
credentials_df = pd.read_excel(credentials_path)
user = credentials_df['usuario'][0]
password = credentials_df['senha'][0]

# Remover arquivo exportado anterior (se existir)
try:
    os.remove(os.path.join(export_folder, "ODS_CRIADAS.xlsx"))
except FileNotFoundError:
    print("Nenhum arquivo antigo encontrado para remoção.")

# Iniciar SAP Logon
os.startfile(saplogon_path)
time.sleep(10)

# Conectar ao SAP via SAP GUI Scripting
SapGui = win32com.client.GetObject("SAPGUI")
application = SapGui.GetScriptingEngine
connection = application.OpenConnection(sap_connection_name, True)
time.sleep(5)
session = connection.Children(0)
session.findById("wnd[0]").maximize()

# Fazer login
session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "400"
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
session.findById("wnd[0]").sendVKey(0)
print("Login realizado com sucesso!")

# Executar transação VA05N
session.StartTransaction(Transaction="VA05N")
time.sleep(2)

# Chamar variante (exemplo genérico, adapte conforme necessidade)
session.findById("wnd[0]").sendVKey(17)
alv_table = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")

# Buscar variante desejada
variant_to_select = "VARIANT_NAME"
num_rows = alv_table.RowCount
for i in range(num_rows):
    if alv_table.GetCellValue(i, 'VARIANT') == variant_to_select:
        alv_table.selectedRows = str(i)
        break

session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = alv_table.selectedRows
session.findById("wnd[1]").sendVKey(2)
time.sleep(2)
session.findById("wnd[0]").sendVKey(8)

# Exportar resultados
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = export_folder
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ODS_CRIADAS.xlsx"
session.findById("wnd[1]/tbar[0]/btn[0]").press()
time.sleep(15)

# Fechar Excel caso tenha aberto automaticamente
pyautogui.hotkey('alt', 'f4')

# Encerrar sessão SAP
print("Encerrando SAP...")
session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
session.findById("wnd[0]").sendVKey(0)

# Processar o arquivo exportado
output_file = os.path.join(export_folder, "ODS_CRIADAS.xlsx")
df = pd.read_excel(output_file, skipfooter=2)
df.columns = ['Doc.SD', 'TpDV', 'Denominação', 'EmissorOrd', 'Dt/criação', 'Data doc/', 'Valor líquido', 'Criado por']

df['Dt/criação'] = pd.to_datetime(df['Dt/criação'], errors='coerce').dt.strftime('%d.%m.%Y')
df['Data doc/'] = pd.to_datetime(df['Data doc/'], errors='coerce').dt.strftime('%d.%m.%Y')

final_file = os.path.join(export_folder, "Base_Ordens_Devolucao.xlsx")
df.to_excel(final_file, index=False)
print("Processo concluído com sucesso!")
