# Programa para calcular ventas en catnidad de partes y USD, acomdo de partes en los alamacenes, 
# Diferencia Due-date Calculado vs. Due-Date en Sistema de Produccion
# Abril 5/ 2023

import pandas as pd
import numpy as np
import xlwt
import openpyxl
from datetime import datetime
#import dt
import warnings
import os
import pygsheets
import glob
import chardet

from oauth2client.service_account import ServiceAccountCredentials
import json
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import gspread_dataframe as gd
import gspread
from gspread_dataframe import set_with_dataframe
from gooey import Gooey, GooeyParser

warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)

scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]

credentials = ServiceAccountCredentials.from_json_keyfile_name("compras-380500-0f03f7c142be.json", scopes) 
#file = gspread.authorize(credentials)# authenticate the JSON key with gspread
file = pygsheets.authorize(service_file="compras-380500-0f03f7c142be.json")
#gc = pygsheets.authorize(service_file='/Users/erikrood/desktop/QS_Model/creds.json')
ss = file.open('EficienciaInventarios')

print(ss)
##V3m = ss["Ventas3m"] 
V3m = ss[0]     
#V3m = ss.worksheet("Ventas3m") 
#V1y = ss.worksheet("Ventas1y") 
##V1y = ss["Ventas1y"]
V1y = ss[1]
#InDB = ss.worksheet("InventoryDB")
##InDB = ss["InventoryDB"]
InDB = ss[4]
#AcoDiario = ss.worksheet("AcomodoDiario")
##AcoDiario = ss["AcomodoDiario"]
AcoDiario = ss[3]
#Vdiaria = ss.worksheet("Venta Diaria")
#Vdiaria = ss["Venta Diaria"]
Vdiaria = ss[2]
#VentasDiaframe = pd.DataFrame()
#VentasDiaframe['name'] = ['Pedro', 'Steve', 'Sarah','jose']
#print(VentasDiaframe)
#print(Vdiaria)
# Obtiene la lista de hojas
#sheets_list = ss.worksheets()

# Imprime el nombre de cada hoja
#for sheet in sheets_list:
#    print(sheet.title)
#AcoDiario.clear(start='A1', end=None, fields='*')
#AcoDiario.set_dataframe(VentasDiaframe,(0,0))
#set_with_dataframe(worksheet=Vdiaria, dataframe=VentasDiaframe, include_index=False,
#include_column_header=True, resize=True)

@Gooey(program_name="Capacidad diaria Tiendas")
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Actualiza capacidad de Almacenes de Google Sheets')
    #parser.add_argument('Archivo_CSV',
    #                    action='store',
    #                    default=stored_args.get('cust_file'),
    #                    widget='FileChooser',
    #                    help='Ej. JobsReport_20230317081108.csv')
    parser.add_argument('Directorio_de_trabajo',
                        action='store',
                        default=stored_args.get('data_directory'),
                        widget='DirChooser',
                        help="Directorio con los archivos .XLSX/.CSV ")
    #parser.add_argument('Archivo_Corte',
    #                    action='store',
    #                    default=stored_args.get('cust_file'),
    #                    widget='FileChooser',
    #                    help='Ej. horaCortesFestivos.csv')
    parser.add_argument('Fecha', help='Seleccione Fecha del Reporte',
                        default=stored_args.get('Fecha'),
                        widget='DateChooser')
    args = parser.parse_args()
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    return args

def Principal(Directorio_de_trabajo):
	#path = os.getcwd()
	global V1y
	global V3m
	global Vdiaria
	global AcoDiario
	global InDB
	path = Directorio_de_trabajo
	xls_files = glob.glob(os.path.join(path, "*.xls"))
	csv_files = glob.glob(os.path.join(path, "*.csv"))
	print("Total de archivos .XLS: ",len(xls_files))
	print("Total de archivos .CSV: ",len(csv_files))

	dl0 = pd.DataFrame()
	dl = pd.DataFrame()
	dl8 = pd.DataFrame()
	dl9 = pd.DataFrame()
	AcomodoDiaframe = pd.DataFrame(columns=['codigo','Tienda','cantidad de partes'])
	Xfiles = []
	Cfiles= []
	tiendas=[1,2,4,6,10,15,20]
	tiendastodas=[1,2,4,6,7,8,10,15,20]
	tiendasTransito=[9,11,13,14,16,18,21]
	tiendas78=[7,8]
	for filename in xls_files:
	    df = pd.read_excel(filename,header=None)
	    Xfiles.append(df)
	    ####dl9 = pd.concat([dl9, Xfiles[j]])
	    #x= [Xfiles,df]
	    #Xfiles = pd.concat(x,ignore_index=True)




	for filename in csv_files:
	    #with open(filename, 'rb') as f:
	    #	enc = chardet.detect(f.read())
	    #df = pd.read_csv(filename,header=None,encoding = enc['encoding'])
	    df = pd.read_csv(filename,header=None,encoding='latin-1')
	    Cfiles.append(df)
	    #Cfiles = pd.concat(Cfiles,df)   

	#print(Xfiles)
	#print(Cfiles)
	#print(len(Xfiles))
	#print(Xfiles[1])
	#print(list(Xfiles[1].columns.values))


	#print(fecha)
	for j in range(len(Xfiles)):
		if Xfiles[j][0][0]=="Inventory By Location":               #Busca archivo de inventario general al dia
			xls_files
			#dlini=Xfiles[j].copy()
			dlini = pd.read_excel(xls_files[j],sheet_name=None,header=None)
			dla = pd.concat(dlini, axis=0, ignore_index=False)
			print("Archivo ",xls_files[j]," de Inventario encontrado")
		elif Xfiles[j][0][0]=="Inventory Report for Parts Sold - Summary":  #Busca archivo de ventas Anuales y 3 meses
			Xfiles[j][2] = Xfiles[j][5][0]
			#print("j=",j,"Store sold",Xfiles[j][5][0]) 
			fecha = Xfiles[j][1][1].split("-")
			fecha[0]=datetime.strptime(fecha[0], '%m/%d/%Y ').date()
			fecha[1]=datetime.strptime(fecha[1], ' %m/%d/%Y').date()
			#print(fecha) 
			#calculate difference between dates
			dias = (fecha[1] - fecha[0]) #/ np.timedelta64(1, 'D')
			#print(dias.days)
			#print(dias.dtypes)
			if dias.days in range(89,93) :
				#dl=dl.append(Xfiles[j])
				list_of_dfs = [dl, Xfiles[j]]
				dl = pd.concat(list_of_dfs, ignore_index=True)
				#dl=pd.concat(dl,Xfiles[j])
				print("Archivo ",xls_files[j]," de ventas ultimos 3 meses encontrado")
			elif dias.days==365:
				#dl8=dl8.append(Xfiles[j])
				#dl=dl.append(Xfiles[j])
				list_of_dfs1 = [dl8, Xfiles[j]]
				dl8 = pd.concat(list_of_dfs1, ignore_index=True)
				#dl8 = pd.concat(dl8,Xfiles[j])

				print("Archivo ",xls_files[j]," de ventas ultimo año encontrado")

		elif Xfiles[j][0][0]=="Inventory Audit Trail":  #Busca archivo de auditoria del dia
			fecha = Xfiles[j][1][1].split("-")
			fecha[0]=datetime.strptime(fecha[0], '%m/%d/%Y ').date()
			fecha[1]=datetime.strptime(fecha[1], ' %m/%d/%Y').date()
			#print(fecha) 
			#calculate difference between dates
			dias = (fecha[1] - fecha[0]) #/ np.timedelta64(1, 'D')
			#print(dias.days)
			if dias.days==0:
			    list_of_dfs2 = [dl9, Xfiles[j]]
			    dl9 = pd.concat(list_of_dfs2, ignore_index=True)

			    #dl9=dl9.append(Xfiles[j])
			    #dl9 = pd.concat(dl9,Xfiles[j])
			    indexDeleted = dl9[dl9[8] != 'CategorizingStoreNumber'].index
			    dl9.drop(indexDeleted,inplace=True)
			    result = dl9[0].str.extract(r'(\d{1,3})').squeeze().str.zfill(3)
			    dl9[1] = result
			    indexDeleted = dl9[dl9[1] == '957'].index
			    result = dl9[10].str.extract(r'(\d{1,4})').squeeze().str.zfill(4)
			    dl9[10] = result
			    dl9[10]=pd.to_numeric(dl9[10], errors='coerce')
			    result = dl9[11].str.extract(r'(\d{1,4})').squeeze().str.zfill(4)
			    dl9[11] = result
			    dl9[11]=pd.to_numeric(dl9[11], errors='coerce')
			    dl9.drop(indexDeleted,inplace=True)
			    #print("dl9: ",dl9)
			    #dl9.to_excel('AuditTrial.xlsx',index=False,header=True)
			    print("Archivo ",csv_files[j]," de auditoria de almacen encontrado")
			else: 
				print(" No se encontro archivo de Auditoria")
				exit()		


	for j in range(len(Cfiles)):
		#print(Cfiles[j].head())
		#print(Cfiles[j][11][0])
		if Cfiles[j][11][0]=="All":  #Busca archivo de ventas del dia 
			fecha = Cfiles[j][7][0].split("-")
			fecha[0]=datetime.strptime(fecha[0], '%m/%d/%Y ').date()
			fecha[1]=datetime.strptime(fecha[1], ' %m/%d/%Y').date()
			#calculate difference between dates
			print(fecha)
			print(fecha[0])
			print(fecha[1])
			dias = (fecha[1] - fecha[0]) #/ np.timedelta64(1, 'D')
			print(dias.days)
			if dias.days==0:
			    #dl0=dl0.append(Cfiles[j])
			    #dl=dl.append(Xfiles[j])
			    list_of_dfs3 = [dl0, Cfiles[j]]
			    dl0 = pd.concat(list_of_dfs3, ignore_index=True)
			    dl0.to_excel('dl0.xlsx',index=False,header=True)
			    dl0 = dl0.iloc[:,[3,7,46,47,48,60,61,62]]
			    #dl0.to_excel('entas_partes_diarias.xlsx',index=False,header=True)
			    #dl0.dropna(subset = [3],inplace=True)
			    #dl0 = dl0.drop(dl0.columns[[15,14,13,12,11,10,9,8,7,6,5,2,0]],axis = 1)
			    print("Archivo ",csv_files[j]," de ventas del dia encontrado")
			else: 
				print(" No se encontro archivo de Ventas diarias")
				exit()
		if Cfiles[j][11][0]==253:  #Busca archivo de ventas del dia 
			fecha = Cfiles[j][7][0].split("-")
			fecha[0]=datetime.strptime(fecha[0], '%m/%d/%Y ').date()
			fecha[1]=datetime.strptime(fecha[1], ' %m/%d/%Y').date()
			#calculate difference between dates
			dias = (fecha[1] - fecha[0]) #/ np.timedelta64(1, 'D')
			print(dias.days)
			if dias.days==365:
				dbgy = pd.read_csv(csv_files[j],header=None,encoding='latin-1')
				print("Archivo ",csv_files[j]," de ventas anual de Bolsas encontrado")
			elif dias.days in range(89,93):
				dbgm = pd.read_csv(csv_files[j],header=None,encoding='latin-1')
				print("Archivo ",csv_files[j]," de ventas ultimos 3 meses de Bolsas encontrado")
			else: 
				print(" No se encontraron archivos de ventas de Bolsas ")
				exit()

	indexDeleted = dl[dl[0] == '253'].index
	dl.drop(indexDeleted,inplace=True)
	indexDeleted = dl8[dl8[0] == '253'].index
	dl8.drop(indexDeleted,inplace=True)			

	#Summary D3  ######### AFTERMARKET Y LLANTAS ##############


	#####dlini = pd.read_excel(r'inventorylocation.xls',sheet_name=None,header=None)
	#####dla = pd.concat(dlini, axis=0, ignore_index=False)
	#with open(r'invpartsold253yoriginal.csv', 'rb') as f:
	#    	enc = chardet.detect(f.read())
	#dbgy = pd.read_csv(r'invpartsold253yoriginal.csv',header=None,encoding = enc['encoding'])encoding='latin-1'
	#################33333dbgy = pd.read_csv(r'invpartsold253y.csv',header=None,encoding='latin-1')
	#with open(r'invpartsold253moriginal.csv', 'rb') as f:
	#    	enc = chardet.detect(f.read())
	#dbgm = pd.read_csv(r'invpartsold253moriginal.csv',header=None,encoding = enc['encoding'])encoding='latin-1')
	##################dbgm = pd.read_csv(r'invpartsold253m.csv',header=None,encoding='latin-1')

	#print(list(dla.columns.values))
	dlanew = dla.copy()
	indexDeleted = dlanew[dlanew[18] == 'FOTOS DE VEICULO'].index
	dlanew.drop(indexDeleted,inplace=True)

	#indexNames = dlanew[dlanew[0].isin(['Group Totals','Grand Totals','Page -1 of 1'])].index
	#dlanew.drop(indexNames,inplace=True)

	indexDeleted = dlanew[dlanew[0] == 'Group Totals'].index
	dlanew.drop(indexDeleted,inplace=True)
	indexDeleted = dlanew[dlanew[0] == 'Grand Totals'].index
	dlanew.drop(indexDeleted,inplace=True)
	indexDeleted = dlanew[dlanew[0] == 'Page -1 of 1'].index
	dlanew.drop(indexDeleted,inplace=True)
	dlanew.dropna(subset = [1],inplace=True)
	result = dlanew[1].str.extract(r'(\d{1,3})').squeeze().str.zfill(3)
	dlanew.insert(1,'Codigo',result)
	#print(list(dlanew.columns.values))
	#dlanew.to_excel('inventorylocationwarehouses_update.xlsx', index=False,header=True)

	# Guardo en Variable "codigos" todos los codigos que estan en powerlink
	dlanewcopy = dlanew.copy()	
	dlaAudit = dlanew.copy()
	dlanewcopy = dlanewcopy.drop_duplicates(subset=['Codigo'], keep='first')
	codigos = dlanewcopy['Codigo'].tolist()
	#print('codigos',codigos)
	############ Ventas del dia anterior OK ######################33
	VentasDiaframe = pd.DataFrame(columns=['codigo','Tienda','cantidad','ventas'])
	#print(list(dl0.columns.values))
	result = dl0[47].str.extract(r'(\d{1,3})').squeeze().str.zfill(3)
	dl0[61]=dl0[61].str.replace('$','')
	dl0[61]=dl0[61].str.replace(',','')
	dl0[61]=pd.to_numeric(dl0[61], errors='coerce')
	dl0[46]=result
	#dl0.to_excel('Ventas_partes_diarias antes.xlsx',index=False,header=True)
	for k in tiendastodas:
		for j in codigos:
			if j!='253':
				Tventas=dl0[(dl0[48]==k)&(dl0[46]==j)][61].sum()
				Tpventas=dl0[(dl0[48]==k)&(dl0[46]==j)][60].count()
				VentasDiaframe.loc[len(VentasDiaframe.index)]= [j,k,Tpventas,Tventas]
			if j=='253':
				TventasB=dl0[(dl0[62].str.upper().str.contains('DASH')==False)&(dl0[46]==j)&(dl0[48]==k)][61].sum()
				TpventasB=dl0[(dl0[62].str.upper().str.contains('DASH')==False)&(dl0[46]==j)&(dl0[48]==k)][60].count()
				TventasBD=dl0[(dl0[62].str.upper().str.contains('DASH'))&(dl0[46]==j)&(dl0[48]==k)][61].sum()
				TpventasBD=dl0[(dl0[62].str.upper().str.contains('DASH'))&(dl0[46]==j)&(dl0[48]==k)][60].count()
				VentasDiaframe.loc[len(VentasDiaframe.index)]= ['253D',k,TpventasBD,TventasBD]
				VentasDiaframe.loc[len(VentasDiaframe.index)]= ['253',k,TpventasB,TventasB]

	indexDeleted = VentasDiaframe[VentasDiaframe["cantidad"] == 0].index
	VentasDiaframe.drop(indexDeleted,inplace=True)			
	VentasDiaframe.dropna(subset = ["cantidad"],inplace=True)	
	print("Creando archivo Ventas_partes_diarias.xlsx")
	VentasDiaframe.to_excel('Ventas_partes_diarias.xlsx',index=False,header=True)
	print("Archivo creado")
	#########################################################################################################
	# Limpiando Audi trial file
	#dlaAudit.to_excel('dlaAfitbe.xlsx', index=False,header=True) #copia del inventario al dia 
	indexDeleted = dlaAudit[dlaAudit['Codigo'] != '253'].index #dejando solo las 253 en copia de inventario
	dlaAudit.drop(indexDeleted,inplace=True)
	dlaAudit.drop(dlaAudit.columns[[19,18,17,16,15,14,13,12,11,10,9,8,7,6,4,1,0]],axis = 1,inplace=True)
	dlaAudit.rename(columns={dlaAudit.columns[0]: 'Intercambio', dlaAudit.columns[1]: 'Store',dlaAudit.columns[2]: 'Stock',dlaAudit.columns[3]: 'Details'},inplace=True)
	#dlaAudit.to_excel('dlaAudit.xlsx', index=False,header=True)
	####dlaAudit = dlaAudit.rename(columns={'Name': 'EmpName', 'ID': 'EmpID', 'Role': 'EmpRole'})
	dl9temp = dl9.copy() # Copia de AuditTrial
	indexDeleted = dl9temp[dl9temp[1] != '253'].index  # dejando solo las 253 en copia de audit trial
	dl9temp.drop(indexDeleted,inplace=True)
	dl9temp.rename(columns={dl9.columns[0]: 'Intercambio', dl9.columns[1]: 'Code',dl9.columns[2]: 'Stock',dl9.columns[3]: 'Year',
		                dl9.columns[4]: 'Model', dl9.columns[5]: 'Location',dl9.columns[6]: 'date',
		                dl9.columns[7]: 'User', dl9.columns[8]: 'Movement',dl9.columns[9]: 'Empty',
		                dl9.columns[10]: 'StoreOld', dl9.columns[11]: 'StoreNew'},inplace=True)
	#dl9temp.to_excel('dl9.xlsx', index=False,header=True)
	merged_data = pd.merge(dlaAudit, dl9temp, on=['Intercambio', 'Stock'], how='inner')
	#dl9temp=dl0[(dl9[0]==dlaAudit[1])&(dl9[2]==dlaAudit[4])][19]
	print("Creando archivo merged.xlsx")
	merged_data.to_excel('meged.xlsx', index=False,header=True)
	print("Archivo creado")
	##################################dl = pds.DataFrame() ##### Debug key
	for k in tiendastodas:
		for j in codigos:
			if j!='253':
				Tparte=dl9[(dl9[11]==k)&(dl9[1]==j)&(dl9[10].isin(tiendasTransito))][8].count()
				AcomodoDiaframe.loc[len(AcomodoDiaframe.index)]= [j,k,Tparte]
			if j=='253':
				TparteB=merged_data[(merged_data['StoreNew']==k)&(merged_data['Details'].str.upper().str.contains('DASH')==False)&(merged_data['StoreOld'].isin(tiendasTransito))]['User'].count()
				TparteBD=merged_data[(merged_data['StoreNew']==k)&(merged_data['Details'].str.upper().str.contains('DASH'))&(merged_data['StoreOld'].isin(tiendasTransito))]['User'].count()
				AcomodoDiaframe.loc[len(AcomodoDiaframe.index)]= ['253',k,TparteB]
				AcomodoDiaframe.loc[len(AcomodoDiaframe.index)]= ['253D',k,TparteBD] 		
	print("Creando archivo AcomodoDiaria.xlsx")
	AcomodoDiaframe.to_excel('AcomodoDiario.xlsx',index=False,header=True)
	print("Archivo creado")

	#dl = pds.DataFrame() ##### Debug key

	dbgy = dbgy.iloc[:,[43,44,48,50,61,62]]
	dbgm = dbgm.iloc[:,[43,44,48,50,61,62]]
	dbgy[61]=dbgy[61].str.replace('$','')
	dbgy[61]=dbgy[61].str.replace(',','')
	dbgy[61]=pd.to_numeric(dbgy[61], errors='coerce')
	dbgm[61]=dbgm[61].str.replace('$','')
	dbgm[61]=dbgm[61].str.replace(',','')
	dbgm[61]=pd.to_numeric(dbgm[61], errors='coerce')

	today = datetime.now()
	#print("Cdebug")
	for j in tiendas:
		TotalSoldy=dbgy[dbgy[48]==j][61].sum()
		TotalSoldqy=dbgy[dbgy[48]==j][61].count()
		#print(TotalSoldqy)
		TotalSoldm=dbgm[dbgm[48]==j][61].sum()
		TotalSoldqm=dbgm[dbgm[48]==j][61].count()
		#print(TotalSoldm)
		TotalSoldDy=dbgy[(dbgy[62].str.upper().str.contains('DASH'))&(dbgy[48]==j)][61].sum()
		TotalSoldDqy=dbgy[(dbgy[62].str.upper().str.contains('DASH'))&(dbgy[48]==j)][61].count()
		#print(TotalSoldDy)
		TotalSoldDm=dbgm[(dbgm[62].str.upper().str.contains('DASH'))&(dbgm[48]==j)][61].sum()
		TotalSoldDqm=dbgm[(dbgm[62].str.upper().str.contains('DASH'))&(dbgm[48]==j)][61].count()
		#print(TotalSoldDm)
		dl8.loc[len(dl8.index)] = ['253','AIR BAG',j, TotalSoldqy-TotalSoldDqy, TotalSoldy-TotalSoldDy,""]
		dl8.loc[len(dl8.index)] = ['253D','AIR BAG DASH',j, TotalSoldDqy, TotalSoldDy,""]
		dl.loc[len(dl.index)] = ['253','AIR BAG',j, TotalSoldqm-TotalSoldDqm, TotalSoldm-TotalSoldDm,""]
		dl.loc[len(dl.index)] = ['253D','AIR BAG DASH',j, TotalSoldDqm, TotalSoldDm,""]


	dlanew.drop(dlanew.columns[[14,12,11,10,9,8,7,6,4,2,0]],axis = 1,inplace=True)
	dlanew.drop(dlanew.index[:4], inplace=True)
	dlanew[14]= today
	dlanew[14] = dlanew[14].apply(pd.to_datetime)
	dlanew[16] = dlanew[16].apply(pd.to_datetime)
	dlanew[17]= (dlanew[14]-dlanew[16]).dt.days
	#print(list(dlanew.columns.values))
	dlanew.drop(dlanew.columns[4],axis = 1,inplace=True)
	dlanewcopy = dlanew.copy()
	#drop all except Bags (253) 
	dlanewcopy= dlanewcopy.query("Codigo == '253'")#,inplace=True)
	indexDeleted = dlanew[dlanew['Codigo'] == '253'].index
	dlanew.drop(indexDeleted,inplace=True)
	Inventarioframe = pd.DataFrame(columns=['Codigo','CantidadInventario','Tienda','365','1.5 Años'])#,'En transito'
	#VentasDiaframe = pd.DataFrame(columns=['Dummy','Intercambio','Tienda','cantidad','ventas'])

	for k in tiendas:

		TotalBQty=dlanewcopy[dlanewcopy[2]==k][12].sum() 
		#Cantidad de Bolsas Dash
		TotalBDQty=dlanewcopy[(dlanewcopy[19].str.upper().str.contains('DASH'))&(dlanewcopy[2]==k)][12].sum()
		TotalBBQty= TotalBQty - TotalBDQty 

		DiasBD365 = dlanewcopy[(dlanewcopy[2]==k)&(dlanewcopy[17].astype(int) < 366)&(dlanewcopy[19].str.upper().str.contains('DASH'))][12].count()
		DiasBD547 = dlanewcopy[(dlanewcopy[2]==k)&(dlanewcopy[17].astype(int) >365)&(dlanewcopy[19].str.upper().str.contains('DASH'))][12].count()
		DivisionBD365 = TotalBDQty and DiasBD365 / TotalBDQty or 0  # a / b
		DivisionBD547 = TotalBDQty and DiasBD547 / TotalBDQty or 0  # a / b

		DiasB365 = dlanewcopy[(dlanewcopy[2]==k)&(dlanewcopy[17].astype(int) < 366)&(dlanewcopy[19].str.upper().str.contains('DASH')==False)][12].count()
		DiasB547 = dlanewcopy[(dlanewcopy[2]==k)&(dlanewcopy[17].astype(int) >365)&(dlanewcopy[19].str.upper().str.contains('DASH')==False)][12].count()
		DivisionB365 = TotalBBQty and DiasBD365 / TotalBBQty or 0  # a / b
		DivisionB547 = TotalBBQty and DiasBD547 / TotalBBQty or 0  # a / b

		
		dlanew.loc[len(dlanew.index)] = ['253',k,'',TotalBDQty,'','',0,'AIR BAG DASH','']

		dlanew.loc[len(dlanew.index)] = ['253',k,'',TotalBQty-TotalBDQty,'','',0,'AIR BAG','']

		Inventarioframe.loc[len(Inventarioframe.index)]= ['253',TotalBBQty,k,"{:.2%}".format(DivisionB365),"{:.2%}".format(DivisionB547)]
		Inventarioframe.loc[len(Inventarioframe.index)]= ['253D',TotalBDQty,k,"{:.2%}".format(DivisionBD365),"{:.2%}".format(DivisionBD547)]


	TotalBQty7=dlanewcopy[dlanewcopy[2]==7][12].sum() 
	TotalBQty8=dlanewcopy[dlanewcopy[2]==8][12].sum() 
	TotalBQty = TotalBQty7 + TotalBQty8
	#Cantidad de Bolsas Dash
	TotalBDQty7=dlanewcopy[(dlanewcopy[19].str.upper().str.contains('DASH'))&(dlanewcopy[2]==7)][12].sum()
	TotalBDQty8=dlanewcopy[(dlanewcopy[19].str.upper().str.contains('DASH'))&(dlanewcopy[2]==8)][12].sum()
	TotalBDQty = TotalBDQty7 + TotalBDQty8

	TotalBBQty= TotalBQty - TotalBDQty 
	TotalBBQty7= TotalBQty7 - TotalBDQty7
	TotalBBQty8= TotalBQty8 - TotalBDQty8

	DiasBD3657 = dlanewcopy[(dlanewcopy[2]==7)&(dlanewcopy[17].astype(int) < 366)&(dlanewcopy[19].str.upper().str.contains('DASH'))][12].count()
	DiasBD3658 = dlanewcopy[(dlanewcopy[2]==8)&(dlanewcopy[17].astype(int) < 366)&(dlanewcopy[19].str.upper().str.contains('DASH'))][12].count()
	DiasBD365 = DiasBD3657 + DiasBD3658

	DiasBD5477 = dlanewcopy[(dlanewcopy[2]==7)&(dlanewcopy[17].astype(int) >365)&(dlanewcopy[19].str.upper().str.contains('DASH'))][12].count()
	DiasBD5478 = dlanewcopy[(dlanewcopy[2]==8)&(dlanewcopy[17].astype(int) >365)&(dlanewcopy[19].str.upper().str.contains('DASH'))][12].count()
	DiasBD547  = DiasBD5477 + DiasBD5478

	DivisionBD365 = TotalBDQty and DiasBD365 / TotalBDQty or 0  # a / b
	DivisionBD547 = TotalBDQty and DiasBD547 / TotalBDQty or 0  # a / b

	DiasB3657 = dlanewcopy[(dlanewcopy[2]==7)&(dlanewcopy[17].astype(int) < 366)&(dlanewcopy[19].str.upper().str.contains('DASH')==False)][12].count()
	DiasB3658 = dlanewcopy[(dlanewcopy[2]==8)&(dlanewcopy[17].astype(int) < 366)&(dlanewcopy[19].str.upper().str.contains('DASH')==False)][12].count()
	DiasB365 = DiasB3657 + DiasB3658

	DiasB5477 = dlanewcopy[(dlanewcopy[2]==7)&(dlanewcopy[17].astype(int) >365)&(dlanewcopy[19].str.upper().str.contains('DASH')==False)][12].count()
	DiasB5478 = dlanewcopy[(dlanewcopy[2]==8)&(dlanewcopy[17].astype(int) >365)&(dlanewcopy[19].str.upper().str.contains('DASH')==False)][12].count()
	DiasB547  = DiasB5477 + DiasB5478 


	DivisionB365 = TotalBBQty and DiasBD365 / TotalBBQty or 0  # a / b
	DivisionB547 = TotalBBQty and DiasBD547 / TotalBBQty or 0  # a / b

		
	dlanew.loc[len(dlanew.index)] = ['253',7,'',TotalBDQty7,'','',0,'AIR BAG DASH','']
	dlanew.loc[len(dlanew.index)] = ['253',8,'',TotalBDQty8,'','',0,'AIR BAG DASH','']

	dlanew.loc[len(dlanew.index)] = ['253',7,'',TotalBQty7-TotalBDQty7,'','',0,'AIR BAG','']
	dlanew.loc[len(dlanew.index)] = ['253',8,'',TotalBQty8-TotalBDQty8,'','',0,'AIR BAG','']

	Inventarioframe.loc[len(Inventarioframe.index)]= ['253',TotalBBQty7,7,"{:.2%}".format(DivisionB365),"{:.2%}".format(DivisionB547)]
	Inventarioframe.loc[len(Inventarioframe.index)]= ['253',TotalBBQty8,8,"{:.2%}".format(DivisionB365),"{:.2%}".format(DivisionB547)]
	Inventarioframe.loc[len(Inventarioframe.index)]= ['253D',TotalBDQty7,7,"{:.2%}".format(DivisionBD365),"{:.2%}".format(DivisionBD547)]
	Inventarioframe.loc[len(Inventarioframe.index)]= ['253D',TotalBDQty8,8,"{:.2%}".format(DivisionBD365),"{:.2%}".format(DivisionBD547)]

	for k in tiendas:
		for j in codigos:
			if j!='253':
				TotalPartes=dlanew[(dlanew[2]==k)&(dlanew['Codigo']==j)][12].sum()
				Dias365 = dlanew[(dlanew[2]==k)&(dlanew[17].astype(int) < 366)&(dlanew['Codigo']==j)][12].count()
				Dias547 = dlanew[(dlanew[2]==k)&(dlanew[17].astype(int) >365)&(dlanew['Codigo']==j)][12].count()
				#print('Tienda: ',k,' Codigo:',j," Cantidad:",TotalPartes)
				#numero1 =
				Division365 = TotalPartes and Dias365 / TotalPartes or 0  # a / b
				Division547 = TotalPartes and Dias547 / TotalPartes or 0  # a / b

				#Tventas=dl0[(dl0[2]==k)&(dl0['Codigo']==j)][17].sum()
				#Tpventas=dl0[(dl0[2]==k)&(dl0['Codigo']==j)][16].sum()

				Inventarioframe.loc[len(Inventarioframe.index)]= [j,TotalPartes,k,"{:.2%}".format(Division365),"{:.2%}".format(Division547)]

	for j in codigos:
		if j!='253':
			TotalPartes7=dlanew[(dlanew[2]==7)&(dlanew['Codigo']==j)][12].sum()
			TotalPartes8=dlanew[(dlanew[2]==8)&(dlanew['Codigo']==j)][12].sum()
			TotalPartes= TotalPartes7 + TotalPartes8
			Dias3657 = dlanew[(dlanew[2]==7)&(dlanew[17].astype(int) < 366)&(dlanew['Codigo']==j)][12].count()
			Dias3658 = dlanew[(dlanew[2]==8)&(dlanew[17].astype(int) < 366)&(dlanew['Codigo']==j)][12].count()
			Dias365 = Dias3657 + Dias3658
			Dias5477 = dlanew[(dlanew[2]==7)&(dlanew[17].astype(int) >365)&(dlanew['Codigo']==j)][12].count()
			Dias5478 = dlanew[(dlanew[2]==8)&(dlanew[17].astype(int) >365)&(dlanew['Codigo']==j)][12].count()
			Dias547 = Dias5477 + Dias5478
			#print('Tienda: ',k,' Codigo:',j," Cantidad:",TotalPartes)
			#numero1 =
			Division365 = TotalPartes and Dias365 / TotalPartes or 0  # a / b
			Division547 = TotalPartes and Dias547 / TotalPartes or 0  # a / b
			Inventarioframe.loc[len(Inventarioframe.index)]= [j,TotalPartes7,7,"{:.2%}".format(Division365),"{:.2%}".format(Division547)]
			Inventarioframe.loc[len(Inventarioframe.index)]= [j,TotalPartes8,8,"{:.2%}".format(Division365),"{:.2%}".format(Division547)]

	########################################################################

	# write to dataframe

	print("Updating Google Sheet Report......")
	#Vdiaria.clear()
	#Vdiaria.to_excel('Vdiaria.xlsx',index=False,header=True)
	#InDB.to_excel('InDB.xlsx',index=False,header=True)
	#Vdiaria.to_excel('Vdiaria.xlsx',index=False,header=True)
	#Vdiaria.to_excel('Vdiaria.xlsx',index=False,header=True)
	#Vdiaria.to_excel('Vdiaria.xlsx',index=False,header=True)
	#Vdiaria.to_excel('Vdiaria.xlsx',index=False,header=True)
	#Vdiaria.set_dataframe(VentasDiaframe)
	Vdiaria.clear(start='A1', end=None, fields='*')
	Vdiaria.set_dataframe(VentasDiaframe,(0,0))
	#set_with_dataframe(worksheet=Vdiaria, dataframe=VentasDiaframe, include_index=False,
	#include_column_header=True, resize=True)
	print("Updating Google Sheet Report......")
	#InDB.clear()
	InDB.clear(start='A1', end=None, fields='*')
	InDB.set_dataframe(Inventarioframe,(0,0))
	#set_with_dataframe(worksheet=InDB, dataframe=Inventarioframe, include_index=False,
	#include_column_header=True, resize=True)
	print("Updating Google Sheet Report......")
	#AcoDiario.clear()
	AcoDiario.clear(start='A1', end=None, fields='*')
	AcoDiario.set_dataframe(AcomodoDiaframe,(0,0))
	#set_with_dataframe(worksheet=AcoDiario, dataframe=AcomodoDiaframe, include_index=False,
	#include_column_header=True, resize=True)
	print("Updating Google Sheet Report......")
	#V1y.clear()
	V1y.clear(start='A1', end=None, fields='*')
	V1y.set_dataframe(dl8,(0,0))
	#set_with_dataframe(V1y, dl8)
	#set_with_dataframe(worksheet=V1y, dataframe=dl8, include_index=False,
	#include_column_header=True, resize=True)  
	print("Updating Google Sheet Report......")
	V3m.clear(start='A1', end=None, fields='*')
	V3m.set_dataframe(dl,(0,0))
	#V3m.clear()
	#dl.to_excel('dl.xlsx',index=False,header=True)
	#valores_a_eliminar = ["Stores: ", "Part Types: ", "Sort By: "]
	# Selecciona las filas que no tienen los valores a eliminar en la 5ta columna
	#dl = dl.loc[~dl.iloc[:, 4].isin(valores_a_eliminar)]
	# El índice del DataFrame puede quedar desordenado, puedes resetearlo si lo deseas
	#dl = dl.reset_index(drop=True)
	#dl = dl.dropna(subset=[4])
	# El índice del DataFrame puede quedar desordenado, puedes resetearlo si lo deseas
	#dl= dl.reset_index(drop=True)
	#column_types = dl.dtypes
	# Filtrar las columnas que son de tipo int64
	#int64_columns = column_types[column_types == 'int64']

	# Imprimir las columnas de tipo int64
	#print(int64_columns)
	#dl.to_excel('dltype.xlsx',index=False,header=True)
	#set_with_dataframe(V3m, dl)
	#set_with_dataframe(worksheet=V3m, dataframe=dl, include_index=False,
	#include_column_header=True, resize=True)
	#print("Updating Google Sheet Report......")
	#V1y.clear()
	#set_with_dataframe(worksheet=V1y, dataframe=dl8, include_index=False,
	#include_column_header=True, resize=True)  
	#print("Updating Google Sheet Report......")  
	#InDB.clear()
	#set_with_dataframe(worksheet=InDB, dataframe=Inventarioframe, include_index=False,
	#include_column_header=True, resize=True)
	#print("Updating Google Sheet Report......")
	#AcoDiario.clear()
	#set_with_dataframe(worksheet=AcoDiario, dataframe=AcomodoDiaframe, include_index=False,
	#include_column_header=True, resize=True)
	#print("Updating Google Sheet Report......")
	#Vdiaria.clear()
	#set_with_dataframe(worksheet=Vdiaria, dataframe=VentasDiaframe, include_index=False,
	#include_column_header=True, resize=True)


	#print(list(dlanew.columns.values))
	#creo el archivo excel limpio , solo con los datos 
	dl8.to_excel('ventas_partes_year.xlsx', index=False,header=True)
	dl.to_excel('ventas_partes_3m.xlsx', index=False,header=True)
	#
	dlanew.to_excel('inventorylocationwarehouses_update.xlsx', index=False,header=True)
	Inventarioframe.to_excel('inventorylocationwarehouses_google.xlsx', index=False,header=True)
	#######dlanewcopy.to_excel('inventorylocationwarehouses_codes.xlsx', index=False,header=True)

if __name__ == '__main__':
	conf = parse_args()
	Principal(conf.Directorio_de_trabajo)