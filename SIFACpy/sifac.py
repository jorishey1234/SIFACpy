#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 20 15:36:56 2026

@author: joris
"""
import numpy as np
import pandas as pd
import glob
import re
import os
import sys

try: 
	os.remove('sifac_log.txt')
except:
	print('no previous log to remove')

sys.stdout = open('sifac_log.txt', 'w')


params=pd.read_excel('parametres_sifac.xls')
file = params.iloc[0,1]
file_existing = params.iloc[1,1]
sheet_existing = params.iloc[2,1]
outputfile= params.iloc[3,1]
column_name=params.iloc[4,1]
dic_libelles=params.iloc[5,1]
dic_names=params.iloc[6,1]
sifac_libelle=params.iloc[7,1]
sifac_num=params.iloc[8,1]
#gestion_num=params.iloc[9,1]
# file_existing='AO_2025-DPT EAU_O1-O4-O5-COMMUN_Etat des Consommations SE.xlsm'
# file='EXTRACTION_SIFAC.xlsx'
# outputfile='EXTRACTION_SIFAC_proc.xlsx'

F=pd.read_excel(file)

try:
	Fexisting=pd.read_excel(file_existing,sheet_name=sheet_existing,skiprows=3)
except:
	print('Fichier de Gestion existant non trouvé !')
	file_existing=0

def highlight_col(x,color):
	return np.where(x=='', f"background color: {color};", None)

def make_search(Dic):
	#create list of re
	Codes=[]
	Expr=[]
	for i in range(len(Dic)):
		expr=Dic.iloc[i,1].split(',')
		print(expr)
		for e in expr:
			if len(e)>0:
				Codes.append(Dic.iloc[i,0])
				Expr.append(re.compile(e))
	return Codes,Expr

Dic=pd.read_excel(dic_libelles)
Codes,Expr = make_search(Dic)

Dic=pd.read_excel(dic_names)
Codes_names,Expr_names = make_search(Dic)


highlight=[]
		
F.insert(0, column_name, '')
#O4=re.compile('O4')
for i,f in enumerate(F[sifac_libelle]):
#	print(f)
#	m = re.search(r'O4', str(f))
	Res=[]

	# First look in the existing file for a manual change
	if file_existing:
		if len(Res)==0:
			print('looking in '+ file_existing +'/'+sheet_existing +' for the N° de commande',F.loc[i,sifac_num])
			s=Fexisting.iloc[:,0][Fexisting.iloc[:,3]==F.loc[i,sifac_num]]
			if len(s)>0:
				print('Found : ',s.iloc[0])
				Res.append(s.iloc[0])

	# If nothing has been found, look for the Libelle col in the sifac extraction file
	if len(Res)==0:
		for j in range(len(Codes)):
			m = Expr[j].search(str(f))
			if m:
				print(Codes[j],' <<< ',f)
				Res.append(Codes[j])
	# If len(Res)>1 possibly raise a warning !
	
	# If nothing has been found, look for the Name col  in the sifac extraction file
	if len(Res)==0:
		name=F['Nom du tiers'][i]
#		print(name)
		for j in range(len(Codes_names)):
			m = Expr_names[j].search(str(name))
			if m:
				print(Codes_names[j],' <<< ',name)
				Res.append(Codes[j])
	
	if len(Res)>0:
		F.loc[i,column_name]=Res[0]
	# If nothing has been found, put a warning
	else:
			print('Nothing found higlight cell')
# 			highlight.append(1)
# 		else:
# 			highlight.append(0) 
		
F.style.apply(highlight_col,color="yellow")
F.to_excel(outputfile, index=False)

sys.stdout.close()
