# -*- coding: utf-8 -*-
"""
Created on Tue Apr 25 16:24:05 2023

@author: LuPerez
"""
from __future__ import unicode_literals
from zipfile import ZipFile
import os


from xlwt import Workbook
import io

import pandas as pd


import xml.etree.ElementTree as ET

from xml.dom import minidom





def unZip(dir_origen, path_destino):
   contenido=getListaArchivos(dir_origen)
   for fichero in contenido:
        path_zip=getOrigenArchivos()+"/"+fichero
        with ZipFile(path_zip, 'r') as zip:
            zip.extractall(path_destino)
            print("File is unzipped in "+path_destino+" folder")
            

def getDestino():
    return "D:/Lucia/buro/buroZip/Descomprimido"

def getOrigenArchivos():
    return 'D:/Lucia/buro/buroZip'

def getListaArchivos(dir_carpeta):
    contenido=os.listdir(dir_carpeta)
    return contenido
    

def getStringArchivoXML(dirArchivoXLS):
    file1=io.open(dirArchivoXLS,"r",encoding="UTF-8")
    info=file1.readlines()
    return info[0]

def crearDataFrame():
    df = pd.DataFrame()
    df['Ti']=None
    df['Nu']=None
    df['Ex']=None
    df['No']=None
    df['En']=None
    df['De']=None
    df['fe']=None
    df['Ca']=None
    df['Sa']=None
    df['Es']=None
    return df


def xmlParcer(dirArchivoXLS):
    doc =getStringArchivoXML(dirArchivoXLS)
    tree = ET.ElementTree(ET.fromstring(doc))
    root=tree.getroot()
    for child in root:
        print(child.tag, child.attrib)
    mi_lista = []
    
    df = crearDataFrame()
    
    for element in root:
        for subelement in element:
            for subelements in subelement:
                contador=100
                for items in subelements:
                   
                    for data in items:
                        dato=data.text
                        if dato=="CI":
                            contador=0
                            if len(mi_lista)>0:
                                mi_lista.clear()
                        if contador <10:
                            mi_lista.append(dato)
                        contador=contador+1
                         
                if len(mi_lista)>0:
                    df.loc[len(df)] = mi_lista

    df=df.drop_duplicates()
    print(df)
    return df

def eliminarArcivos(dir):
    for file in os.scandir(dir):
        os.remove(file.path)



unZip(getOrigenArchivos(), getDestino())

lista_xls=getListaArchivos(getDestino())
for archivo in lista_xls:
    print("********* "+archivo)
    df=xmlParcer(getDestino()+"/"+ archivo)
    print(df)

    
eliminarArcivos(getOrigenArchivos())
    



