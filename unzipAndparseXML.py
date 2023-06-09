# -*- coding: utf-8 -*-
"""
Created on Tue Apr 25 16:24:05 2023

@author: LuPerez
"""
from __future__ import unicode_literals
from zipfile import ZipFile
import os
import shutil


import io

import pandas as pd

import xml.etree.ElementTree as ET






def unZip(dir_origen, path_destino):
   contenido=getListaArchivos(dir_origen)
   for fichero in contenido:
        path_zip=getOrigenArchivos()+"/"+fichero
        with ZipFile(path_zip, 'r') as zip:
            zip.extractall(path_destino)
            print("File is unzipped in "+path_destino+" folder")
        
def getDirRepositorio():
    return "D:\\Lucia\\buro\\buroZip\\ZipRepositorio"         

def getDestino():
    return "D:\\Lucia\\buro\\buroZip\\Descomprimido"

def getOrigenArchivos():
    return 'D:\\Lucia\\buro\\buroZip\\Zip'

def getListaArchivos(dir_carpeta):
    contenido=os.listdir(dir_carpeta)
    return contenido
    

def getStringArchivoXML(dirArchivoXLS):
    file1=io.open(dirArchivoXLS,"r",encoding="UTF-8")
    info=file1.readlines()
    return info[0]

def crearDataFrame():
    df = pd.DataFrame()
    df['Tipo ID']=None
    df['Numero ID']=None
    df['Extencion']=None
    df['Nombre']=None
    df['Entidad']=None
    df['Departamento']=None
    df['fecha']=None
    df['Calificacion']=None
    df['Saldo']=None
    df['Estado']=None
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
                    print(mi_lista)
                    df.loc[len(df)] = mi_lista

    df=df.drop_duplicates()
    df['nombre_buro']="infocred"
    print(df)
    return df

def eliminarArchivos(dir_eliminar):
    for file in os.scandir(dir_eliminar):
        print("este archivo"+file.path)
        os.remove(file.path)
        
def copiarArchivos(dir_origen, dir_destino):
    contenido=getListaArchivos(dir_origen)
    for archivo in contenido:
        shutil.copy(dir_origen+"\\"+archivo, 
                    dir_destino+"\\"+archivo)


def rutaOrigenInfocenter():
    return "D:\\Lucia\\buro\\buroZip\\infocenter"

def rutaRepoInfocenter():
    return "D:\\Lucia\\buro\\buroZip\\InfocenterRepositorio"




unZip(getOrigenArchivos(), getDestino())

lista_xls=getListaArchivos(getDestino())
for archivo in lista_xls:
    print("********* "+archivo)
    df=xmlParcer(getDestino()+"\\"+ archivo)
    print(df)
 
copiarArchivos(getOrigenArchivos(), getDirRepositorio())
eliminarArchivos(getOrigenArchivos())
eliminarArchivos(getDestino())

## CAGAR INFOCENTER
ruta=rutaOrigenInfocenter()
listaInfo=getListaArchivos(ruta)
for file in listaInfo:
    ruta=ruta+"\\"+file
    df = pd.read_excel(ruta)
    df['nombre_buro']="infocenter"
    
    print(df)

copiarArchivos(rutaOrigenInfocenter(), rutaRepoInfocenter())
eliminarArchivos(rutaOrigenInfocenter())
