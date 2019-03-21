from netmiko import Netmiko
import xlrd
import os
import time
import getpass
import paramiko
import re
from pandas import DataFrame

paramiko.util.log_to_file("Paramiko_log.log")

def Generar_Script(Nombre, Contenido):
    file = open(Nombre + ".txt", 'a')
    file.write(Contenido)

def Abrir_Excel(path, hoja):
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_name(hoja)
    return (wb, sheet)

# Retorna los valores de una columna en un vector
def Obtener_Columna(Hoja, columna):
    vector = []
    for i in range(1, Hoja.nrows):
        vector.append(str(Hoja.cell_value(i, columna)))
    return vector

def Reemplazo_lista(lista, str1, str2):
    nueva_lista = [a.replace(str1, str2) for a in lista]
    return nueva_lista


def connect_huawei(ip,password):
    huawei_device = {
        "host": ip,
        "username": "estebauribe.huawei",
        "password": password,
        "device_type": "huawei",
    }
    net_connect = Netmiko(**huawei_device)
    return net_connect

def get_description(input):
    try:
        return input.split('description')[1]
    except:
        return 'N/A'


def process_OSPF (input):
    text2 = input.split('\n')
    Interfaces = []
    Costos = []
    for line in text2:
        substring = line.split()
        if 'P-2-P' in substring:
            if 'GE' in substring[0]:
                Interfaces.append(substring[0].replace('GE','Gi'))
            else:
                Interfaces.append(substring[0])
            Costos.append(substring[4])
    Interfaces = [str(item) for item in Interfaces]
    Costos = [str(item) for item in Costos]
    return Interfaces, Costos

File = 'C:\Users\e80047212\Documents\Entel_Ingenieria\On-site-Tuning\Control-Homologacion-IP-SWAP-v1_0.xlsx'
Sheet = 'Actividad 16'
(Archivo, Hoja) = Abrir_Excel(File, Sheet)
# Retorna los valores de la columna 0 en un vector
Hostnames = Obtener_Columna(Hoja, 0)
IPs = Obtener_Columna(Hoja, 4)
# Nombre de usuario
Username = 'estebauribe.huawei'
password = getpass.getpass()
i = 0
comando = 'dis ospf interface'
Interfaces = []
Costos = []
Equipos = []
Description = []

for ip in IPs:
    if '.' in ip:
        print(Hostnames[i])
        try:
            net_connect = connect_huawei(ip,password)
        except:
            Output = "Verificar conexion"
            i += 1
            print("No fue posible conectarse al equipo")
            continue
        Generar_Script('Revision_Actividad_14', comando + '\n')
        Output = net_connect.send_command(comando)
        L1_If, L2_Costs = process_OSPF(Output)
        for item in L1_If:
            Interfaces.append(item)
            try:
                Output = net_connect.send_command('dis cu int ' + item + ' | inc descr ')
            except:
                print("No fue posible conectarse al equipo")
                Output = 'N/A'
            Description.append(str(get_description(Output)))
            Equipos.append(Hostnames[i])
        for item in L2_Costs:
            Costos.append(item)
        i += 1
        #print len(Description)
        #print len(Equipos)
        #print len(Interfaces)
        #print len(Costos)
    else:
        print ("No es una IP")
        break

#print Interfaces
#print Costos
#print Equipos
#print (len(Interfaces))
#print (len(Costos))
#print (len(Equipos))
df = DataFrame({'Hostname':Equipos,'Costos':Costos,'Interfaces':Interfaces, 'Descripcion':Description})
print df
try:
    df.to_excel('test.xlsx',sheet_name='sheet1', index=False)
except IOError:
    print ("No se puede escribir sobre el archivo, tal vez esta abierto")