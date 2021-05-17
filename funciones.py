
from os import path
import openpyxl as xl
from io import open
import pathlib
#Creo una funcion para obtener un archivo
def getFile( name, permissions = "a+"):
    rutaAbsoluta = str(pathlib.Path().absolute())
    name = "/" + name
    archivo = open( rutaAbsoluta + name, permissions )
    return archivo
#Creo otra funcion para obtener la informacion de un archivo json
def getJSON( archivo ):
    return eval( archivo.read() )
#Retorno los metodos para obtener y guardar el documento respectivamente
def getDocData( name ):
    doc = xl.load_workbook( name )
    return {
        'get': doc,
        'save': lambda: doc.save( name )
    }
#FUNCIÓN PARA OBTENER LOS VALORES DEL INVENTARIO
def getInventario( doc ):
    sheet = doc['Inventario']
    cont = 2
    registro = {}
    while True:
        referencia = sheet[f'A{cont}'].value
        if not referencia:
            break
        
        #OBTENEMOS LAS VENTAS Y LA CANTIDAD
        fila = {
            "VENTAS": sheet[f'C{cont}'].value,
            "DESCRIPCION": sheet[f'B{cont}'].value,
            "CANTIDAD": sheet[f'D{cont}'].value,
            "VALOR VENTA": sheet[f'F{cont}'].value,
        }
        registro[referencia] = fila
        cont += 1
    return registro

#ACTUALIZAMOS EL EXCEL DEL INVENTARIO SEGÚN LOS NUEVOS DATOS DE VENTAS
def uploadInventario( doc, newInv ):
    sheet = doc['Inventario']
    cont = 2
    while True:
        referencia = sheet[f'A{cont}'].value
        if not referencia:
            break
        newProducto = newInv[referencia]
        sheet[f'C{cont}'] = newProducto["VENTAS"]
        sheet[f'D{cont}'] = newProducto["CANTIDAD"]
        cont += 1

#AÑADIMOS A BANCOLOMBIA TODAS LAS TRANSACCIONES DE VENTAS
def uploadBancolombia( doc, transacciones ):
    sheet = doc['Bancolombia']
    puntoActual = 1
    while True:
        if not sheet[f'A{puntoActual}'].value:
            break
        puntoActual += 1
    for transaccion in transacciones:
        fecha = transaccion["FECHA"]
        sheet[f'A{puntoActual}'] = f'{fecha["DIA"]}/{fecha["MES"]}/{fecha["AÑO"]}'
        sheet[f'B{puntoActual}'] = f'Venta de {transaccion["CANTIDAD"]} unidades de {transaccion["DESCRIPCION"]}'
        sheet[f'C{puntoActual}'] = f'${transaccion["VALOR PRODUCTO"] * transaccion["CANTIDAD"]}'
        puntoActual += 1

def barrierVentas( doc ):
    sheet = doc['Ventas']
    cont = 2
    bancolombia = []
    registro = {}
    #OBTENEMOS EL INVENTARIO PARA EDITAR LOS VALORES DE LOS PRODUCTOS
    inventario = getInventario( doc )
    while True:
        id_ = sheet[f'A{cont}'].value
        checked = sheet[f'L{cont}']
        if not id_:
            break
        if checked.value == "SI":
            cont += 1
            continue
        sheet[f'L{cont}'] = "SI"
        fila = {
            "REFERENCIA": sheet[f'B{cont}'].value,
            "CANTIDAD": sheet[f'E{cont}'].value,
            "VALOR PRODUCTO": sheet[f'F{cont}'].value,
            "FORMA DE PAGO": sheet[f'J{cont}'].value,
            "FECHA": {
                "DIA": sheet[f'G{cont}'].value,
                "MES": sheet[f'H{cont}'].value,
                "AÑO": sheet[f'I{cont}'].value
            }
        }
        producto = inventario[ fila['REFERENCIA'] ]
        
        #HACEMOS LAS EDICIONES RESPECTIVAS AL PRODUCTO (CON CONDICIONALES EN CASO DE ESTAR VACÍOS)
        if producto["VENTAS"]:
            producto["VENTAS"] += fila["CANTIDAD"]
        else:
            producto["VENTAS"] = fila["CANTIDAD"]

        if producto["CANTIDAD"]:
            producto["CANTIDAD"] -= fila["CANTIDAD"]
        else:
            producto["CANTIDAD"] = -fila["CANTIDAD"]
            
        if fila["FORMA DE PAGO"].upper() == "BANCOLOMBIA":
            fila["DESCRIPCION"] = producto["DESCRIPCION"]
            bancolombia.append( fila )
            cont += 1
            continue
        registro[id_] = fila
        cont += 1
    uploadBancolombia( doc, bancolombia )
    return  registro

def uploadGastosFijos( sheet ):
    json = getJSON( getFile('datos.json', 'r') )
    cont = 3
    total = 0
    for gasto in json['gastos_fijos']:
        sheet[f'A{cont}'] = gasto
        gasto = json['gastos_fijos'][gasto]
        sheet[f'B{cont}'] = gasto
        total += gasto
        cont += 1
    return total

def uploadGastosAdicionales( sheet ):
    cont = 3
    total = 0
    while True:
        gasto = sheet[f'C{cont}'].value
        valor = sheet[f'D{cont}'].value
        if not gasto:
            break
        total += valor
        cont += 1
    return total
    
def uploadGastos( doc ):
    sheet = doc['Gastos']
    totalGastos = uploadGastosFijos( sheet )
    totalGastos += uploadGastosAdicionales( sheet )
    sheet['E3'] = totalGastos


def init( docName ):
    data = getDocData( docName )
    doc = data['get']
    saveDoc = data['save']
    barrierVentas( doc )
    uploadGastos( doc )
    saveDoc()





"""
Barrer Ventas ✅
    Guardar precio ✅
    Guardar los que hayan sido de bancolombia ✅
"""