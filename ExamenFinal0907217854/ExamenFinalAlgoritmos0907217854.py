import openpyxl
import sys
from openpyxl.styles import Border,Side

libro = openpyxl.load_workbook("vehículos.xlsx")
hoja = libro['Listado']


hoja['A1'].value = "Código"
hoja['B1'].value = "Marca"
hoja['C1'].value = "Modelo"
hoja['D1'].value = "Precio"
hoja['E1'].value = "Kilometraje"
hoja['F1'].value = "CantidadFotos"

def mantenimiento_exel():
    print("Los datos van separados por el simbolo |")
    opcion_elegida = sys.argv[1]
    if opcion_elegida == "Crear vehículos":
        Crear_vehículos = sys.argv[2].split("|")
        
    if opcion_elegida == "Editar vehículos":

    if opcion_elegida == "Eliminar vehículos":
    
    if opcion_elegida == "Listar vehículos"




    
datos_entrada_ejemplo = [
   {
      "Código":"CITY01",
      "Marca":"HONDA",
      "Modelo":"2020",
      "Precio":"80000",
      "Kilometraje":"600",
      "CantidadFotos":"0"
   },
   {
      "Código":"CIVIC01",
      "Marca":"HONDA",
      "Modelo":"2021",
      "Precio":"90000",
      "Kilometraje":"0",
      "CantidadFotos":"0"
   },
   {
      "Código":"PILOT01",
      "Marca":"HONDA",
      "Modelo":"2021",
      "Precio":"40000",
      "Kilometraje":"1300",
      "CantidadFotos":"0"
   },
   {
      "Código":"BT50",
      "Marca":"MAZDA",
      "Modelo":"2021",
      "Precio":"50000",
      "Kilometraje":"600",
      "CantidadFotos":"0"
   },
   {
      "Código":"BALENO1",
      "Marca":"SUZUKI",
      "Modelo":"2021",
      "Precio":"60000",
      "Kilometraje":"2000",
      "CantidadFotos":"0"
   },
   {
      "Código":"XL71",
      "Marca":"SUZUKI",
      "Modelo":"2021",
      "Precio":"70000",
      "Kilometraje":"1500",
      "CantidadFotos":"0"
   }
]

proxima_fila = hoja.max_row + 1

top=Side(border_style='thick',color="A52A2A")
left=Side(border_style='thick', color="0000FF")

border=Border(top=top,left=left)

for venta in datos_entrada_ejemplo:
    hoja[f'A{proxima_fila}'].value = venta["Código"]
    hoja[f'B{proxima_fila}'].value = venta["Marca"]
    hoja[f'C{proxima_fila}'].value = venta["Modelo"]
    hoja[f'D{proxima_fila}'].value = venta["Precio"]
    hoja[f'E{proxima_fila}'].value = venta["Kilometraje"]
    hoja[f'F{proxima_fila}'].value = venta["CantidadFotos"]

    hoja[f'A{proxima_fila}'].border = border
    hoja[f'B{proxima_fila}'].border = border
    proxima_fila +=1
    
 libro.save("vehículos.xlsx")