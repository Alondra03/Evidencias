import datetime
import time
import sys
import os
import openpyxl
import csv

def generarID(diccionario):
    lista=[]
    try:
        for key in diccionario:
            lista.append(key)
        return max(lista)+1
    except:
        return 1

def validar_ID(estructura,texto):
  registros = len(estructura)
  while True:
    ID = input(texto)
    try:
        ID = int(ID)
        if ID <= registros and ID > 0:
          return ID
        else:
          print(f"ERROR folio no existente.")
          continue
    except Exception:
        print(f"ERROR folio no valido.")
        print("Ingrese de nuevo...")
        
def Editar_Nombre_Evento ():
    try:
      ID_Evento= validar_ID(eventos,"Ingrese folio de reservacion: ")
      if eventos.get(ID_Evento) is not None:
          while True:
            editar_nombre = input('Nuevo nombre para el evento: ')
            if editar_nombre == "":
              print('Nombre no valido!')
            else:
              break
          eventos [ID_Evento][2]= editar_nombre
          print('\n Se edito correctamente el nombre del evento!')
    except:
      print('*** ERROR! El numero ingresado no existe ***')

def Consultar_Reservaciones ():
    while True:
      try:
        os.system("cls")
        Fecha_reservacion= input('Fecha de la reservacion que desea buscar (dd/mm/aaaa): ')
        fecha_procesada = datetime.datetime.strptime(Fecha_reservacion, "%d/%m/%Y").date() 
        break   
      except:
            print(f"ERROR fecha no valida!")
            print("Ingrese de nuevo...")
    print("")
    print("*"*85)
    print(f'**{"REPORTE DE RESERVACIONES PARA EL DIA "+ Fecha_reservacion:^81}**')
    print("*"*85)  
    print(f'{"SALA":^7}{"CLIENTE":^33}{"EVENTO":^33}{"TURNO":^12}')
    print('*'*85)
    for id_fecha in eventos:
        if Fecha_reservacion == eventos[id_fecha][-1]:
          print(f'{eventos[id_fecha][0]:^7}{eventos[id_fecha][1]:^33}{eventos[id_fecha][2]:^33}{eventos[id_fecha][3]:^12}')
    print(f'{"FIN DEL REPORTE":*^85}')
      
# ESTRUCTURAS
turnos= {1:"Matutino", 2:"Vespertino", 3:"Nocturno"}
Salas = {}
clientes = {}
eventos = {} 
ocupado = {}

# LLENADO DE INFO
try:
  with open("reservaciones.csv","r",newline="") as archivo:
    lector = csv.reader(archivo)
    next(lector)
    for idReservacion,idSala,cliente,evento,id_turno,turno,fecha in lector:
      eventos[int(idReservacion)]=[int(idSala),cliente,evento,id_turno,turno,fecha]

  with open("salas.csv","r",newline="") as archivo:
    lector = csv.reader(archivo)
    next(lector)
    for idSala,nombre,capacidad in lector:
      Salas[int(idSala)]=[nombre,int(capacidad)]

  with open("ocupado.csv","r",newline="") as archivo:
    lector = csv.reader(archivo)
    next(lector)
    for idReservacion, idSala, idTurno, fecha in lector:
      ocupado[int(idReservacion)]=[int(idSala),int(idTurno),fecha]

  with open("clientes.csv","r",newline="") as archivo:
    lector = csv.reader(archivo)
    next(lector)
    for idCliente,nombre in lector:
      clientes[idCliente]=[nombre]

except:
  print("No se encontro informacion anteriormente registrada")
  input("Presione cualquier tecla para continuar ...")
  
#----MENU PRINCIPAL----
while True:
    os.system("cls")
    print("-"*50)
    print("MENÚ PRINCIPAL")
    print("\t [A] Reservaciones.")
    print("\t [B] Reportes.")
    print("\t [C] Registrar Nueva Sala.")
    print("\t [D] Registrar Nuevo Cliente.") 
    print("\t [X] Salir.")
    op_principal = input("OPCION: ")
    os.system("cls")
    if (op_principal == ""):
          print("No se debe omitir.")
    # OPCION A. RESERVACIONES
    if (op_principal.upper() in "A"):
        # submenú  
        while True:
          print("-"*60)
          print("Menu de Reservaciones")
          print("\t [A] Registrar Nueva Reservación.")
          print("\t [B] Modificar Nombre de una Reservación.")
          print("\t [C] Consultar Disponibilidad de salas para una fecha.")
          print("\t [X] Regresar al menú principal. ")
          op_reserva = input("OPCION: ")
          os.system("cls")
          if (op_reserva == ""):
              print("No se debe omitir.")
          # A. Registrar Nueva Reservación
          elif (op_reserva.upper() in "A"):
              ID_Reservacion = generarID(eventos)
              print("---------------  RESERVACION DE SALA  ----------------")
              if not bool(clientes):
                print("Lo sentimos, primero debe registrarse.")
                input("\nPresione cualquier tecla para continuar...")
                continue
              elif not bool(Salas):
                print("No hay salas registradas")
                input("\nPresione cualquier tecla para continuar...")
                continue

              # Validacion Cliente
              print("CLIENTES")
              for id,nombre in clientes.items():
                print(id,'.-',nombre[0])
                print("-"*54)
              ID_Cliente = validar_ID(clientes,"Ingrese folio de cliente: ")
              os.system("cls")
              print("---------------  RESERVACION DE SALA  ----------------")

              # Validacion Fecha
              while True:
                fecha_capturada = input("\nFecha del evento dd/mm/aaaa: ")
                try:
                  fecha_procesada = datetime.datetime.strptime(fecha_capturada,"%d/%m/%Y").date()
                  fecha_actual = datetime.date.today()
                  fecha_invalida = fecha_actual + datetime.timedelta(days=+1)
                  if fecha_procesada <=  fecha_invalida:
                    print(f"Fecha Rechazada.Se necesita al menos 2 días de anticipación.")
                    continue
                  else:
                    os.system("cls")
                    print("---------------  RESERVACION DE SALA  ----------------")
                    print("SALAS")
                    
                    # Imprime salas registradas
                    for id in Salas:
                      print(id,'.-',Salas[id][0])
                    print("-"*54)
                    
                        # Validacion Sala
                    idSala = validar_ID(Salas,"Ingrese folio de sala: ")
                    os.system("cls")
                    print("---------------  RESERVACION DE SALA  ----------------")
                    print("TURNOS")
                    
                    # Imprime turno
                    for turno in turnos:
                        print(turno,".-",turnos[turno])
                    print("-"*54)
                    
                    # Validacion ID        
                    ID_Turno = validar_ID(turnos,"Ingrese folio de turno: ")
                    os.system("cls")
                    print("---------------  RESERVACION DE SALA  ----------------")
                    while True:
                        Nombre_Reservacion= input('\nNombre para la reservacion: ')
                        if Nombre_Reservacion == "":
                            print("No se puede omitir.")
                            continue
                        else:
                            break

                    # Guardar Registros
                    if ID_Reservacion == 1:
                        eventos[ID_Reservacion] = [idSala,clientes[ID_Cliente][0],Nombre_Reservacion,ID_Turno,turnos[ID_Turno], fecha_capturada]
                        ocupado[ID_Reservacion] = [idSala, (ID_Turno), fecha_capturada] # Nos permitirá comprobar que no hayan eventos con misma fecha, sala y turno
                        print("\nLa reservacion se ha guardado exitosamente!")
                    else:
                        # Validar que NO se empalme con otro evento registrado
                        nueva_Reservacion = ([idSala,(ID_Turno),fecha_capturada])
                        for registro in ocupado.values():
                          if registro == nueva_Reservacion:
                            print("Lo sentimos, sala OCUPADA!")
                            break
                        else:
                          eventos[ID_Reservacion] = [idSala, clientes[ID_Cliente][0], Nombre_Reservacion, ID_Turno, turnos[ID_Turno], fecha_capturada]
                          print(type(fecha_capturada))
                    break
                except ValueError:
                    print(f"ERROR fecha no valida")
                    print("Ingrese de nuevo...")

          # B. Editar Nombre de una Reservación
          elif (op_reserva.upper() in "B"):
            print("------------  EDITAR NOMBRE DE SALA  -------------")
            if not bool(eventos):
              print("NO hay eventos registrados.")
              continue
            else:
              for id in eventos:
                  print(id,'.-',eventos[id][2])
              print("_"*50)
              print('--> Seleccione el numero del evento: ')
              Editar_Nombre_Evento()

         # C. Consultar Disponibilidad de Salas por fecha.
          elif (op_reserva.upper() in "C"):
              print("------ CONSULTAR DISPONIBILIDAD DE SALA  -------")
              print(" ")
              while True:
                fecha_ingresada = input("Ingrese la fecha que desea consultar dd/mm/aaaa: ")
                lista_encontrados = []
                eventos_posibles = []
                try:
                  fecha_buscada=datetime.datetime.strptime(fecha_ingresada,"%d/%m/%Y").date()
                  for valor in eventos.items():
                    sala, turno, fecha = (valor[1][0], valor[1][3], valor[1][5])
                    if fecha == fecha_ingresada:
                      lista_encontrados.append((sala, turno))
                    eventos_encontrados = set(lista_encontrados)
                  for sala in Salas.keys():
                    for turno in turnos.keys():
                      eventos_posibles.append((sala, turno))
                      combinaciones_eventos_posibles = set(eventos_posibles)
                      salas_turnos_disponibles = sorted(list(combinaciones_eventos_posibles - eventos_encontrados))
                  os.system("cls")
                  print(f'** Salas disponibles para renta el {fecha_ingresada} **\n')
                  print("Sala\t\tTurno\n")
                  for sala, turno in salas_turnos_disponibles:
                    print(f"{sala}. {Salas[sala][0]}\t\t{turnos[turno]}")
                  break
                except:
                  print(f"ERROR fecha no valida!")
                  print("Ingrese de nuevo...")
          # X. Salida del submenú.
          elif (op_reserva.upper() in "X"):
              print("Has salido del Menú Reservaciones.")
              break
          else:
              print("OPCION NO VALIDA. Debes ingresar la letra que corresponda a la acción deseada.")
          input("\nPresione cualquier tecla para continuar...")

          os.system("cls")
    # OPCION B. REPORTES
    elif (op_principal.upper() in "B"):
        # submenú
        while True:
            print("-"*50)
            print("MENU REPORTES")
            print("\t [A] Reporte en Pantalla")
            print("\t [B] Reporte en Excel.")
            print("\t [X] Regresar al menú principal. ")
            op_reporte = input("OPCION: ")
            os.system("cls")
            if (op_reporte == ""):
              print("No se debe omitir.")
            # A. Reporte en Pantalla
            if (op_reporte.upper() in "A"):
                if not bool(eventos):
                    print("NO hay eventos registrados.")
                    continue
                print("*"*85)
                print(f'**{"REPORTE DE RESERVACIONES":^81}**')
                print("*"*85)  
                print(f'{"SALA":^7}{"CLIENTE":^28}{"EVENTO":^28}{"TURNO":^12}{"FECHA":^10}')
                print('*'*85)
                for id_fecha in eventos:
                    print(f'{eventos[id_fecha][0]:^7}{eventos[id_fecha][1]:^28}{eventos[id_fecha][2]:^28}{eventos[id_fecha][4]:^12}{eventos[id_fecha][5]:^10}')
                print(f'{"FIN DEL REPORTE":*^85}')
                Consultar_Reservaciones()
                
                
                    
            # B. Reporte en Excel
            elif (op_reporte.upper() in "B"):
                print("Reporte en Excel")
                print("")
                if not bool(eventos):
                    print("NO hay eventos registrados.")
                    continue
                libro = openpyxl.Workbook()
                #libro.iso_dates = True #Acepte fechas              
                hoja = libro["Sheet"] 
                hoja.title = "Reporte"
                hoja.append(('Id', 'Cliente', 'Sala', 'Turno','Fecha')) 
                for id_fecha in eventos:
                  lista = [eventos[id_fecha][0],eventos[id_fecha][1],eventos[id_fecha][2],eventos[id_fecha][4],eventos[id_fecha][5]]                
                  hoja.append(lista)
                libro.save("Reporte.xlsx")
            
            # X. Salida del submenú.
            elif (op_reporte.upper() in "X"):
                  print("Has salido del Menú Reportes.")
                  break
            else:
                print("OPCION NO VALIDA. Debes ingresar la letra que corresponda a la acción deseada.")
            input("\nPresione cualquier tecla para continuar...")
            os.system("cls")          
            
    # OPCION C. RESERVAR NUEVA SALA
    elif (op_principal.upper() in "C"):
        print("---------------  REGISTRAR SALA  -----------------")
        ID_Sala = generarID(Salas)     
        # Validaciones
        while True:
            Nombre = (input("Nombre para la sala: "))
            if Nombre == "":
                print("No se puede omitir.")
                continue
            else:
                break  
        while True:
          try:
            Cupo=int(input("Capacidad de la sala: "))
            if Cupo >0:
                break
            else:
                print("Ocurrió un problema no puede ingresar numeros menores a 0")
          except Exception:
            print(f"Ocurrió un problema numero no valido")
            print("Ingrese de nuevo...")
        Salas[ID_Sala]=(Nombre, Cupo)
        print("\nLa sala ha sido guardada exitosamente!")
        
    # OPCION D. REGISTRAR NUEVO CLIENTE
    elif (op_principal.upper() in "D"):
        print("--------------  REGISTRAR CLIENTE  ---------------")
        ID_Cliente = generarID(clientes) 
        
        # Validacion Nombre
        while True:
            Nombre = input("Nombre del cliente: ")
            if Nombre == "":
                print("No se puede omitir.")
                continue
            else:
                break 
        clientes[ID_Cliente]=[Nombre]
        print("\nEl cliente ha sido guardado exitosamente!")
        
    # OPCION X. SALIR 
    elif (op_principal.upper() in "X"):
        print("FIN DE LA EJECUCIÓN! ")
        with open("reservaciones.csv","w",newline="") as archivo:
          grabador = csv.writer(archivo)
          grabador.writerow(( " idRESERVA ", " idSALA "," CLIENTE " , " EVENTO ", "ID TURNO"," TURNO "," FECHA "))
          grabador.writerows([(clave,datos[0],datos[1],datos[2],datos[3],datos[4],datos[5]) for clave,datos in eventos.items()])
        with open("salas.csv","w",newline="") as archivo:
          grabador = csv.writer(archivo)
          grabador.writerow(( "idSALA ", " NOMBRE ", " CAPACIDAD "))
          grabador.writerows([(clave,datos[0],datos[1]) for clave,datos in Salas.items()])
        with open("clientes.csv","w",newline="") as archivo:
          grabador = csv.writer(archivo)
          grabador.writerow(( "idCLIENTE ", " NOMBRE "))
          grabador.writerows([(clave,datos[0]) for clave,datos in clientes.items()])
        with open("ocupado.csv","w",newline="") as archivo:
          grabador = csv.writer(archivo)
          grabador.writerow(( " idRESERVA ", " idSALA ", " idTURNO ", " FECHA "))
          grabador.writerows([(clave,datos[0],datos[1],datos[2]) for clave,datos in ocupado.items()])
        break
    else:
        print("OPCION NO VALIDA. Debes ingresar la letra que corresponda a la acción deseada.")
    input("\nPresione cualquier tecla para continuar...")
    os.system("cls")
