import datetime
import time
import sys
import os

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
    print("*"*60)
    print(f'**           REPORTE DE TODAS LAS RESERVACIONES           **')
    print("*"*60) 
    print("{:<7}{:<17}{:<14}{:<12}{:<9}".format("SALA","CLIENTE","EVENTO","TURNO","FECHA"))
    print("*"*60)
    
    for id_fecha in eventos:
          print("{:<7}{:<17}{:<14}{:<12}{:<7}".format(eventos[id_fecha][0],eventos[id_fecha][1],eventos[id_fecha][2],eventos[id_fecha][3],eventos[id_fecha][4]))
    print('********************* FIN DEL REPORTE **********************')
    while True:
      try:
        input("\nPresione cualquier tecla para continuar...")
        os.system("cls")
        Fecha_reservacion= input('Fecha de la reservacion que desea buscar (dd/mm/aaaa): ')
        fecha_procesada = datetime.datetime.strptime(Fecha_reservacion, "%d/%m/%Y").date() 
        break   
      except:
            print(f"Ocurrió un problema {sys.exc_info()}")
            print("Ingrese de nuevo...")
    print("")
    print("*"*55)
    print(f'**  REPORTE DE RESERVACIONES PARA EL DIA {Fecha_reservacion}  **')
    print("*"*55)  
    print("{:<7}{:<17}{:<16}{:<12}".format("SALA","CLIENTE","EVENTO","TURNO"))
    print('*'*55)
    
    for id_fecha in eventos:
        if Fecha_reservacion == eventos[id_fecha][-1]:
          print("{:<7}{:<17}{:<16}{:<12}".format(eventos[id_fecha][0],eventos[id_fecha][1],eventos[id_fecha][2],eventos[id_fecha][3]))
    print('******************* FIN DEL REPORTE *******************')
      

# ESTRUCTURAS
turnos = ("Matutino","Vespertino","Nocturno")
Salas = {}
cliente = {}
eventos = {} 
ocupado = {}
 


while True:

  # INICIO
  print("-"*50)
  print("Menu de Opciones")
  print("\t [A] Reservar Sala.")
  print("\t [B] Editar Nombre del Evento Registrado.")
  print("\t [C] Consultar Reservaciones.")
  print("\t [D] Registrar Nuevo Cliente.")
  print("\t [E] Registrar Nueva Sala.")
  print("\t [X] Salir.")

  opcion = input("OPCION: ")
  os.system("cls")
  # Validacion
  
  # A. Reservar Sala 
  if (opcion.upper() in "A"):
      ID_Reservacion = generarID(eventos)
      print("---------------  RESERVACION DE SALA  ----------------")
      if not bool(cliente):
        print("Lo sentimos, primero debe registrarse.")
        input("\nPresione cualquier tecla para continuar...")
        continue
        
      elif not bool(Salas):
        print("No hay salas registradas")
        input("\nPresione cualquier tecla para continuar...")
        continue

      # Validacion Cliente
      print("CLIENTES")
      for id in cliente:
        print(id,'.-',cliente[id])
        print("-"*54)
      ID_Cliente = validar_ID(cliente,"Ingrese folio de cliente: ")
      os.system("cls")
      print("---------------  RESERVACION DE SALA  ----------------")
      #print(f"\nBienvenid@ {cliente[ID_Cliente]} :)")

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
            id = 0 # <- para mostrar un numero por turno
            # Imprime turno
            for nombre in turnos:
                id=id + 1
                print(id,".-",nombre)
            print("-"*54)
              # Validacion          
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
                eventos[ID_Reservacion] = [idSala,(cliente[ID_Cliente]),Nombre_Reservacion,turnos[ID_Turno-1], fecha_capturada]
                ocupado[ID_Reservacion] = [idSala, (ID_Turno-1), fecha_capturada] # Nos permitirá comprobar que no hayan eventos con misma fecha, sala y turno
                print("\nLa reservacion se ha guardado exitosamente!")
            else:
                # Validar que NO se empalme con otro evento registrado
                nueva_Reservacion = ([idSala,(ID_Turno-1),fecha_capturada])
                for registro in ocupado.values():
                  if registro == nueva_Reservacion:
                    print("Lo sentimos, sala OCUPADA!")
                    break
                else:
                  eventos[ID_Reservacion] = [idSala, (cliente[ID_Cliente]), Nombre_Reservacion, turnos[ID_Turno-1], fecha_capturada]
            break
        except ValueError:
            print(f"ERROR fecha no valida")
            print("Ingrese de nuevo...")


  # B. Editar Nombre del Evento
  elif (opcion.upper() in "B"):
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


  # C. Consultar Reservaciones."
  elif(opcion.upper() in "C"):
    if not bool(eventos):
      print("NO hay eventos registrados.")
      continue
    Consultar_Reservaciones()


  # D. Registrar nuevo cliente
  elif(opcion.upper() in "D"):
    print("--------------  REGISTRAR CLIENTE  ---------------")
    ID_Cliente = generarID(cliente) 
    
    # Validacion Nombre
    while True:
        Nombre = (input("Nombre del cliente: "))
        if Nombre == "":
            print("No se puede omitir.")
            continue
        else:
            break
        
    cliente[ID_Cliente]=Nombre
    print("\nEl cliente ha sido guardado exitosamente!")


  # E. Registrar nueva sala
  elif(opcion.upper() in "E"):  
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

  # SALIR
  elif(opcion.upper() in "X"):
   print("FIN DE LA EJECUCIÓN! ")
   break

  else:
    print("Opcion no valida!")
  input("\nPresione cualquier tecla para continuar...")
  os.system("cls")