import sqlite3
from sqlite3 import Error
import datetime
import sys
import os
from tarfile import ExtractError
import openpyxl
  
def verificarId(id,tabla):
  with sqlite3.connect("COWORKING.db") as conn:
    mi_cursor = conn.cursor() 
    mi_cursor.execute("SELECT * FROM "+tabla+" WHERE id"+tabla+" = "+str(id))
    valor = mi_cursor.fetchall()
    if not valor:
      return False
    else:
      return True

#Esta funcion se usara para todo lo que tenga un ingreso de informacion como INSERT UPDATE DELETE
def IngresoBD(comando,valores):
  try:
    with sqlite3.connect("COWORKING.db") as conn:
      mi_cursor = conn.cursor()
      mi_cursor.execute(comando,valores)
  except Error as e:
    print(e)
    return False
  except:
    print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    return False
  finally:
      conn.close()
      return True

#Esta funcion se usa para cualquier tipo de extraccion de informacion de la BD como SELECT
def ExtraccionBD(comando,valores):
  try:
    with sqlite3.connect("COWORKING.db") as conn:
      mi_cursor = conn.cursor()
      mi_cursor.execute(comando,valores)
      return mi_cursor.fetchall()
  except Error as e:
    print(e)
  except:
    print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
  finally:
      conn.close()
      
# Generamos la base de datos y las tablas para el script
try:
    with sqlite3.connect("COWORKING.db") as conn:
        mi_cursor = conn.cursor()
        mi_cursor.execute("""CREATE TABLE IF NOT EXISTS Sala (
        idSala INTEGER PRIMARY KEY,
        nombreSala TEXT NOT NULL,
        cupoSala INTEGER);""")

        mi_cursor.execute("""CREATE TABLE IF NOT EXISTS Cliente (
        idCliente INTEGER PRIMARY KEY,
        nombreCliente TEXT NOT NULL);""")

        mi_cursor.execute("""CREATE TABLE IF NOT EXISTS Turno (
        idTurno INTEGER PRIMARY KEY,
        nombreTurno TEXT NOT NULL);""")

        mi_cursor.execute("""CREATE TABLE IF NOT EXISTS Reserva (
        idReserva INTEGER PRIMARY KEY,
        nombreReserva TEXT NOT NULL,
        fechaReserva TEXT NOT NULL,
        cliente INTEGER NOT NULL,
        sala INTEGER NOT NULL,
        turno INTEGER NOT NULL,
        FOREIGN KEY(cliente) REFERENCES sala(idCliente),
        FOREIGN KEY(sala) REFERENCES cliente(idSala),
        FOREIGN KEY(turno) REFERENCES turno(idTurno));""")
        pass
except Error as e:
    print(e)

# Agregamos los turnos a la tabla "turno".
try:
    with sqlite3.connect("COWORKING.db") as conn:
      mi_cursor = conn.cursor()
      mi_cursor.execute("INSERT INTO turno VALUES(1,'Matutino')")
      mi_cursor.execute("INSERT INTO turno VALUES(2,'Vespertino')")
      mi_cursor.execute("INSERT INTO turno VALUES(3,'Nocturno')")
except Error as e:
  print(e)
except:
  print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
finally:
  conn.close()

# ---- M E N U   P R I N C I P A L ----
os.system("cls")
while True:      
    print("-"*50)
    print("MENÚ PRINCIPAL")
    print("\t [A] Reservaciones.")
    print("\t [B] Reportes.")
    print("\t [C] Registrar Nueva Sala.")
    print("\t [D] Registrar Nuevo Cliente.")
    print("\t [X] Salir.")
    op_principal = input("OPCION: ")
    os.system("cls")

   # OPCION A. RESERVACIONES
    if (op_principal.upper() == "A"):
      while True:
        # Recolección de todos los registros en la BD
        valores = {}
        registrosClientes = ExtraccionBD("SELECT * FROM Cliente ORDER BY idCliente",valores)
        registrosSalas = ExtraccionBD("SELECT * FROM Sala ORDER BY idSala",valores)
        registrosTurnos = ExtraccionBD("SELECT * FROM Turno ORDER BY idTurno",valores)
        registrosReservas = ExtraccionBD("SELECT * FROM Reserva ORDER BY idReserva",valores)
        
        # Submenú Reservaciones
        print("-"*60)
        print("Menu de Reservaciones")
        print("\t [A] Registrar Nueva Reservación.")
        print("\t [B] Modificar Nombre de una Reservación.")
        print("\t [C] Consultar Disponibilidad de salas para una fecha.")
        print("\t [D] Eliminar una Reservación.")
        print("\t [X] Regresar al menú principal. ")
        op_reserva = input("OPCION: ")
        os.system("cls")

        # A. Registrar Nueva Reservación
        if(op_reserva.upper() == "A"):
          if not registrosClientes:
            input("\nLo sentimos, primero debe registrarse.")
            continue
          elif not registrosSalas:
            input("\nNo hay salas registradas")
            continue    
          
          # Impresion de cliente
          print("-" * 54)
          print(f'|{"RESERVACION DE SALA":^52}|')
          print("-" * 54)
          print(f'|{"Clientes":^52}|')
          print("-" * 54)
          print(f'|{" Numero Cliente":^15}{"Nombre":^37}|') 
          print("-" * 54)
          for idCliente,nombreCliente in registrosClientes:
            print(f'|{idCliente:^15}{nombreCliente:^37}|')
          print("-" * 54)
          print()
          
          # Validacion de cliente
          while True:
            try:
              cliente = int(input("Ingrese folio de cliente: "))
              if cliente > 0:
                if verificarId(cliente,"Cliente"):
                  break
                else:
                  print("Numero de cliente no valido")
              else:
                  print("Numero de cliente no valido")
            except Exception:
              print("Ocurrió un problema numero no valido")
              print("Ingrese de nuevo...")
          os.system("cls")

          print("-" * 54)
          print(f'|{"RESERVACION DE SALA":^52}|')
          print("-" * 54)

          # Validacion de fecha
          while True:
            fecha_str = input("\nFecha del evento dd/mm/aaaa: ") 
            try: 
              fechaReserva = datetime.datetime.strptime(fecha_str,"%d/%m/%Y").date()               
              fecha_actual = datetime.date.today()
              fecha_invalida = fecha_actual + datetime.timedelta(days=+1)
              if fechaReserva <=  fecha_invalida:
                print(f"Fecha Rechazada.Se necesita al menos 2 días de anticipación.")
                continue
              else:
                break
            except ValueError:
              print(f"ERROR! Formato de fecha NO valida.")
              print("Ingrese de nuevo...")
          os.system("cls")

          # Impresion de sala
          print("-" * 54)
          print(f'|{"RESERVACION DE SALA":^52}|')
          print("-" * 54)          
          print(f'|{"Salas":^52}|')
          print("-" * 54)
          print(f'|{" Numero Sala":^15}{"Nombre":^27}{"Capacidad ":^10}|')
          print("-" * 54)
          for idSala,nombreSala,cupoSala in registrosSalas:
              print(f'|{idSala:^15}{nombreSala:^27}{cupoSala:^10}|')
          print("-" * 54)
          print()

          # Validacion de sala
          while True:
            try:
              sala = int(input("Ingrese folio de sala: "))
              if sala > 0:
                if verificarId(sala,"Sala"):
                  break
                else:
                  print("Numero de sala no valido")
              else:
                print("Numero de sala no valido")
            except Exception:
              print(f"Ocurrió un problema numero no valido")
              print("Ingrese de nuevo...")
          os.system("cls")

          # Impresion de turno         
          print("-" * 54)
          print(f'|{"RESERVACION DE SALA":^52}|')
          print("-" * 54)      
          print(f'|{"Turnos":^52}|')
          print("-" * 54)
          print(f'|{" Numero Turno":^15}{"Turno":^37}|')
          print("-" * 54)
          for idTurno,nombreTurno in registrosTurnos:
              print(f"|{idTurno:^15}{nombreTurno:^37}|")
          print("-" * 54)
          print()
          
          # Validacion de turno
          while True:
            try:
              turno = int(input("Ingrese folio de turno: "))
              if turno > 0:
                if verificarId(turno,"Turno"):
                  break
                else:
                  print("Numero de turno no valido")
              else:
                print("Numero de turno no valido")
            except Exception:
              print(f"Ocurrió un problema numero no valido")
              print("Ingrese de nuevo...")
          os.system("cls")

          # Validar que no se empalmen los eventos
          valores = {"fechaReserva":fechaReserva,"idSala":sala,"idTurno":turno}
          if not ExtraccionBD("SELECT * FROM Reserva WHERE fechaReserva = :fechaReserva and sala = :idSala and turno = :idTurno",valores):
            print("-" * 54)
            print(f'|{"RESERVACION DE SALA":^52}|')
            print("-" * 54)    
            # Validacion de nombre
            while True:
              nombreReserva= input('\nNombre para la reservacion: ')
              if nombreReserva == "":
                  print("No se puede omitir.")
                  continue
              else:
                  break     
            # Guardar Reservación en base de datos
            valores = (nombreReserva, fechaReserva, cliente, sala, turno )
            if IngresoBD("INSERT INTO Reserva(nombreReserva, fechaReserva, cliente, sala, turno) VALUES(?,?,?,?,?)", valores):
              input("\nLa reservacion ha sido guardado exitosamente!")
          else:
            input("\nError la reservacion coincide con una reserva ya realizada.")
          os.system("cls")

        # B. Modificar Nombre de una Reservación.
        elif(op_reserva.upper() == "B"):
          print("-" * 54)
          print(f'|{"EDITAR RESERVACION":^52}|')
          print("-" * 54)
          print()
          if not registrosReservas:
            print("NO hay eventos registrados.")
            continue
          else:
            # Impresion de reservaciones
            print(f'{"  Reservaciones  ":-^54}')
            print("-" * 54)
            print(f'{"Numero Reserva":^14}{"Nombre":^31}{"Fecha":^9}')
            print("-" * 54)
            for idReserva,nombreReserva,fechaReserva,cliente,sala,turno in registrosReservas:
                print(f"{idReserva:^14}{nombreReserva:^29}{fechaReserva:^12}")
            print("-" * 54)
            print()
            
          # Validacion de id de reservacion
          while True:
            try:
              reserva = int(input("Ingrese folio de reservacion: "))
              if reserva > 0:
                if verificarId(reserva,"Reserva"):
                  break
                else:
                  print("Numero de reserva no valido")
              else:
                print("Numero de reserva no valido")
            except Exception:
              print(f"Ocurrió un problema numero no valido")
              print("Ingrese de nuevo...")
          os.system("cls")
          
          # Validacion del nuevo nombre para la reservacion
          while True:
            nuevoNombre = input('Nuevo nombre para el evento: ')
            if nuevoNombre == "":
              print('Nombre no valido!')
            else:
              break
            
          # Guardar en base de datos.
          valores = {"nombre":nuevoNombre,"idReserva":reserva}
          if IngresoBD("UPDATE Reserva SET (nombreReserva) = :nombre WHERE (idReserva) = :idReserva", valores):
            input('\nSe edito correctamente el nombre del evento!')
          os.system("cls")
  
        # C. Consultar Disponibilidad de salas para una fecha.
        elif(op_reserva.upper() == "C"):
          print("-" * 54)
          print(f'|{"DISPONIBILIDAD DE SALAS":^52}|')
          print("-" * 54)
          while True:
            fechaBuscada = input("\nIngrese la fecha que desea consultar dd/mm/aaaa: ")
            try:
              fechaBuscada = datetime.datetime.strptime(fechaBuscada,"%d/%m/%Y").date()
              os.system("cls")
              break
            except:
              print(f"ERROR fecha no valida!")
          
          #Extraer informacion de la base de datos
          valores = {"fechaReserva":fechaBuscada}
          reservaciones_actuales = ExtraccionBD("SELECT sala.idSala,sala.nombreSala, turno.nombreTurno FROM Reserva INNER JOIN sala on Reserva.sala = sala.idSala INNER JOIN turno on Reserva.turno = turno.idTurno where Reserva.fechaReserva = :fechaReserva",valores)
          valores = {}
          eventos_posibles = ExtraccionBD("SELECT sala.idSala,sala.nombreSala, turno.nombreTurno FROM sala, turno",valores)
          
          print(f'**{"Salas disponibles para renta el dia "+str(fechaBuscada):^52}**')
          print()
          print(f'{"Sala":<26}{"Turno":<26}')
          print()
          for evento in reservaciones_actuales:
            eventos_posibles.remove(evento)
          for idSala,nombreSala, nombreTurno in eventos_posibles:
            print(f'{str(idSala)+". "+nombreSala:<26}{nombreTurno:<26}')
          input()
          os.system("cls")
          
        # D. Eliminar una Reservación.
        elif (op_reserva.upper() == "D"):
          if not registrosReservas:
            print("NO hay eventos registrados.")
            continue
          else:
            # Impresión de todas las reservaciones.
            print("-" * 54)
            print(f'|{"ELIMINAR RESERVACION":^52}|')
            print("-" * 54)
            print(f'|{"Reservaciones  ":^52}|')
            print("-" * 54)
            print(f'|{"Numero Reserva":^14}{"Nombre":^28}{"Fecha":^10}|')
            print("-" * 54)
            for idReserva,nombreReserva,fechaReserva,cliente,sala,turno in registrosReservas:
                print(f'|{idReserva:^14}{nombreReserva:^28}{fechaReserva:^10}|')
            print("-" * 54)
            print()
            
          # Validación del id de la reservacion
          while True:
            try:
              reserva = int(input("Ingrese folio de sala: "))
              if verificarId(reserva,"Reserva"):
                os.system("cls")
                break
              else:
                print("Opcion no valida.")
                continue
            except:
              print("ERROR! Respuesta no valida")


          # Extraccion de informacion de la reservacion
          valores = {"idReserva":reserva}
          reservacion = ExtraccionBD("SELECT * FROM Reserva WHERE (idReserva) = :idReserva",valores)
          
          # Validacion de fecha para saber si se puede eliminar
          fecha_actual = datetime.date.today()
          fecha_invalida = fecha_actual + datetime.timedelta(days=+2)
          for idReserva,nombreReserva,fechaReserva,cliente,sala,turno in reservacion:
            if fechaReserva <=  str(fecha_invalida):
              print(f"\nFecha Rechazada.Se necesitan al menos 3 días de anticipación para cancelar la reservacion.")
              input()
              break
            else:
              # Impresion de reservacion a eliminar
              print("-" * 54)
              print(f'|{"ELIMINAR RESERVACION":^52}|') 
              print("-" * 54)
              print(f'|{"Numero Reserva":^14}{"Nombre":^28}{"Fecha":^10}|')
              print("-" * 54)    
              for idReserva,nombreReserva,fechaReserva,cliente,sala,turno in reservacion:
                print(f'|{idReserva:^14}{nombreReserva:^28}{fechaReserva:^10}|')
                print("-" * 54)
              while True:
                op_eliminar=input("\n¿Esta seguro de eliminar la reservacion? (A = Si / B = No): ")
                if (op_eliminar.upper() == "A"):
                  # Eliminacion de la reservacion
                  valores = {"idReserva":reserva}
                  if IngresoBD("DELETE FROM Reserva WHERE (idReserva) = :idReserva",valores):
                    print("\nSe cancelo correctamente la reservacion!")
                    break
                if (op_eliminar.upper() == "B"):
                  break
                else:
                  input("OPCION NO VALIDA. Debe ingresar la letra que corresponda a la acción deseada.")
          os.system("cls")
              
        # Salir del submenú.
        elif(op_reserva.upper() == "X"):
          break
        
        else:
          input("OPCION NO VALIDA. Debe ingresar la letra que corresponda a la acción deseada.")
          os.system("cls")

   # OPCION B. REPORTES
    elif (op_principal.upper() == "B"):
      while True:
        print("-"*50)
        print("MENU REPORTES")
        print("\t [A] Reporte en Pantalla")
        print("\t [B] Reporte en Excel.")
        print("\t [X] Regresar al menú principal. ")
        op_reporte = input("OPCION: ")
        os.system("cls")

        # A. Reporte en Pantalla
        if (op_reporte.upper() == "A"):
          # Extraccion de informacion de reportes de la BD
          valores={}
          todasReservas = ExtraccionBD("SELECT Reserva.sala, Cliente.nombreCliente, Reserva.nombreReserva, Turno.nombreTurno, Reserva.fechaReserva FROM Reserva INNER JOIN Cliente ON Reserva.cliente = Cliente.idCliente INNER JOIN Turno on Reserva.turno = Turno.idTurno",valores)
          
          #Impresion de Reporte
          print("*"*85)
          print(f'**{"REPORTE DE RESERVACIONES ":^81}**')
          print("*"*85)
          print(f'{"SALA":<10}{"CLIENTE":<20}{"EVENTO":<20}{"TURNO":<20}{"FECHA":<26}')
          print('*'*85)
          for idSala,nombreCliente,nombreReserva,Turno,fechaReserva in todasReservas:
            print(f'{idSala:<10}{nombreCliente:<20}{nombreReserva:<20}{Turno:<20}{fechaReserva:<26}')
          print(f'{"FIN DEL REPORTE":*^85}')

          # Validacion de fecha
          while True:
            fecha_str = input("\nIngrese la fecha del evento a buscar dd/mm/aaaa: ") 
            try: 
              fechaReserva = datetime.datetime.strptime(fecha_str,"%d/%m/%Y").date()               
              break
            except ValueError:
              print(f"ERROR! Formato de fecha NO valida.")
              print("Ingrese de nuevo...")
          os.system("cls")

          # Extraccion de informacion de reportes de la BD con fecha especifica
          valores = {"fechaReserva":fechaReserva}
          registrosReservasFecha = ExtraccionBD("SELECT Reserva.sala, Cliente.nombreCliente, Reserva.nombreReserva, Turno.nombreTurno FROM Reserva INNER JOIN Cliente ON Reserva.cliente = Cliente.idCliente INNER JOIN Turno on Reserva.turno = Turno.idTurno WHERE Reserva.fechaReserva = :fechaReserva",valores)
          
          #Impresion de Reporte por fecha
          print("*"*85)
          print(f'**{"REPORTE DE RESERVACIONES PARA EL DIA "+fecha_str:^81}**')
          print("*"*85)
          print(f'{"SALA":<7}{"CLIENTE":<33}{"EVENTO":<33}{"TURNO":<12}')
          print('*'*85)
          for idSala,nombreCliente,nombreReserva,Turno in registrosReservasFecha:
            print(f'{idSala:<7}{nombreCliente:<33}{nombreReserva:<33}{Turno:<12}')
          print(f'{"FIN DEL REPORTE":*^85}')
          input()
          os.system("cls")
        
        # B. Reporte en Excel
        elif (op_reporte.upper() == "B"):
          #Extraer reservaciones de la BD
          valores = {}
          registrosreserva = ExtraccionBD("SELECT Reserva.sala, Cliente.nombreCliente, Reserva.nombreReserva, Turno.nombreTurno, Reserva.fechaReserva FROM Reserva INNER JOIN Cliente ON Reserva.cliente = Cliente.idCliente INNER JOIN Turno on Reserva.turno = Turno.idTurno",valores) 
          
          #Exportaccion para un excel
          libro = openpyxl.Workbook()
          #libro.iso_dates = True #Acepte fechas              
          hoja = libro["Sheet"] 
          hoja.title = "Reporte"
          hoja.append(('SALA', 'CLIENTE', 'EVENTO', 'TURNO','FECHA'))
          for idSala,nombreCliente,nombreReserva,Turno,fechaReserva in registrosreserva:
            reporte=idSala,nombreCliente,nombreReserva,Turno,fechaReserva
            hoja.append(reporte)
          libro.save("Reporte.xlsx")
          input("\nSe exporto en excel las reservaciones exitosamente!")
          os.system("cls")
        
        # X. Salida del menú reportes.
        elif(op_reporte.upper() == "X"):
          print("Has salido del Menú Reportes.")
          break
        else:
          input("OPCION NO VALIDA. Debe ingresar la letra que corresponda a la acción deseada.")
          os.system("cls")
 
    # OPCION C. RESERVAR NUEVA SALA
    elif (op_principal.upper() == "C"):
      print("-" * 54)
      print(f'|{"REGISTRAR SALA":^52}|')
      print("-" * 54)
      print()
      while True:
        nombreSala = (input("Nombre para la sala: "))
        if nombreSala == "":
          print("No se puede omitir.")
          continue
        else:
          break
      while True:
        try:
          cupoSala=int(input("Capacidad de la sala: "))
          if cupoSala >0:
              break
          else:
              print("Ocurrió un problema no puede ingresar numeros menores a 0")
        except Exception:
          print(f"Ocurrió un problema numero no valido")
          print("Ingrese de nuevo...")

      # Guardar en base de datos.
      valores = (nombreSala, cupoSala)
      if IngresoBD("INSERT INTO Sala( nombreSala, cupoSala ) VALUES(?,?)",valores):
        input("\nLa sala ha sido guardada exitosamente!")
      os.system("cls")

    # OPCION D. Registrar un cliente
    elif (op_principal.upper() == "D"):
      print("-" * 54)
      print(f'|{"REGISTRAR CLIENTE":^52}|')
      print("-" * 54)
      print()
      while True:
        nombreCliente = input("Nombre del cliente: ")
        if nombreCliente == "":
          print("No se puede omitir.")
          continue
        else:
          break
      # Guardar en base de datos.
      valores = (nombreCliente,)
      if IngresoBD("INSERT INTO Cliente(nombreCliente) VALUES(?)",valores):
        input("\nEl cliente ha sido guardado exitosamente!")
      os.system("cls")

    # OPCION X Salir del menú principal.
    elif (op_principal.upper() == "X"):
      break
    else:
      input("OPCION NO VALIDA. Debe ingresar la letra que corresponda a la acción deseada.")
      os.system("cls")
