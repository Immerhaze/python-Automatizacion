import pandas as pd
import pyautogui
import time

#Listado original del dropdown en Kawak, con estos indices se puede medir la altura en la que se encuentra cada opción para poder, 
#si es que es necesario scrollear para buscar la opción y mover le mouse con pyautogui hacia ella
cargo_dropdown = [
    "Seleccionar",
    "",
    "ADMINISTRATIVA REVERSA BDF",
    "Administrativo",
    "Administrativo de decomiso P&G",
    "Administrativo de Expedición Unilever",
    "Administrativo de Facturación Unilever",
    "Administrativo de Gestión de Stock",
    "Administrativo de Preparación Unilever",
    "Administrativo de Preparacion y Recepcion Tesa",
    "Administrativo de Recepción",
    "Administrativo de Recepción Unilever",
    "Administrativo de Recuperado Unilever",
    "Administrativo de RRHH",
    "Administrativo de Transporte",
    "Administrativo de Transporte",
    "Administrativo Logística Reversa",
    "Administrativo Logístico P&G",
    "Administrativo Noviciado",
    "Administrativo Operación Noviciado",
    "Administrativo Recuperado Unilever",
    "Administrativo VAS Noviciado",
    "Analista",
    "Analista de Calidad",
    "ANALISTA DE CALIDAD VAS P&G",
    "Analista de Facturación",
    "Analista de Facturación Unilever",
    "Analista de Gestión de Stock",
    "Analista de Gestión de Stock",
    "Analista de IT",
    "Analista de Planificación",
    "Analista de Planificación VAS P&G",
    "Analista de Recuperado P&G",
    "Analista de SHE Unilever",
    "Analista DO y Selección",
    "Analista Gestión de Personas",
    "Analista Gestión de Stock Unilever",
    "Analista Logistica Reversa Unilever",
    "Analista Logística Reversa Unilever",
    "Analista VAS",
    "Analistas en prevención de riesgos",
    "Apilador",
    "Apilador Unilever",
    "Aplilador Noviciado",
    "APOYO",
    "Asistente de Mantención Unilever",
    "Asistente de Mantención Unilever",
    "Asistente de Mantención Unilever",
    "ASISTENTE DE MANTENCION VAS P&G",
    "Auxiliar de Aseo",
    "BACKUP LIDER DE LINEA VAS P&G",
    "Bunkero Noviciado",
    "Bunkero Unilever",
    "Contract Manager",
    "Contratista MACLEAND",
    "Contratista R2G",
    "Contratista Royal América",
    "Control",
    "Control de Calidad Cliente Unilever",
    "Control de Calidad Noviciado",
    "Control de carga y descarga P&G",
    "Control de Embarque Unilever",
    "Control de Gestión de Stock",
    "Control de insumos P&G",
    "Control de Insumos VAS",
    "Control de Merma Unilever",
    "Control de Operaciones BDF - TESA",
    "Control de Operaciones doTerra",
    "Control de Operaciones SEB - PEYA",
    "Control de Preparación",
    "Control de Preparación Unilever",
    "Control de Recepción Noviciado",
    "Control de Recepción Unilever",
    "Control Documental de Transporte",
    "Control GDS Noviciado",
    "Control GDS P&G",
    "Control Noviciado",
    "Control Patrimonial",
    "Control Patrimonial Unilever",
    "Control Unilever",
    "COORDINADOR DE CARGA Y DESCARGA P&G",
    "Coordinador de Cuadrilla Unilever",
    "Coordinador de Mantención y Servicios Generales Noviciado",
    "Coordinador de Mantenimiento",
    "Coordinador de Mantenimiento Unilever",
    "Coordinador VAS",
    "Coordinadora de Transporte",
    "Coordinadores de Cuadrilla",
    "Copacking",
    "COPACKING VAS P&G ",
    "CORDINADOR DE CARGA Y DESCARGA",
    "COUNTER BALANCE P&G",
    "Cuadrillero",
    "Desarrollador Full Stack IT",
    "Despicking Unilever",
    "Director General",
    "Director Técnico Droguería Noviciado",
    "Encargada de Calidad VAS Unilever",
    "Encargada de LLRR, Mermas y Recuperado Unilever",
    "Encargada de Mantención y Servicios Generales Centro Lampa - Unilever",
    "Encargada de Operaciones Lipton",
    "Encargada de Operaciones P&G",
    "Encargada de Recepción Unilever",
    "Encargada de Reclutamiento y Selección",
    "Encargada de VAS",
    "ENCARGADA DE VAS P&G",
    "Encargada Medioambiente",
    "Encargada Seguridad Patrimonial",
    "Encargada/0 de Planificación VAS",
    "Encargado",
    "Encargado de BMP y Super Market P&G",
    "Encargado de Calidad Noviciado",
    "Encargado de Calidad P&G",
    "Encargado de Gestión de Stock",
    "Encargado de HSE (BAT)",
    "Encargado de Operaciones",
    "Encargado de Operaciones BAT",
    "Encargado de operaciones BMP",
    "Encargado de Operaciones Lipton",
    "Encargado de Operaciones Noviciado",
    "Encargado de Operaciones Unilever",
    "Encargado de Remuneraciones",
    "Encargado de SHE P&G",
    "ENCARGADO DE SHE VAS P&G",
    "Encargado de Transporte",
    "Encargado IT",
    "Experto en prevención de riesgos",
    "Experto SHE Unilever",
    "Facturador P&G",
    "Gerente",
    "Gerente de Calidad y CSR",
    "Gerente de Cuenta P&G",
    "Gerente de Cuenta Unilever",
    "Gerente de Cuentas - Noviciado",
    "Gerente de Finanzas",
    "Gerente de Gestión de Stock",
    "Gerente de Implementación",
    "Gerente de Ingeniería",
    "Gerente de IT",
    "Gerente de Operaciones",
    "Gerente de Operaciones BAT",
    "Gerente de RRHH",
    "Gerente de SHE y Seguridad",
    "Gerente de Valor Agregado (VAS)",
    "Gestión Finanzas",
    "Gestión Personas P&G",
    "Grúa Frontal (Unilever)",
    "Guardia",
    "Ingeniera Desarrollo Negocios",
    "Ingeniero de Procesos",
    "Ingeniero de Transporte",
    "Ingeniero P&G",
    "Inspector MHE",
    "Instructor MHE Unilever",
    "Jefe",
    "Jefe de Calidad Noviciado",
    "Jefe de Calidad Unilever",
    "Jefe de Gestión Patrimonial",
    "Jefe de Operaciones",
    "Jefe de Operaciones BAT",
    "Jefe de Operaciones Maipo",
    "Jefe de Operaciones P&G",
    "Jefe de SHE Unilever",
    "Jefe de SHE y Seguridad",
    "Jefe de Transporte",
    "Jefe de Turno Unilever",
    "Jefe de VAS",
    "Jefe DO y Selección",
    "Jefe Prevención de Riesgos",
    "Líder Control de Gestión",
    "LIDER DE CALIDAD VAS P&G",
    "Líder de Línea VAS",
    "LIDER DE LÍNEA VAS P&G",
    "Mantención y Servicios Generales Noviciado",
    "Mantenimiento Unilever",
    "Noviciado",
    "Operador carga y descarga P&G",
    "Operador CCTV Unilever",
    "Operador de Mantención",
    "Operador de Maquina BMP P&G",
    "Operador de recuperado P&G",
    "Operador Maquinista Unilever",
    "OPERADOR MAQUINISTA VAS P&G",
    "Operario BAT",
    "Operario Movilizador Súper Market P&G",
    "Operario VAS Noviciado",
    "Operario VAS Unilever",
    "Picking (Unilever)",
    "Picking - BAT",
    "Picking Carga y Descarga",
    "Picking doTerra",
    "Picking Noviciado",
    "Picking P&G",
    "Planificación P&G",
    "Procter",
    "Reach truck P&G",
    "Repaletizador",
    "Sales and Marketing Director",
    "Supervisor",
    "Supervisor de Expedición (UL)",
    "Supervisor de Gestión de Stock P&G",
    "Supervisor de Operaciones",
    "Supervisor de Preparación",
    "Supervisor de Rechazo y Devoluciones Unilever",
    "Supervisor de rechazos y devoluciones comerciales P&G",
    "Supervisor de Turno P&G",
    "Supervisor de Turno Unilever",
    "Supervisor GDS Unilever",
    "Supervisor Guardia",
    "Supervisor Operaciones BAT",
    "Supervisor Operaciones P&G",
    "Supervisor Operaciones SEB",
    "Supervisor Unilever",
    "Supervisor VAS",
    "Supervisora de Devoluciones y Rechazos P&G",
    "Supervisora Operaciones VAS Noviciado",
    "Unilever"
]



# CARGAR DATOS DEL ARCHIVO EXCEL A UNA VARIABLE
try:
    df = pd.read_excel('profile_data.xlsx', sheet_name='Hoja1')
except FileNotFoundError:
    print("Error: File 'profile_data.xlsx' not found.")
    exit()

# FUNCIÓN PARA CALCULAR LA ALTURA DE LA OPCIÓN DENTRO DELDROPDOWN MENU DE CARGO 
def select_cargo(cargo_name):
    try:
        # Click on the cargo input field at the given coordinates
        pyautogui.press('enter')
        # Wait to ensure the dropdown menu opens
        time.sleep(2)

        # Convert both cargo_name and cargo_dropdown items to lowercase for case-insensitive search
        cargo_dropdown_lower = [item.lower() for item in cargo_dropdown]
        cargo_name_lower = cargo_name.lower()

        if cargo_name_lower in cargo_dropdown_lower:
            index = cargo_dropdown_lower.index(cargo_name_lower) + 1
            print(f"Index of '{cargo_name}': {index}")
        else:
            print(f"'{cargo_name}' not found in the dropdown options.")
            return False

        # Press the down arrow key the required number of times to select the item
        pyautogui.press('down',presses=index)
        time.sleep(0.5)
        # Press enter to select the item
        pyautogui.press('enter')
        pyautogui.press('tab')

        return True

    except Exception as e:
        print(f"Error while selecting cargo: {e}")
        return False

def selec_jefe(jefe_name):
    try:
         pyautogui.write(jefe_name)
         pyautogui.press('tab')
         pyautogui.press('down')
         pyautogui.press('tab')
         pyautogui.press('enter')
         pyautogui.press('tab',presses = 3)


    except Exception as e:
        print(f"Error choosing boss : {e}")


def selec_usuario():
    try:
         pyautogui.press('enter')
         pyautogui.press('down', presses= 3)
         time.sleep(0.5)
         pyautogui.press('enter')
         pyautogui.press('tab',presses = 2)

    except Exception as e:
        print(f"Error choosing user type : {e}")

def selec_area():
    try:

         pyautogui.press('enter')
         pyautogui.press('down')
         time.sleep(0.5)
         pyautogui.press('enter')
         pyautogui.press('tab')

    except Exception as e:
        print(f"Error choosing work area  : {e}")



def selec_centro():
    try:

         centro = "procter"
         gen = "general"

         pyautogui.write(centro)
         pyautogui.press('tab')
         pyautogui.press('down')
         pyautogui.press('tab')
         pyautogui.press('enter')
         pyautogui.hotkey('shift','tab')
         pyautogui.hotkey('shift','tab')
         pyautogui.press('backspace')
         pyautogui.write(gen)
         pyautogui.press('tab')
         pyautogui.press('down')
         pyautogui.press('tab')
         pyautogui.press('enter')




    except Exception as e:
        print(f"Error selecting work center : {e}")

def crear_perfil(row):
    try:
        time.sleep(4)  # Wait to prepare the environment

        # Navigate to the page to add
        pyautogui.press('tab',presses = 9)
        pyautogui.press('enter')

        time.sleep(4)  # Wait to prepare the environment
        pyautogui.press('tab',presses = 9)

        pyautogui.write(str(row['NOMBRE']))  # Assuming 'NOMBRE' is the correct column name

        pyautogui.press('tab')

        # Write the surname
        pyautogui.write(str(row['APELLIDOS']))  # Assuming 'APELLIDOS' is the correct column name

        pyautogui.press('tab')

        # Select the cargo
        select_cargo(row['CARGO'])

        selec_jefe(row['Jefe inmediato'])

        selec_usuario()

        selec_area()

        pyautogui.press('enter')


        pyautogui.write(str(row['CARGO']))  # Assuming 'NOMBRE' is the correct column name
        pyautogui.press('tab')

        # selec_permisos()
        selec_centro()
        pyautogui.press('tab',presses = 7)

        # pyautogui.sleep(5)
        # pyautogui.press('enter')
        # pyautogui.sleep(5)


    except Exception as e:
        print(f"Error creating profile for {row['NOMBRE']}: {e}")

# Extract unique cargo items from the DataFrame
cargo_items = df['CARGO'].unique()

# # Iterate over each row in the DataFrame and create the profile
for index, row in df.iloc[13:].iterrows():
    crear_perfil(row)







# input_nombre(row['CORREO ELECTRONICO'])

        # Write the name
        # pyautogui.click(email)
        # # Split email address and input before/after '@'
        # email_address = str(row['CORREO ELECTRONICO'])
        # email_parts = email_address.split('@')
        # if len(email_parts) == 2:
        #     user_email, email_domain = email_parts
        #     pyautogui.click(email)
        #     pyautogui.write(user_email)
        #     pyautogui.hotkey('alt', 'ctrl', 'q')
        #     pyautogui.write(email_domain)
        # else:
        #     print(f"Invalid email format: {email_address}")

        # # Write the name
        # input_contraseña(row['APELLIDOS'])

        # input_contraseña_confirmacion(row['APELLIDOS'])
        
        
        # def input_nombre(email):
#     try:
        
#         pyautogui.click(nombre_usuario_coords)
#           # Split nombre and apellidos into part
#         email_parts = email.split('@')
#         if len(email_parts) == 2:
#             user_email, email_domain = email_parts
#             pyautogui.click(email)
#             pyautogui.write(user_email)
#             pyautogui.hotkey('alt', 'ctrl', 'q')
#             pyautogui.write(email_domain)
#         else:
#             print(f"Invalid email format: {email}")
        
  

#     except Exception as e:
#         print(f"Error imputing  user name : {e}")
        
# def input_contraseña(apellido):
#     try:
#           # Split nombre and apellidos into parts
#         apellido_parts = apellido.split()

#         # Extract first name and first surname
#         first_surname = apellido_parts[0] if apellido_parts else ''

#         # Create new variable with first name and first surname
#         user_name = f"{first_surname}12345678"
#         pyautogui.click(contraseña)
#         pyautogui.write(user_name)

#     except Exception as e:
#         print(f"Error imputing password : {e}")
        
# def input_contraseña_confirmacion(apellido):
#     try:
#           # Split nombre and apellidos into parts
#         apellido_parts = apellido.split()

#         # Extract first name and first surname
#         first_surname = apellido_parts[0] if apellido_parts else ''

#         # Create new variable with first name and first surname
#         user_name = f"{first_surname}12345678"
#         pyautogui.click(confirmacion_contraseña)
#         pyautogui.write(user_name)

#     except Exception as e:
#         print(f"Error imputing confirmation : {e}")

# def selec_permisos():
#     try:
#          permiso = "consulta"
#          pyautogui.click(filtro_permisos)
#          pyautogui.write(permiso)
#          pyautogui.click(permisos)
#          pyautogui.click(btn_permisos)

#     except Exception as e:
#         print(f"Error selecting permissions : {e}")