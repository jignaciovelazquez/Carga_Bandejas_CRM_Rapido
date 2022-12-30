from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from bs4 import BeautifulSoup

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import pandas as pd

import re

import datetime
import time

columna = 0
campo = ""
cierre = ""
vector = []


path = "/chromedriver.exe"
Service = Service(executable_path=path)
driver = webdriver.Chrome(service=Service)
wait = WebDriverWait(driver, 5)
# driver.maximize_window()
driver.minimize_window()
# ------------------------------- loggin -----
driver.get("http://jvelazquez:Nacho123-@crm.telecentro.local/MembersLogin.aspx")

wait.until(EC.element_to_be_clickable(
    (By.XPATH, '//*[@id="txtPassword"]'))).send_keys("Nacho123-")

wait.until(EC.element_to_be_clickable(
    (By.XPATH, '//*[@id="btnAceptar"]'))).click()

# -------------------------------Bandeja de Cierre de Relevamiento-----

driver.get(
    "http://crm.telecentro.local/Edificio/Gt_Edificio/BandejaEntradaDeRelevamiento.aspx?TituloPantalla=Descarga%20De%20Relevamiento&EstadoGestionId=5&TipoGestionId=3&TipoGestion=OPERACIONES%20DE%20RED%20-%20CIERRE%20DE%20RELEVAMIENTO")

wait.until(EC.element_to_be_clickable(
    (By.XPATH, '//*[@id="btnBuscar"]'))).click()

wait.until(EC.presence_of_element_located(
    (By.XPATH, '/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table')))
time.sleep(1)

tabla = driver.find_element(
    by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table')
filas1 = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr'))
content = tabla.get_attribute("outerHTML")

for x in range(0, len(content)):
    if (content[x] == "<") and (content[x+1] == "t") and (content[x+2] == "d") and (content[x+3] == ">"):
        y = x+4
        cierre = content[y]+content[y+1]+content[y+2]+content[y+3]+content[y+4]
        while (cierre != "</td>"):
            campo += content[y]
            y += 1
            cierre = content[y]+content[y+1] + \
                content[y+2]+content[y+3]+content[y+4]
        # ------------ Casilla de verificacion ---------------------------
        if (columna == 0) and (campo[0] == chr(10)):
            campo = "V"
        # ------------ separar altura y localidad ---------------------------
        if (columna == 6) and (campo[0] == "<"):
            campo = campo[52:]
            comilla = campo.index("'")
            campo = campo[:comilla]
            guion = campo.index("-")
            altura = campo[:guion-1]
            vector.append(altura)
            guion = campo.rfind("-")
            campo = campo[guion+2:]
        # ------------ Subtipo ---------------------------
        if (columna == 7):
            if campo == "NORMALIZADO (ORE)":
                campo = "NORMALIZADO"
            if campo == "ARMADO (ORE)":
                campo = "ARMADO"
        # ------------ Bandeja previa y Observacion ---------------------------
        if (columna == 13) and (campo[0] == "<"):
            campo = campo[52:]
            comilla = campo.index("'")
            obs = campo[:comilla]
            flecha1 = campo.index(">")
            flecha2 = campo.index("<")
            campo = campo[flecha1+1:flecha2]
            if campo == "PENDIENTE DE RE...":
                campo = "PENDIENTE DE RELEVAMIENTO"
            if campo == "PENDIENTE DE DI...":
                campo = "PENDIENTE DE DISEÑO DE RED"
            if campo == "PLANIFICACION D...":
                campo = "PLANIFICACION DE TAREAS"
            if campo == "EN CERTIFICACIO...":
                campo = "EN CERTIFICACION"
            if campo == "ANALISIS DE FAC...":
                campo = "ANALISIS DE FACTIBILIDAD"
            vector.append(campo)
            campo = obs
        # ------------ inserta columna Obs Anterior ---------------------------
        if (columna == 14):
            vector.append(" ")
        # ------------ Cambiar el Nodo GPON ---------------------------
        if (columna == 14) and (campo != "&nbsp;"):
            # si el nodo GPON esta y el nodo HFC esta vacio
            if vector[len(vector) - 13] == " ":
                vector[len(vector) - 13] = campo
            # si el nodo GPON esta y el nodo HFC son solo numeros (ejemplo= 001100)
            if not re.search(r'[a-zA-Z]', vector[len(vector) - 13]):
                vector[len(vector) - 13] = campo
            # si el nodo GPON esta y el nodo HFC posee guiones (ejemplo= HF-SUR)
            if re.search('-', vector[len(vector) - 13]):
                vector[len(vector) - 13] = campo
        # ------------ Cambiar campos vacios ---------------------------
        if (campo == "&nbsp;"):
            campo = " "

        vector.append(campo)
        columna += 1
        if columna == 17:
            columna = 0

    campo = ""

# -------------------------------Bandeja de Reconversion Tecnologica -----
columna = 0

driver.get(
    "http://crm.telecentro.local/Edificio/Gt_Edificio/BandejaEntradaDeRelevamiento.aspx?TituloPantalla=CIERRE%20DE%20RELEVAMIENTO&EstadoGestionId=303&TipoGestionId=6&TipoGestion=RECONVERSION%20TECNOLOGICA%20-%20CIERRE%20DE%20RELEVAMIENTO")

wait.until(EC.presence_of_element_located(
    (By.XPATH, '/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table')))
time.sleep(1)

tabla2 = driver.find_element(
    by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table')
filas2 = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr'))
content2 = tabla2.get_attribute("outerHTML")

for x in range(0, len(content2)):
    if (content2[x] == "<") and (content2[x+1] == "t") and (content2[x+2] == "d") and ((content2[x+3] == ">") or (content2[x+3] == " ")):
        y = x+4
        cierre = content2[y]+content2[y+1] + \
            content2[y+2]+content2[y+3]+content2[y+4]
        while (cierre != "</td>"):
            campo += content2[y]
            y += 1
            cierre = content2[y]+content2[y+1] + \
                content2[y+2]+content2[y+3]+content2[y+4]
        # ------------ Casilla de verificacion ---------------------------
        if (columna == 0) and (campo[0] == chr(10)):
            campo = "V"
        # ------------ separar altura y localidad ---------------------------
        if (columna == 6) and (campo[0] == "<"):
            campo = campo[52:]
            comilla = campo.index("'")
            campo = campo[:comilla]
            guion = campo.index("-")
            altura = campo[:guion-1]
            vector.append(altura)
            guion = campo.rfind("-")
            campo = campo[guion+2:]
        # ------------ Subtipo ---------------------------
        if (columna == 7):
            if campo == "NORMALIZADO (ORE)":
                campo = "NORMALIZADO"
            if campo == "ARMADO (ORE)":
                campo = "ARMADO"
        # ------------ Saltar la columna contratista que no aplica para reconversion ---------------------------
        if (columna == 12):
            vector.append(" ")
        # ------------ Bandeja previa y Observacion ---------------------------
        if (columna == 12) and (campo[0] == "<"):
            campo = campo[52:]
            comilla = campo.index("'")
            obs = campo[:comilla]
            flecha1 = campo.index(">")
            flecha2 = campo.index("<")
            campo = campo[flecha1+1:flecha2]
            if campo == "PENDIENTE DE RE...":
                campo = "PENDIENTE DE RELEVAMIENTO"
            if campo == "PENDIENTE DE DI...":
                campo = "PENDIENTE DE DISEÑO DE RED"
            if campo == "PLANIFICACION D...":
                campo = "PLANIFICACION DE TAREAS"
            if campo == "EN CERTIFICACIO...":
                campo = "EN CERTIFICACION"
            if campo == "ANALISIS DE FAC...":
                campo = "ANALISIS DE FACTIBILIDAD"
            vector.append(campo)
            campo = obs
        # ------------ Observacion Anterior---------------------------
        if (columna == 13):
            campo = campo[7:]
            comilla = campo.index('"')
            campo = campo[:comilla]
        # ------------ Cambiar el Nodo GPON ---------------------------
        if (columna == 14) and (campo != "&nbsp;"):
            vector[len(vector) - 13] = campo
        # ------------ Cambiar campos vacios ---------------------------
        if (campo == "&nbsp;"):
            campo = " "
        vector.append(campo)
        # ------------ Agrega Campos Vacios para igualar columnas de Cierre ---------------------------
        if (columna == 14):
            vector.append(" ")
            vector.append(" ")

        columna += 1
        if columna == 15:
            columna = 0

    campo = ""


# ------------- filas en cierre + filas en reconversion + vector de metaDatos -----------------------
filasTotal = filas1+filas2+1

hora_actual = datetime.datetime.now()
hora_formateada = hora_actual.strftime('%H:%M:%S')

# ------------------------------- Bandeja de Pendiente de Relevamiento ----------
driver.get(
    "http://crm.telecentro.local/Edificio/Gt_Edificio/BandejaEntradaDeRelevamiento.aspx?TituloPantalla=Pendiente%20De%20Relevamiento&EstadoGestionId=4&TipoGestionId=3&TipoGestion=OPERACIONES%20DE%20RED%20-%20PENDIENTE%20DE%20RELEVAMIENTO")
wait.until(EC.presence_of_element_located(
    (By.XPATH, '//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr')))
time.sleep(1)
filas3 = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr'))
# ------------------------------- Bandeja de Analisis de Factibilidad ----------
driver.get(
    "http://crm.telecentro.local/Edificio/Gt_Edificio/BandejaEntradaDeRelevamiento.aspx?TituloPantalla=An%c3%a1lisis%20De%20Factibilidad&EstadoGestionId=6&TipoGestionId=3&TipoGestion=OPERACIONES%20DE%20RED%20-%20ANALISIS%20DE%20FACTIBILIDAD")
wait.until(EC.presence_of_element_located(
    (By.XPATH, '//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr')))
time.sleep(1)
filas4 = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr'))


# --- hora_formateada: Hora actural del sistema (Hora de ultima actualizacion) -----------------------
# --- filas1: filas en cierre de Relevamiento -----------------------
# --- filas2: filas en cierre de Reconversion Tecnologica -----------------------


metaData = ["mD", hora_formateada, filas1, filas2, filas3, filas4, "", "",
            "", "", "", "", "", "", "", "", "", "", "", ""]
vector = vector + metaData


serie = pd.Series(vector)
df = pd.DataFrame(serie.values.reshape(filasTotal, 20))
df.columns = ["N", "Gestion", "ID", "Nodo", "Zona", "Prioridad", "Direccion", "Localidad", "Subtipo", "Ult Visita", "Estado Edificio",
              "Cant Gestiones", "Usuario", "Contratista", "Bandeja Previa", "Observacion", "Observacion Anterior", "Nodo Gpon", "Dias Faltantes", "Dias Retraso"]
print(df)

driver.quit()

# ------------------------- Subir a Google Sheet -----------------------------

scope = ['https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive"]

credenciales = ServiceAccountCredentials.from_json_keyfile_name(
    "carga-de-bandeja-de-entrada-01ec277da545.json", scope)

cliente = gspread.authorize(credenciales)
# ------------- Crea y comparte la Google Sheet  -------------------------
#libro = cliente.create("AutocargaGestiones")
#libro.share("ignaciogproce3@gmail.com", perm_type="user", role="writer")
# ------------------------------------------------------------------------

hoja = cliente.open("AutocargaGestiones").sheet1

hoja.clear()

hoja.update([df.columns.values.tolist()] + df.values.tolist())


# ------------------------------------


"""



"""
