import matplotlib
matplotlib.use('Agg')  
import tkinter as tk
from tkinter import ttk, messagebox
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import pandas as pd
import threading
from PIL import Image, ImageTk
from urllib.parse import urljoin
import time
import logging
import os
import matplotlib.pyplot as plt  # Para generar gráficos
import numpy as np  # Para manejar los ticks del eje y
import unicodedata  # Para normalizar acentos
import re  # Para sanitizar nombres de hojas

# Configurar logging para depuración detallada
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Clase principal de la aplicación
class QureoApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Script de revisión de avance - Múltiples Colegios")
        self.master.geometry("400x250")
        self.master.configure(bg="#003366")

        self.boton_iniciar = tk.Button(master, text="Iniciar proceso con credenciales XLSX", bg="#3399FF", fg="white", command=self.iniciar_proceso)
        self.boton_iniciar.pack(pady=10)

        self.progress = ttk.Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        self.estado = tk.Label(master, text="", bg="#003366", fg="white")
        self.estado.pack()

    def iniciar_proceso(self):
        if not os.path.exists("credenciales_colegios.xlsx"):
            messagebox.showerror("Error", "No se encontró 'credenciales_colegios.xlsx'. Crea el archivo con columnas: Colegio,Usuario,Contraseña.")
            return

        self.boton_iniciar.config(state="disabled")
        self.estado.config(text="Iniciando proceso...")
        threading.Thread(target=self.procesar_colegios).start()

    def update_gui(self, text):
        """Función para actualizar la GUI de forma segura desde un hilo."""
        self.master.after(0, lambda: self.estado.config(text=text))

    def show_error(self, title, message):
        """Función para mostrar un mensaje de error de forma segura desde un hilo."""
        self.master.after(0, lambda: messagebox.showerror(title, message))

    def truncate_sheet_name(self, name, suffix=""):
        """Trunca el nombre de la hoja a 31 caracteres, considerando el sufijo, elimina caracteres inválidos y asegura unicidad."""
        name = str(name).strip()
        name = re.sub(r'[:\\\/?*\[\]]', '', name)  # Eliminar :, \, /, ?, *, [, ]
        max_length = 31 - len(suffix)
        if max_length < 1:
            max_length = 31
        truncated = name[:max_length].strip()
        if not truncated or truncated.lower() == "nan":
            truncated = "Sheet_" + str(hash(name))[:8]
        full_name = f"{truncated}{suffix}"[:31].strip()
        return full_name

    def normalize_text(self, text):
        """Normaliza texto eliminando acentos y convirtiendo a mayúsculas."""
        text = ''.join(c for c in unicodedata.normalize('NFKD', text) if unicodedata.category(c) != 'Mn')
        return text.strip().upper()

    def procesar_colegios(self):
        start_time = time.time()
        BASE_URL = "https://sa-admin.qureo.education"
        datos_por_colegio = {}
        estudiantes_omitidos_global = []
        colegios_con_errores = []

        colegios_especiales = [
            "CARLOS PHILLIPS",
            "19 DE JUNIO",
            "8 DE DICIEMBRE",
            "JOSÉ BAQUIJANO Y CARRILLO"
        ]
        colegios_especiales_normalized = [self.normalize_text(c) for c in colegios_especiales]

        grupo_mapping = {
            "CARLOS PHILLIPS": "GRUPO 1",
            "JOSÉ BAQUIJANO Y CARRILLO": "GRUPO 2",
            "19 DE JUNIO": "GRUPO 3",
            "8 DE DICIEMBRE": "GRUPO 4"
        }

        try:
            credenciales_df = pd.read_excel("credenciales_colegios.xlsx", sheet_name=0)
            required_columns = ["Colegio", "Usuario", "Contraseña"]
            if not all(col in credenciales_df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in credenciales_df.columns]
                self.show_error("Error", f"El archivo XLSX no contiene las columnas requeridas: {', '.join(missing_cols)}")
                self.update_gui("Error en el archivo XLSX")
                self.master.after(0, lambda: self.boton_iniciar.config(state="normal"))
                return

            if credenciales_df[required_columns].isna().any().any():
                self.show_error("Error", "El archivo 'credenciales_colegios.xlsx' contiene valores vacíos o NaN en las columnas Colegio, Usuario o Contraseña.")
                self.update_gui("Error: Valores inválidos en el archivo XLSX")
                self.master.after(0, lambda: self.boton_iniciar.config(state="normal"))
                return

            credenciales_df = credenciales_df.dropna(subset=required_columns)
            total_colegios = len(credenciales_df)
            if total_colegios == 0:
                self.show_error("Error", "No hay colegios válidos en 'credenciales_colegios.xlsx'.")
                self.update_gui("Error: No hay colegios válidos")
                self.master.after(0, lambda: self.boton_iniciar.config(state="normal"))
                return

            logger.info(f"Procesando {total_colegios} colegios.")

            reporte_anterior_path = "reporte_anterior.xlsx"
            df_anterior = {}
            if os.path.exists(reporte_anterior_path):
                try:
                    df_anterior = pd.read_excel(reporte_anterior_path, sheet_name=None)
                except Exception as e:
                    logger.warning(f"Error al leer reporte_anterior.xlsx: {str(e)}. Se ignorará.")
                    os.remove(reporte_anterior_path)

            self.progress["maximum"] = total_colegios * 100
            self.progress["value"] = 0

            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)

                for idx, row in credenciales_df.iterrows():
                    colegio = row["Colegio"]
                    usuario = row["Usuario"]
                    contrasena = row["Contraseña"]
                    logger.info(f"Procesando colegio: {colegio}")

                    colegio_normalized = self.normalize_text(colegio)
                    datos = []
                    estudiantes_omitidos = []
                    context = browser.new_context(viewport={"width": 1920, "height": 1080}, no_viewport=False)
                    page = context.new_page()

                    try:
                        page.goto("https://sa-admin.qureo.education/login")
                        page.wait_for_selector("input[name='userId']", timeout=20000)
                        page.fill("input[name='userId']", str(usuario))
                        page.fill("input[name='userPassword']", str(contrasena))
                        page.click("button[type='submit']")

                        try:
                            page.wait_for_load_state("networkidle", timeout=20000)
                            if "login" in page.url:
                                raise Exception("Fallo en el inicio de sesión.")
                        except PlaywrightTimeoutError:
                            raise Exception("No se pudo cargar la página después del login.")

                        if colegio_normalized in colegios_especiales_normalized:
                            logger.info(f"Colegio especial {colegio}: Lista de estudiantes ya visible.")
                            page.wait_for_load_state("networkidle", timeout=30000)
                            try:
                                page.wait_for_selector("a[href*='/students/']", timeout=10000)
                            except PlaywrightTimeoutError:
                                logger.error(f"No se encontraron enlaces de estudiantes en {colegio}.")
                                raise Exception("No se encontraron enlaces de estudiantes.")
                        else:
                            try:
                                estudiante_link = None
                                try:
                                    estudiante_link = page.wait_for_selector("a[href='/schoolinfo/students']", timeout=20000)
                                except PlaywrightTimeoutError:
                                    logger.warning(f"Enlace 'a[href=/schoolinfo/students]' no encontrado en {colegio}.")
                                    try:
                                        estudiante_link = page.wait_for_selector("a:has-text('Estudiantes')", timeout=10000)
                                    except PlaywrightTimeoutError:
                                        try:
                                            estudiante_link = page.wait_for_selector("a:has-text('Students')", timeout=10000)
                                        except PlaywrightTimeoutError:
                                            logger.error(f"No se encontró enlace a estudiantes en {colegio}.")
                                            raise Exception("No se encontró el enlace a estudiantes.")

                                if estudiante_link:
                                    estudiante_link.click()
                                    page.wait_for_load_state("networkidle", timeout=20000)
                                else:
                                    raise Exception("No se encontró enlace a estudiantes.")
                            except PlaywrightTimeoutError:
                                logger.error(f"Timeout al buscar enlace a estudiantes en {colegio}.")
                                raise Exception("No se encontró el enlace a estudiantes.")

                        total_estudiantes = 0
                        estudiantes_vistos = set()
                        estudiantes_data = []

                        while True:
                            estudiantes = page.query_selector_all("a[href*='/students/']")
                            logger.info(f"Encontrados {len(estudiantes)} enlaces de estudiantes en esta página para {colegio}")
                            for e in estudiantes:
                                text = e.text_content().strip()
                                if text.lower() != "añadir estudiante" and text not in estudiantes_vistos:
                                    estudiantes_vistos.add(text)
                                    href = e.get_attribute("href")
                                    if href and not href.startswith("log:"):
                                        full_url = urljoin(BASE_URL, href)
                                        row = e.query_selector("xpath=ancestor::tr")
                                        aula_cell = row.query_selector("td:nth-child(2)") if row else None
                                        if colegio_normalized in colegios_especiales_normalized:
                                            nombre_aula = grupo_mapping.get(colegio, "Desconocida")
                                        else:
                                            nombre_aula = aula_cell.text_content().strip() if aula_cell else "Desconocida"
                                        logger.info(f"Estudiante: {text}, Aula: {nombre_aula} en {colegio}")
                                        estudiantes_data.append((nombre_aula, text, full_url))
                                        total_estudiantes += 1
                            logger.info(f"Total de estudiantes contados hasta ahora en {colegio}: {total_estudiantes}")

                            try:
                                siguiente_boton = page.query_selector("button[aria-label*='next page']")
                                if siguiente_boton and "Mui-disabled" in (siguiente_boton.get_attribute("class") or ""):
                                    break
                                if siguiente_boton:
                                    siguiente_boton.click()
                                    page.wait_for_load_state("networkidle", timeout=15000)
                                    page.wait_for_selector("a[href*='/students/']", timeout=15000)
                                else:
                                    break
                            except PlaywrightTimeoutError:
                                logger.warning(f"No se pudo avanzar a la siguiente página en {colegio}")
                                break

                        self.progress["maximum"] = self.progress["maximum"] + total_estudiantes - 100

                        # Procesa los datos de cada estudiante (optimizado)
                        for nombre_aula, nombre, url in estudiantes_data:
                            if "-" in nombre_aula:
                                try:
                                    grado, seccion = nombre_aula.split("-", 1)[:2]
                                except ValueError:
                                    logger.warning(f"Formato de aula inválido: {nombre_aula} en {colegio}")
                                    grado, seccion = "Desconocido", "Desconocida"
                            else:
                                logger.warning(f"Aula sin guion: {nombre_aula} en {colegio}")
                                grado, seccion = "Desconocido", "Desconocida"

                            new_page = context.new_page()
                            success = False
                            for attempt in range(2):  # Reducir a 2 intentos
                                try:
                                    logger.info(f"Intento {attempt + 1} para estudiante {nombre} en aula {nombre_aula} ({colegio})")
                                    new_page.goto(url, timeout=12000)  # Reducir timeout
                                    new_page.wait_for_load_state("domcontentloaded", timeout=12000)  # Usar domcontentloaded

                                    # Esperar explícitamente a que los acordeones estén presentes
                                    try:
                                        new_page.wait_for_selector(
                                            "//div[contains(@class,'MuiAccordionSummary-root') and .//h3]", 
                                            state="visible", 
                                            timeout=8000  # Reducir timeout
                                        )
                                    except PlaywrightTimeoutError:
                                        logger.warning(f"No se encontraron acordeones válidos para {nombre} en aula {nombre_aula} ({colegio})")
                                        raise Exception("No se encontraron acordeones válidos")

                                    # Obtener acordeones y filtrar títulos relevantes de una vez
                                    acordeones = new_page.query_selector_all("//div[contains(@class,'MuiAccordionSummary-root') and .//h3]")
                                    logger.info(f"Estudiante {nombre} en aula {nombre_aula} ({colegio}): Encontrados {len(acordeones)} acordeones")

                                    # Cachear títulos de acordeones para evitar consultas repetidas
                                    titulos_acordeones = []
                                    for i, acordeon in enumerate(acordeones, 1):
                                        h3_element = acordeon.query_selector("xpath=.//h3")
                                        titulo = h3_element.text_content().strip() if h3_element else "Curso sin título"
                                        titulos_acordeones.append((acordeon, titulo))
                                        logger.info(f"Acordeón {i}: {titulo}")

                                    cursos_encontrados = False
                                    cursos_validos = []
                                    for i, (acordeon, titulo) in enumerate(titulos_acordeones, 1):
                                        # Filtrar cursos no relevantes primero
                                        if not any(keyword in titulo.lower() for keyword in ["principiante", "javascript", "beginner", "js", "básico", "intro"]):
                                            logger.info(f"Curso {titulo} descartado para {nombre} en aula {nombre_aula} ({colegio})")
                                            continue

                                        # Evitar procesar cursos duplicados
                                        if titulo in cursos_validos:
                                            logger.warning(f"Curso {titulo} ya procesado para {nombre} en aula {nombre_aula} ({colegio}), omitiendo")
                                            continue

                                        try:
                                            # Verificar visibilidad
                                            if not acordeon.is_visible():
                                                logger.warning(f"Acordeón {i} no visible para {nombre} en aula {nombre_aula} ({colegio})")
                                                continue

                                            # Expandir acordeón si está colapsado
                                            if acordeon.get_attribute("aria-expanded") == "false":
                                                try:
                                                    acordeon.scroll_into_view_if_needed(timeout=4000)
                                                    acordeon.click(timeout=4000)
                                                    new_page.wait_for_timeout(700)  # Reducir espera
                                                except Exception as e:
                                                    logger.warning(f"Error al expandir acordeón {titulo} para {nombre} en aula {nombre_aula} ({colegio}): {str(e)}")
                                                    continue

                                            # Obtener contenedor y progreso
                                            contenedor = acordeon.query_selector("xpath=./ancestor::div[contains(@class, 'MuiAccordion-root')]")
                                            if not contenedor:
                                                logger.warning(f"No se encontró contenedor para {titulo}")
                                                progreso_texto = "0"
                                            else:
                                                progreso = contenedor.query_selector(
                                                    "xpath=.//div[contains(text(),'finalizados') or contains(text(),'completed')]/following-sibling::div"
                                                )
                                                progreso_texto = progreso.text_content().strip() if progreso else "0"
                                                logger.info(f"Progreso en {titulo}: {progreso_texto}")

                                            datos.append({
                                                "Aula": nombre_aula,
                                                "Grado": grado,
                                                "Sección": seccion,
                                                "Estudiante": nombre,
                                                "Curso": titulo,
                                                "Capítulos finalizados": progreso_texto
                                            })
                                            cursos_validos.append(titulo)
                                            cursos_encontrados = True

                                        except Exception as e:
                                            logger.error(f"Error al procesar curso {titulo}: {str(e)}")
                                            continue

                                    # Registrar cursos faltantes
                                    if not cursos_encontrados or len(cursos_validos) < 2:
                                        logger.info(f"Sin cursos válidos suficientes, registrando ambos cursos")
                                        for curso in ["Curso para principiantes", "Curso de JavaScript"]:
                                            if curso not in cursos_validos:
                                                datos.append({
                                                    "Aula": nombre_aula,
                                                    "Grado": grado,
                                                    "Sección": seccion,
                                                    "Estudiante": nombre,
                                                    "Curso": curso,
                                                    "Capítulos finalizados": "0"
                                                })
                                                logger.info(f"Registrado {curso} con 0 capítulos finalizados")
                                        cursos_encontrados = True

                                    if cursos_encontrados:
                                        success = True
                                        break
                                    else:
                                        logger.warning(f"No se encontraron cursos válidos, reintentando...")
                                        time.sleep(1)  # Reducir espera

                                except Exception as e:
                                    logger.error(f"Intento {attempt + 1} fallido para estudiante {nombre} en aula {nombre_aula} ({colegio}): {str(e)}")
                                    if attempt < 1:  # Solo reintentar una vez
                                        time.sleep(1)
                                    continue

                            if not success:
                                logger.error(f"Fallo tras reintentos, estudiante {nombre} omitido")
                                estudiantes_omitidos.append(nombre)
                                datos.append({
                                    "Aula": nombre_aula,
                                    "Grado": grado,
                                    "Sección": seccion,
                                    "Estudiante": nombre,
                                    "Curso": "Error",
                                    "Capítulos finalizados": "0"
                                })

                            new_page.close()
                            self.master.after(0, lambda: self.progress.__setitem__("value", self.progress["value"] + 1))
                            if self.progress["value"] % 10 == 0:
                                self.update_gui(f"Procesado {self.progress['value']} estudiantes (colegio {idx+1}/{total_colegios})")

                        datos_por_colegio[colegio] = pd.DataFrame(datos)
                        estudiantes_omitidos_global.extend([f"{nombre} ({colegio})" for nombre in estudiantes_omitidos])

                    except Exception as e:
                        logger.error(f"Error general en {colegio}: {str(e)}")
                        self.show_error("Error", f"Error en {colegio}: {str(e)}")
                        colegios_con_errores.append(colegio)
                    finally:
                        context.close()

                browser.close()

            if datos_por_colegio:
                try:
                    with pd.ExcelWriter("reporte_avances.xlsx", engine='xlsxwriter') as writer_avances:
                        for colegio, df_actual in datos_por_colegio.items():
                            sheet_name = self.truncate_sheet_name(colegio)
                            def parse_capitulos(caps):
                                try:
                                    if "/" in str(caps):
                                        completados, total = map(int, caps.split("/"))
                                        return completados, total
                                    return 0, 0
                                except:
                                    return 0, 0

                            df_actual[["Capítulos completados", "Total capítulos"]] = df_actual["Capítulos finalizados"].apply(parse_capitulos).apply(pd.Series)
                            if colegio in df_anterior:
                                df_prev = df_anterior.get(colegio, pd.DataFrame())
                                df_prev[["Capítulos completados_prev", "Total capítulos_prev"]] = df_prev["Capítulos finalizados"].apply(parse_capitulos).apply(pd.Series)
                                df_merged = df_actual.merge(df_prev, on=["Aula", "Grado", "Sección", "Estudiante", "Curso"], suffixes=("_actual", "_prev"), how="left")
                                df_merged["Capítulos completados_prev"] = df_merged["Capítulos completados_prev"].fillna(0)
                                df_merged["Total capítulos_prev"] = df_merged["Total capítulos_prev"].fillna(0)
                                df_merged["Avance"] = df_merged.apply(lambda row: "Sí" if row["Capítulos completados_actual"] > row["Capítulos completados_prev"] else "No", axis=1)
                                df_avance = df_merged[["Aula", "Grado", "Sección", "Estudiante", "Curso", "Capítulos completados_actual", "Total capítulos_actual", "Capítulos completados_prev", "Total capítulos_prev", "Avance"]]
                                df_avance.to_excel(writer_avances, sheet_name=sheet_name, index=False)
                            else:
                                df_actual["Avance"] = "N/A (primera ejecución)"
                                df_avance = df_actual[["Aula", "Grado", "Sección", "Estudiante", "Curso", "Capítulos completados", "Total capítulos", "Avance"]]
                                df_avance.to_excel(writer_avances, sheet_name=sheet_name, index=False)

                            plataformas = df_actual["Curso"].unique()
                            for plataforma in plataformas:
                                if "principiante" in plataforma.lower() or "beginner" in plataforma.lower() or "básico" in plataforma.lower() or "intro" in plataforma.lower():
                                    plat_name = "Qureo"
                                elif "javascript" in plataforma.lower() or "js" in plataforma.lower():
                                    plat_name = "Curso de JavaScript"
                                else:
                                    continue

                                df_plataforma = df_actual[df_actual["Curso"] == plataforma]
                                if df_plataforma.empty:
                                    continue

                                resumen_por_aula = df_plataforma.groupby("Aula").agg(
                                    Total_Estudiantes=("Estudiante", "count"),
                                    Avance_Promedio=("Capítulos completados", "mean"),
                                    Capítulo_Más_Común=("Capítulos completados", lambda x: x.mode()[0] if not x.mode().empty else 0),
                                    Total_Capítulos=("Total capítulos", "max")
                                ).reset_index()
                                resumen_por_aula["Mensaje_Mayoría"] = resumen_por_aula.apply(
                                    lambda row: f"La mayoría de estudiantes está por el capítulo {int(row['Capítulo_Más_Común'])} de {int(row['Total_Capítulos'])}", axis=1
                                )
                                resumen_sheet_name = self.truncate_sheet_name(colegio, f"_Resumen_{plat_name}")
                                resumen_por_aula.to_excel(writer_avances, sheet_name=resumen_sheet_name, index=False)
                                logger.info(f"Resumen por aula generado para {colegio} ({plat_name}): {resumen_por_aula}")

                    os.makedirs("graficos", exist_ok=True)
                    colegios_list = list(datos_por_colegio.keys())
                    avances_globales_qureo = [0] * len(colegios_list)
                    avances_globales_js = [0] * len(colegios_list)

                    for idx, (colegio, df_actual) in enumerate(datos_por_colegio.items()):
                        df_actual["Capítulos completados"] = pd.to_numeric(df_actual["Capítulos completados"], errors='coerce').fillna(0)
                        plataformas = df_actual["Curso"].unique()

                        for plataforma in plataformas:
                            if "principiante" in plataforma.lower() or "beginner" in plataforma.lower() or "básico" in plataforma.lower() or "intro" in plataforma.lower():
                                plat_name = "Qureo"
                            elif "javascript" in plataforma.lower() or "js" in plataforma.lower():
                                plat_name = "Curso de JavaScript"
                            else:
                                continue

                            df_plataforma = df_actual[df_actual["Curso"] == plataforma]
                            if df_plataforma.empty:
                                continue

                            resumen_por_aula = df_plataforma.groupby("Aula")["Capítulos completados"].mean().reset_index()
                            avance_promedio = df_plataforma["Capítulos completados"].mean()
                            if plat_name == "Qureo":
                                avances_globales_qureo[idx] = avance_promedio
                            else:
                                avances_globales_js[idx] = avance_promedio

                            plt.figure(figsize=(10, 6))
                            plt.bar(resumen_por_aula["Aula"], resumen_por_aula["Capítulos completados"])
                            plt.title(f"Avance Promedio por Aula en {colegio} ({plat_name})")
                            plt.xlabel("Aula")
                            plt.ylabel("Capítulos Completados (Promedio)")
                            plt.xticks(rotation=45)
                            plt.tight_layout()
                            grafico_path = f"graficos/avance_{self.truncate_sheet_name(colegio)}_{self.truncate_sheet_name(plat_name)}.png"
                            plt.savefig(grafico_path)
                            plt.close()
                            logger.info(f"Gráfico generado para {colegio} ({plat_name}): {grafico_path}")

                    plt.figure(figsize=(12, 6))
                    x = range(len(colegios_list))
                    plt.bar([i - 0.2 for i in x], avances_globales_qureo, width=0.4, label="Qureo", color="blue")
                    plt.bar([i + 0.2 for i in x], avances_globales_js, width=0.4, label="Curso de JavaScript", color="purple")
                    plt.title("Avance Promedio por Colegio (Global)")
                    plt.xlabel("Colegio")
                    plt.ylabel("Capítulos Completados (Promedio)")
                    plt.xticks(x, colegios_list, rotation=90)
                    plt.yticks(np.arange(0, max(max(avances_globales_qureo, default=0), max(avances_globales_js, default=0)) + 1, 1))
                    plt.legend()
                    plt.tight_layout()
                    grafico_global_path = "graficos/avance_global.png"
                    plt.savefig(grafico_global_path)
                    plt.close()
                    logger.info(f"Gráfico global generado: {grafico_global_path}")

                    self.update_gui("¡Proceso completado! Reporte de avances y gráficos generados.")
                    self.master.after(0, lambda: messagebox.showinfo("Éxito", f"Avances exportados a 'reporte_avances.xlsx'. Gráficos en carpeta 'graficos/'. Errores en: {', '.join(colegios_con_errores) if colegios_con_errores else 'Ningún colegio'}"))

                except Exception as e:
                    logger.error(f"Error al generar reporte o gráficos: {str(e)}")
                    self.show_error("Error", f"Error al generar reporte o gráficos: {str(e)}.")
                    self.update_gui("Error al generar reporte o gráficos")
                    self.master.after(0, lambda: self.boton_iniciar.config(state="normal"))
                    return

            else:
                self.update_gui("Error: No se procesó ningún colegio correctamente.")
                self.show_error("Error", "No se pudo procesar ningún colegio.")
                self.master.after(0, lambda: self.boton_iniciar.config(state="normal"))

            if estudiantes_omitidos_global:
                logger.error(f"Estudiantes omitidos: {estudiantes_omitidos_global}")
                self.show_error("Advertencia", f"Se omitieron {len(estudiantes_omitidos_global)} estudiantes: {', '.join(estudiantes_omitidos_global)}")

            logger.info(f"Tiempo total: {time.time() - start_time} segundos")
            self.master.after(0, lambda: self.boton_iniciar.config(state="normal"))

        except Exception as e:
            logger.error(f"Error al leer el archivo XLSX: {str(e)}")
            self.show_error("Error", f"Error al leer 'credenciales_colegios.xlsx': {str(e)}.")
            self.update_gui("Error en el archivo XLSX")
            self.master.after(0, lambda: self.boton_iniciar.config(state="normal"))

# Función para mostrar una pantalla de carga
def mostrar_splash(callback):
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.geometry("600x400+500+200")

    try:
        imagen = Image.open("qureo_1.jpg")
        imagen = imagen.resize((600, 400))
        foto = ImageTk.PhotoImage(imagen)
        label = tk.Label(splash, image=foto)
        label.image = foto
        label.pack()
    except Exception as e:
        logger.error(f"No se pudo cargar la imagen: {e}")
        splash.destroy()
        callback()
        return

    def cerrar_splash():
        splash.destroy()
        callback()

    splash.after(3000, cerrar_splash)
    splash.mainloop()

# Punto de entrada del programa
if __name__ == "__main__":
    def iniciar_aplicacion():
        root = tk.Tk()
        app = QureoApp(root)
        root.mainloop()

    mostrar_splash(iniciar_aplicacion)