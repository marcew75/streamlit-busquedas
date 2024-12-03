from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
from selenium.webdriver.chrome.service import Service as ChromeService
from openpyxl import load_workbook
# Configuración del WebDriver
options = Options()
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36")
options.add_argument("--disable-search-engine-choice-screen")  # Desactiva la selección del motor de búsqueda
options.add_argument("--disable-notifications")               # Desactiva notificaciones emergentes
options.add_argument("--disable-popup-blocking")              # Desactiva bloqueos de pop-ups
options.add_argument("--start-maximized")                     # Inicia el navegador maximizado
options.add_argument("--disable-infobars")                    # Oculta barras informativas
options.add_argument("--headless=new")                        # Ejecuta el navegador en modo sin cabeza (opcional)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

# URL de inicio
url = "https://pe.wiautomation.com/"
driver.get(url)

# Espera explícita
wait = WebDriverWait(driver, 10)

# Manejar la ventana de cookies
try:
    botones_cookies = [
        "button#accept-cookies",       # Selector 1
        "div.cookie-consent button",  # Selector 2
        "div[class*='cookie'] button" # Selector 3
        "#usercentrics-root"
    ]
    for selector in botones_cookies:
        try:
            boton_cookies = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
            boton_cookies.click()
            print("Ventana de cookies cerrada.")
            break
        except Exception as e:
            print(f"Intento fallido con el selector {selector}: {e}")
    else:
        print("No se encontró ninguna ventana de cookies.")
except Exception as e:
    print(f"Error general al manejar las cookies: {e}")

# Seleccionar la marca
marca = "ABB"  # Cambia este valor según la marca deseada
try:
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, marca))).click()
    print(f"Marca seleccionada: {marca}")
except Exception as e:
    print(f"No se pudo seleccionar la marca: {marca}. Error: {e}")

# Seleccionar la categoría
categoria = "Fuente-de-alimentación"  # Cambia este valor según la categoría deseada
try:
    wait.until(EC.element_to_be_clickable((By.XPATH, f"//h3[text()='{categoria}']"))).click()
    print(f"Categoría seleccionada: {categoria}")
except Exception as e:
    print(f"No se pudo seleccionar la categoría: {categoria}. Error: {e}")

# Scroll infinito para cargar todos los productos
last_height = driver.execute_script("return document.body.scrollHeight")
while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)  # Espera para cargar más contenido
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height

# Extraer datos de los productos
productos = []
contenedores = driver.find_elements(By.CSS_SELECTOR, "div.item_content")  # Ajusta este selector según el HTML

if not contenedores:
    print("No se encontraron productos.")
else:
    for contenedor in contenedores:
        try:
            # Extraer datos usando selectores de Selenium
            descripcion = contenedor.find_element(By.CSS_SELECTOR, "span.name").text
            imagen_url = contenedor.find_element(By.CSS_SELECTOR, "img.image.lozad.loaded").get_attribute("src")
            precio = contenedor.find_element(By.CSS_SELECTOR, "div.price").text #if contenedor.find_elements(By.CSS_SELECTOR, "div.stock") else "No disponible"

            # Guardar datos en la lista
            productos.append({
                "Descripción": descripcion,
                "URL Imagen": imagen_url,
                "Categoría": categoria,
                "precio": precio
            })
        except Exception as e:
            print(f"Error al procesar un producto: {e}")
            continue

# Cerrar el navegador
driver.quit()

# Guardar datos en un archivo Excel
if productos:
    df = pd.DataFrame(productos)
    
    df.to_excel("productos_con_precio.xlsx", index=False)
    print("Archivo productos.xlsx creado con éxito.")
else:
    print("No se encontraron productos.")
