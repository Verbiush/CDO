import time
import pandas as pd
import json
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import sys
import threading

# Variable global para mantener la instancia del navegador
DRIVER_INSTANCE = None
DRIVER_LOCK = threading.Lock()

# Lista de pasos memorizados
# Estructura de cada paso:
# {
#   "accion": "escribir" | "click" | "tecla" | "espera",
#   "xpath": "...", (solo para escribir/click)
#   "columna": "...", (solo para escribir)
#   "tecla": "...", (solo para tecla: ENTER, TAB, etc.)
#   "tiempo": 1.0 (solo para espera)
#   "descripcion": "Texto legible"
# }
PASOS_MEMORIZADOS = []
PASOS_ALTERNATIVOS = [] # Legacy, mantenido por compatibilidad pero se prefiere FLUJOS_CONDICIONALES
CONDICION_EJECUCION = None # Legacy
FLUJOS_CONDICIONALES = {} # Diccionario {indice: {pasos, condicion, nombre}}

# Control de Ejecución
EJECUCION_ACTIVA = False
EJECUCION_LOCK = threading.Lock()
ULTIMO_ERROR = None

def detener_ejecucion():
    global EJECUCION_ACTIVA
    with EJECUCION_LOCK:
        EJECUCION_ACTIVA = False
    return True, "Solicitud de detención enviada."

def cargar_pasos_alternativos(nuevos_pasos):
    # Wrapper legacy para cargar en el slot 0
    return set_flujo_condicional(0, nuevos_pasos, CONDICION_EJECUCION, "Alternativo 1")

def set_flujo_condicional(index, pasos, condicion, nombre="Alternativo"):
    global FLUJOS_CONDICIONALES, PASOS_ALTERNATIVOS, CONDICION_EJECUCION
    if isinstance(pasos, list):
        FLUJOS_CONDICIONALES[index] = {
            "pasos": pasos,
            "condicion": condicion,
            "nombre": nombre
        }
        # Sync Legacy vars si es indice 0
        if index == 0:
            PASOS_ALTERNATIVOS = pasos
            CONDICION_EJECUCION = condicion
            
        guardar_sesion()
        return True, f"Flujo '{nombre}' configurado correctamente ({len(pasos)} pasos)."
    return False, "Formato de pasos inválido."

def update_flujo_condicional(index, pasos=None, condicion=None, nombre=None):
    global FLUJOS_CONDICIONALES
    if index not in FLUJOS_CONDICIONALES:
        if pasos is None: return False, "Debe cargar pasos primero."
        FLUJOS_CONDICIONALES[index] = {"pasos": [], "condicion": None, "nombre": f"Alternativo {index+1}"}
    
    if pasos is not None:
        FLUJOS_CONDICIONALES[index]["pasos"] = pasos
        if index == 0: 
            global PASOS_ALTERNATIVOS
            PASOS_ALTERNATIVOS = pasos
            
    if condicion is not None:
        FLUJOS_CONDICIONALES[index]["condicion"] = condicion
        if index == 0:
            global CONDICION_EJECUCION
            CONDICION_EJECUCION = condicion
            
    if nombre is not None:
        FLUJOS_CONDICIONALES[index]["nombre"] = nombre
        
    guardar_sesion()
    return True, "Actualizado."

def get_flujo_condicional(index):
    return FLUJOS_CONDICIONALES.get(index)


def limpiar_pasos_alternativos():
    global PASOS_ALTERNATIVOS, CONDICION_EJECUCION, FLUJOS_CONDICIONALES
    PASOS_ALTERNATIVOS = []
    CONDICION_EJECUCION = None
    FLUJOS_CONDICIONALES = {}
    guardar_sesion()
    return True, "Todos los flujos alternativos limpiados."

def set_condicion_ejecucion(tipo, valor, columna=None):
    # Legacy wrapper para slot 0
    global CONDICION_EJECUCION, FLUJOS_CONDICIONALES
    
    cond = None
    if valor:
        cond = {"tipo": tipo, "valor": valor, "columna": columna}
    
    CONDICION_EJECUCION = cond
    
    # Actualizar slot 0 si existe
    if 0 in FLUJOS_CONDICIONALES:
        FLUJOS_CONDICIONALES[0]["condicion"] = cond

def obtener_driver(create_if_missing=True):
    global DRIVER_INSTANCE
    with DRIVER_LOCK:
        if DRIVER_INSTANCE is not None:
            try:
                # Verificar si sigue vivo
                DRIVER_INSTANCE.title
                return DRIVER_INSTANCE
            except UnexpectedAlertPresentException:
                # Si hay una alerta, el navegador está vivo pero bloqueado.
                # Retornamos la instancia para poder manejar la alerta.
                return DRIVER_INSTANCE
            except:
                DRIVER_INSTANCE = None
        
        if not create_if_missing:
            return None

        # Iniciar nuevo si no existe
        print("Iniciando nuevo driver...")
        try:
            service = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            # options.add_argument("--detach=true") 
            DRIVER_INSTANCE = webdriver.Chrome(service=service, options=options)
        except Exception as e_chrome:
            print(f"Chrome no encontrado o error: {e_chrome}. Intentando Edge...")
            try:
                service = Service(EdgeChromiumDriverManager().install())
                options = webdriver.EdgeOptions()
                options.add_argument("--start-maximized")
                DRIVER_INSTANCE = webdriver.Edge(service=service, options=options)
            except Exception as e_edge:
                print(f"Error iniciando Edge: {e_edge}")
                return None
        return DRIVER_INSTANCE

def abrir_navegador_inicial():
    """Solo abre el navegador y va al login. No bloquea."""
    driver = obtener_driver()
    if not driver:
        return False, "No se pudo iniciar el navegador (Chrome/Edge)."
    
    try:
        url = "https://ovidazs.siesacloud.com/ZeusSalud/ips/iniciando.php"
        driver.get(url)
        return True, "Navegador abierto. Por favor Inicie Sesión y navegue hasta la pantalla donde desea ingresar los datos."
    except Exception as e:
        return False, f"Error al navegar: {e}"

def limpiar_pasos():
    global PASOS_MEMORIZADOS
    PASOS_MEMORIZADOS = []
    guardar_sesion()
    return True, "Pasos limpiados."

def eliminar_ultimo_paso():
    global PASOS_MEMORIZADOS
    if PASOS_MEMORIZADOS:
        eliminado = PASOS_MEMORIZADOS.pop()
        guardar_sesion()
        return True, f"Eliminado paso: {eliminado['descripcion']}"
    return False, "No hay pasos para eliminar."

def eliminar_paso_indice(indice):
    global PASOS_MEMORIZADOS
    if 0 <= indice < len(PASOS_MEMORIZADOS):
        eliminado = PASOS_MEMORIZADOS.pop(indice)
        guardar_sesion()
        return True, f"Paso {indice+1} eliminado."
    return False, "Índice fuera de rango."

def mover_paso(indice_origen, direccion):
    """direccion: -1 (arriba), 1 (abajo)"""
    global PASOS_MEMORIZADOS
    if not PASOS_MEMORIZADOS: return False, "No hay pasos."
    
    nuevo_indice = indice_origen + direccion
    if 0 <= nuevo_indice < len(PASOS_MEMORIZADOS):
        PASOS_MEMORIZADOS[indice_origen], PASOS_MEMORIZADOS[nuevo_indice] = PASOS_MEMORIZADOS[nuevo_indice], PASOS_MEMORIZADOS[indice_origen]
        guardar_sesion()
        return True, "Paso movido."
    return False, "No se puede mover más allá de los límites."

def alternar_opcional_paso(indice):
    """Alterna el estado 'opcional' de un paso."""
    global PASOS_MEMORIZADOS
    if 0 <= indice < len(PASOS_MEMORIZADOS):
        paso = PASOS_MEMORIZADOS[indice]
        nuevo_estado = not paso.get("opcional", False)
        paso["opcional"] = nuevo_estado
        estado_str = "Opcional (Si falla, continúa)" if nuevo_estado else "Obligatorio (Si falla, se detiene)"
        guardar_sesion()
        return True, f"Paso {indice+1} ahora es: {estado_str}"
    return False, "Índice fuera de rango."

def _insertar_paso(paso, indice=None):
    global PASOS_MEMORIZADOS
    if indice is not None and 0 <= indice <= len(PASOS_MEMORIZADOS):
        PASOS_MEMORIZADOS.insert(indice, paso)
        guardar_sesion()
        return True
    else:
        PASOS_MEMORIZADOS.append(paso)
        guardar_sesion()
        return True

def obtener_pasos():
    return PASOS_MEMORIZADOS

# Script JS para generar XPath (reutilizable)
# MEJORADO: Evita usar IDs que contengan números (probablemente dinámicos/generados)
JS_XPATH_SCRIPT = """
function getXPath(element) {
    // Heurística: Si el ID tiene números o es muy largo, probablemente es dinámico. Ignorarlo.
    // Solo usar ID si es puramente texto (letras, guiones, guiones bajos) y corto.
    var isValidId = element.id !== '' && !/[0-9]/.test(element.id);
    
    if (isValidId)
        return 'id("' + element.id + '")';
    
    if (element === document.body)
        return '/html/body';

    var ix = 0;
    var siblings = element.parentNode.childNodes;
    for (var i = 0; i < siblings.length; i++) {
        var sibling = siblings[i];
        if (sibling === element)
            return getXPath(element.parentNode) + '/' + element.tagName + '[' + (ix + 1) + ']';
        if (sibling.nodeType === 1 && sibling.tagName === element.tagName)
            ix++;
    }
}
return getXPath(arguments[0]);
"""

def _generar_xpath_elemento_activo(driver):
    # OBSOLETO: Usar _detectar_foco_con_frames
    try:
        active = driver.switch_to.active_element
        if not active: return None, "No active"
        return driver.execute_script(JS_XPATH_SCRIPT, active), None
    except:
        return None, "Error"

def _detectar_foco_con_frames(driver):
    """
    Recorre desde el contexto principal (default_content) hacia adentro de los iframes
    hasta encontrar el elemento que realmente tiene el foco (el 'leaf node').
    Retorna: (xpath_elemento, lista_xpaths_frames, error)
    """
    frames_path = []
    
    try:
        # 1. Empezar desde el contexto principal para asegurar ruta completa
        driver.switch_to.default_content()
        
        while True:
            active = driver.switch_to.active_element
            if not active:
                return None, [], "No se detectó elemento activo."
            
            tag = active.tag_name.lower()
            
            # Si el elemento activo es un frame/iframe, nos metemos en él
            if tag in ["frame", "iframe"]:
                # Guardamos el xpath de este frame para poder volver a él durante la ejecución
                frame_xpath = driver.execute_script(JS_XPATH_SCRIPT, active)
                frames_path.append(frame_xpath)
                
                # Cambiamos contexto al frame
                driver.switch_to.frame(active)
            else:
                # Es un elemento 'real' (input, button, body, etc.)
                final_xpath = driver.execute_script(JS_XPATH_SCRIPT, active)
                return final_xpath, frames_path, None
                
    except Exception as e:
        return None, [], f"Error detectando foco en frames: {e}"

def agregar_paso_foco(accion, columna=None, formato=None, indice_insercion=None, saltar_al_final=False):
    """
    Agrega un paso basado en el elemento que tiene el foco actual.
    Soporta iframes anidados.
    """
    global PASOS_MEMORIZADOS
    driver = obtener_driver(create_if_missing=False)
    if not driver:
        return False, "Navegador no conectado. Por favor inicie el navegador primero."
    
    # Usar nueva lógica con frames
    xpath, frames, error = _detectar_foco_con_frames(driver)
    
    if error:
        return False, error
    
    if accion == "escribir":
        if not columna:
            return False, "Debe especificar una columna para la acción 'Escribir'."
        paso = {
            "accion": "escribir",
            "xpath": xpath,
            "frames": frames, # Lista de xpaths de frames padres
            "columna": columna,
            "descripcion": f"Escribir '{columna}' en {xpath}" + (f" (dentro de {len(frames)} frames)" if frames else "")
        }
    elif accion == "escribir_fecha":
        if not columna:
            return False, "Debe especificar una columna para la acción 'Escribir Fecha'."
        paso = {
            "accion": "escribir_fecha",
            "xpath": xpath,
            "frames": frames,
            "columna": columna,
            "formato": formato or "%d/%m/%Y",
            "descripcion": f"Escribir Fecha '{columna}' ({formato}) en {xpath}"
        }
    elif accion == "click":
        paso = {
            "accion": "click",
            "xpath": xpath,
            "frames": frames,
            "descripcion": f"Click en {xpath}" + (f" (dentro de {len(frames)} frames)" if frames else "")
        }
    elif accion == "limpiar_campo":
        paso = {
            "accion": "limpiar_campo",
            "xpath": xpath,
            "frames": frames,
            "descripcion": f"Limpiar contenido de {xpath}" + (f" (dentro de {len(frames)} frames)" if frames else "")
        }
    else:
        return False, "Acción no válida."
    
    if saltar_al_final:
        paso["saltar_al_final"] = True
        paso["descripcion"] += " [⏩ SALTAR AL FINAL]"

    _insertar_paso(paso, indice_insercion)
    return True, f"Paso agregado: {paso['descripcion']}"

def agregar_paso_tecla(tecla, indice_insercion=None, saltar_al_final=False):
    """Agrega un paso de pulsar tecla especial."""
    global PASOS_MEMORIZADOS
    paso = {
        "accion": "tecla",
        "tecla": tecla,
        "descripcion": f"Presionar tecla {tecla}"
    }
    if saltar_al_final:
        paso["saltar_al_final"] = True
        paso["descripcion"] += " [⏩ SALTAR AL FINAL]"
        
    _insertar_paso(paso, indice_insercion)
    return True, f"Paso agregado: {paso['descripcion']}"

def agregar_paso_espera(segundos, indice_insercion=None, saltar_al_final=False):
    """Agrega un paso de espera."""
    global PASOS_MEMORIZADOS
    paso = {
        "accion": "espera",
        "tiempo": segundos,
        "descripcion": f"Esperar {segundos} segundos"
    }
    if saltar_al_final:
        paso["saltar_al_final"] = True
        paso["descripcion"] += " [⏩ SALTAR AL FINAL]"
        
    _insertar_paso(paso, indice_insercion)
    return True, f"Paso agregado: {paso['descripcion']}"

def agregar_paso_alerta(accion="aceptar", indice_insercion=None, saltar_al_final=False):
    """Agrega un paso para manejar alertas nativas (popups JS)."""
    global PASOS_MEMORIZADOS
    
    desc = "Aceptar Alerta (OK)" if accion == "aceptar" else "Cancelar Alerta"
    paso = {
        "accion": "alerta",
        "subaccion": accion,
        "descripcion": desc
    }
    if saltar_al_final:
        paso["saltar_al_final"] = True
        paso["descripcion"] += " [⏩ SALTAR AL FINAL]"

    _insertar_paso(paso, indice_insercion)
    return True, f"Paso agregado: {desc}"

def agregar_paso_scroll(direccion, cantidad=0, indice_insercion=None, saltar_al_final=False):
    """
    Agrega un paso de scroll.
    direccion: "arriba", "abajo", "inicio", "fin"
    cantidad: pixels (solo para arriba/abajo)
    """
    global PASOS_MEMORIZADOS
    
    desc = ""
    if direccion == "inicio":
        desc = "Scroll al Inicio de página"
    elif direccion == "fin":
        desc = "Scroll al Final de página"
    elif direccion == "arriba":
        desc = f"Scroll Arriba {cantidad}px"
    elif direccion == "abajo":
        desc = f"Scroll Abajo {cantidad}px"
        
    paso = {
        "accion": "scroll",
        "direccion": direccion,
        "cantidad": cantidad,
        "descripcion": desc
    }
    if saltar_al_final:
        paso["saltar_al_final"] = True
        paso["descripcion"] += " [⏩ SALTAR AL FINAL]"

    _insertar_paso(paso, indice_insercion)
    return True, f"Paso agregado: {desc}"

def cargar_pasos_externos(lista_pasos):
    global PASOS_MEMORIZADOS
    if isinstance(lista_pasos, list):
        PASOS_MEMORIZADOS = lista_pasos
        guardar_sesion()
        return True, f"Cargados {len(lista_pasos)} pasos."
    return False, "Formato inválido."

def agregar_paso_cambiar_ventana(indice=-1, indice_insercion=None):
    """Agrega un paso para cambiar de ventana/pestaña y CAMBIA EL FOCO ACTUAL."""
    global PASOS_MEMORIZADOS
    driver = obtener_driver(create_if_missing=False)
    
    desc = "Cambiar a la última ventana abierta" if indice == -1 else f"Cambiar a ventana índice {indice}"
    paso = {
        "accion": "cambiar_ventana",
        "indice": indice,
        "descripcion": desc
    }
    _insertar_paso(paso, indice_insercion)
    
    # Cambiar el foco del driver en tiempo real para permitir seguir grabando
    if driver:
        try:
            handles = driver.window_handles
            if indice == -1:
                target = handles[-1]
            elif 0 <= indice < len(handles):
                target = handles[indice]
            else:
                return False, f"Paso agregado, pero índice {indice} fuera de rango."
            
            driver.switch_to.window(target)
            return True, f"Paso agregado y FOCO CAMBIADO a: {driver.title}"
        except Exception as e:
            return False, f"Paso agregado pero falló cambio de foco: {e}"
            
    return True, f"Paso agregado: {desc}"

# --- SELECTOR VISUAL (INSPECTOR MODE) ---
JS_SELECTOR_SETUP = """
(function() {
    if (window.__zeus_cleanup) window.__zeus_cleanup();
    
    var lastHighlit = null;
    var tooltip = document.createElement('div');
    tooltip.style.position = 'fixed';
    tooltip.style.padding = '5px 10px';
    tooltip.style.background = 'rgba(0, 0, 0, 0.85)';
    tooltip.style.color = '#fff';
    tooltip.style.borderRadius = '4px';
    tooltip.style.fontSize = '12px';
    tooltip.style.fontFamily = 'monospace';
    tooltip.style.pointerEvents = 'none';
    tooltip.style.zIndex = '999999';
    tooltip.style.display = 'none';
    tooltip.style.boxShadow = '0 2px 5px rgba(0,0,0,0.3)';
    tooltip.style.border = '1px solid #4caf50';
    document.body.appendChild(tooltip);

    function getElementLabel(el) {
        var txt = el.tagName.toLowerCase();
        if (el.id) txt += '#' + el.id;
        if (el.className && typeof el.className === 'string') {
            var classes = el.className.split(' ').filter(c => c.trim().length > 0).slice(0, 2); // Max 2 classes
            if (classes.length) txt += '.' + classes.join('.');
        }
        if (el.getAttribute('name')) txt += '[name="' + el.getAttribute('name') + '"]';
        if (el.getAttribute('title')) txt += ' (Title: ' + el.getAttribute('title').substring(0, 15) + '...)';
        if (el.getAttribute('aria-label')) txt += ' (Label: ' + el.getAttribute('aria-label').substring(0, 15) + '...)';
        
        // Mostrar dimensiones para ayudar con iconos pequeños
        var rect = el.getBoundingClientRect();
        txt += ' [' + Math.round(rect.width) + 'x' + Math.round(rect.height) + ']';
        
        return txt;
    }

    function onMouseOver(e) {
        var target = e.target;
        
        // Ignorar el tooltip mismo
        if (target === tooltip) return;

        if (lastHighlit && lastHighlit !== target) {
            lastHighlit.style.outline = lastHighlit.dataset.zeusOutline || '';
            lastHighlit.style.boxShadow = lastHighlit.dataset.zeusShadow || '';
            lastHighlit = null;
        }

        if (target !== document.body && target !== document.documentElement) {
            lastHighlit = target;
            
            // Guardar estilos originales
            if (lastHighlit.dataset.zeusOutline === undefined) {
                lastHighlit.dataset.zeusOutline = lastHighlit.style.outline;
                lastHighlit.dataset.zeusShadow = lastHighlit.style.boxShadow;
            }
            
            // Resaltado de alta visibilidad (Magenta + Sombra brillante)
            lastHighlit.style.outline = '2px solid #ff00ff'; 
            lastHighlit.style.boxShadow = '0 0 8px #ff00ff, inset 0 0 4px #ff00ff'; 
            lastHighlit.style.cursor = 'crosshair';

            // Actualizar Tooltip
            tooltip.textContent = getElementLabel(target);
            tooltip.style.display = 'block';
            
            // Posicionar tooltip cerca del cursor pero sin estorbar
            var rect = target.getBoundingClientRect();
            var top = rect.top - 30;
            if (top < 0) top = rect.bottom + 10;
            tooltip.style.top = top + 'px';
            tooltip.style.left = rect.left + 'px';
        }
    }
    
    function onMouseOut(e) {
        if (e.target.dataset.zeusOutline !== undefined) {
             e.target.style.outline = e.target.dataset.zeusOutline;
             e.target.style.boxShadow = e.target.dataset.zeusShadow;
             e.target.style.cursor = '';
        }
        tooltip.style.display = 'none';
    }
    
    function onClick(e) {
        e.preventDefault();
        e.stopPropagation();
        
        window.__zeus_selected_element = e.target;
        
        // Feedback visual de éxito (Verde)
        e.target.style.outline = '3px solid #00ff00'; 
        e.target.style.boxShadow = '0 0 10px #00ff00';
        
        tooltip.textContent = "✅ SELECCIONADO: " + getElementLabel(e.target);
        tooltip.style.background = '#006400';
        tooltip.style.borderColor = '#00ff00';
        
        setTimeout(cleanup, 1000); // Dar un momento para ver la confirmación
    }
    
    function cleanup() {
        document.removeEventListener('mouseover', onMouseOver, true);
        document.removeEventListener('mouseout', onMouseOut, true);
        document.removeEventListener('click', onClick, true);
        if (tooltip.parentNode) tooltip.parentNode.removeChild(tooltip);
        window.__zeus_cleanup = null;
    }
    
    window.__zeus_cleanup = cleanup;
    window.__zeus_selected_element = null;
    
    document.addEventListener('mouseover', onMouseOver, true);
    document.addEventListener('mouseout', onMouseOut, true);
    document.addEventListener('click', onClick, true);
})();
"""

def _inject_js_recursive(driver):
    """Inyecta el JS de selector en el frame actual y sus hijos recursivamente."""
    try:
        # Inyectar en contexto actual
        driver.execute_script(JS_SELECTOR_SETUP)
        
        # Buscar frames hijos
        frames = driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame")
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                _inject_js_recursive(driver)
                driver.switch_to.parent_frame()
            except:
                # Recuperar contexto si falla cambio de frame
                try: driver.switch_to.parent_frame()
                except: pass
    except Exception as e:
        print(f"Warning: Error inyectando JS en frame: {e}")

def iniciar_selector_visual():
    """Inyecta JS para permitir selección visual en TODOS los frames."""
    driver = obtener_driver(create_if_missing=False)
    if not driver: return False, "Navegador no conectado."
    
    try:
        # Asegurar que empezamos desde la raíz para cubrir todo
        driver.switch_to.default_content()
        _inject_js_recursive(driver)
        driver.switch_to.default_content() # Volver a raíz
        return True, "Modo Selección Activado. Haga click en el elemento en el navegador."
    except Exception as e:
        return False, f"Error iniciando selector: {e}"

def _check_selection_recursive(driver):
    """Busca recursivamente si hay una selección en algún frame."""
    try:
        # 1. Chequear contexto actual
        element = driver.execute_script("return (window.__zeus_selected_element || null);")
        
        if element:
            # Generar XPath usando el script existente
            xpath = driver.execute_script(JS_XPATH_SCRIPT, element)
            # Limpiar selección
            driver.execute_script("window.__zeus_selected_element = null;")
            return xpath
        
        # 2. Buscar en frames hijos
        frames = driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame")
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                found_xpath = _check_selection_recursive(driver)
                if found_xpath:
                    return found_xpath 
                driver.switch_to.parent_frame()
            except:
                try: driver.switch_to.parent_frame()
                except: pass
                
        return None
    except Exception:
        return None

def obtener_seleccion_visual():
    """Consulta recursivamente todos los frames para ver si el usuario hizo click."""
    driver = obtener_driver(create_if_missing=False)
    if not driver: return False, "Navegador desconectado."
    
    try:
        # Empezar búsqueda desde la raíz
        driver.switch_to.default_content()
        xpath = _check_selection_recursive(driver)
        
        if xpath:
            return True, xpath
        else:
            return False, None
            
    except Exception as e:
        return False, f"Error obteniendo selección: {e}"

def _generar_xpath_texto(texto, exacto, tag="*", ignore_case=False):
    """Genera un XPath para buscar elementos por texto/atributos con filtro de etiqueta."""
    
    # Helpers para case-insensitive
    def to_lower(attr):
        return f"translate({attr}, 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúñ')"
    
    texto_clean = texto.strip()
    texto_lower = texto_clean.lower()
    
    # Mejora: Si el usuario busca 'x' o 'X', buscar también el símbolo '×' (común en botones cerrar)
    if texto_lower == "x":
        extra_match = " or normalize-space(text())='×' or @title='×' or @aria-label='Close' or @aria-label='Cerrar'"
    else:
        extra_match = ""

    if exacto:
        if ignore_case:
            # Búsqueda exacta pero ignorando mayúsculas/minúsculas
            # Usamos normalize-space() normal y comparamos contra el lower del texto
            cond = f"{to_lower('normalize-space(text())')}='{texto_lower}'"
            cond += f" or {to_lower('@title')}='{texto_lower}'"
            cond += f" or {to_lower('@aria-label')}='{texto_lower}'"
            # Fallback para elementos anidados (spans dentro de div) - Busca nodos hoja que contengan el texto
            cond += f" or ({to_lower('normalize-space(.)')}='{texto_lower}' and count(*)=0)" 
            return f"//{tag}[({cond}){extra_match}]"
        else:
            return f"//{tag}[normalize-space(text())='{texto_clean}' or @title='{texto_clean}' or @id='{texto_clean}' or @value='{texto_clean}' or @aria-label='{texto_clean}' or (normalize-space(.)='{texto_clean}' and count(*)=0){extra_match}]"
    else:
        # Contiene (Parcial)
        if ignore_case:
            cond = f"contains({to_lower('text()')}, '{texto_lower}')"
            cond += f" or contains({to_lower('@title')}, '{texto_lower}')"
            cond += f" or contains({to_lower('@aria-label')}, '{texto_lower}')"
            # Robustez tipo Visual: Buscar también en el contenido de texto (string-value) de nodos hoja
            cond += f" or (contains({to_lower('.')}, '{texto_lower}') and count(*)=0)"
            return f"//{tag}[({cond}){extra_match}]"
        else:
            return f"//{tag}[contains(text(), '{texto_clean}') or contains(@title, '{texto_clean}') or contains(@id, '{texto_clean}') or contains(@class, '{texto_clean}') or contains(@alt, '{texto_clean}') or contains(@aria-label, '{texto_clean}') or (contains(., '{texto_clean}') and count(*)=0){extra_match}]"

def agregar_paso_click_texto(texto, exacto=False, es_dinamico=False, tag="*", tipo_seleccion="texto", ignore_case=False, indice_insercion=None, xpath_contenedor=None, usar_indice_contenedor=False, saltar_al_final=False):
    """
    Agrega un paso de click basado en búsqueda de texto, opción de lista (select) o XPath directo.
    
    tipo_seleccion: 
      - "texto": busca en el DOM por texto visible/atributos (comportamiento original)
      - "lista": busca <option> dentro de <select> (o emulados) cuyo texto coincida
      - "xpath": usa el selector proporcionado directamente
      
    xpath_contenedor: (Opcional) XPath de un elemento padre para restringir la búsqueda.
    usar_indice_contenedor: (Opcional) Si es True, el valor (Excel/Fijo) se usa como índice (1,2,3) para elegir cual contenedor clickear.
    """
    global PASOS_MEMORIZADOS
    
    # Base del paso
    paso = {
        "accion": "click_texto" if tipo_seleccion != "seleccionar_opcion" else "click_texto", # Unificamos lógica
        "es_dinamico": es_dinamico,
        "xpath_contenedor": xpath_contenedor,
        "usar_indice_contenedor": usar_indice_contenedor
    }
    
    # Soporte para lista de contenedores (Multi-Visual)
    # Si 'xpath_contenedor' es una lista, la guardamos como 'contenedores_visuales'
    if isinstance(xpath_contenedor, list):
        paso["contenedores_visuales"] = xpath_contenedor
        paso["xpath_contenedor"] = None # Limpiamos el campo legacy simple
    else:
        paso["contenedores_visuales"] = []

    if tipo_seleccion == "xpath":
        xpath = texto # En este caso 'texto' es el xpath crudo
        desc = f"Click Personalizado (XPath): {xpath}"
        paso.update({
            "accion": "click", # Mantenemos 'click' simple para xpath directo
            "xpath": xpath,
            "descripcion": desc
        })
    
    elif usar_indice_contenedor:
        # NUEVO MODO: Selección por Índice de Contenedor
        if es_dinamico:
            desc = f"Click por Índice de Contenedor (Excel: '{texto}')"
        else:
            desc = f"Click por Índice de Contenedor (Fijo: '{texto}')"
            
        if paso.get("contenedores_visuales"):
            desc += f" [De Lista de {len(paso['contenedores_visuales'])} Opciones]"
            
        paso.update({
            "accion": "click_texto",
            "columna": texto if es_dinamico else None,
            "valor": texto if not es_dinamico else None,
            "descripcion": desc
        })

    elif tipo_seleccion == "lista":
        # Modo selección de lista desplegable
        if es_dinamico:
            desc = f"Seleccionar de Lista (Dinámico): Columna '{texto}'"
        else:
            desc = f"Seleccionar de Lista (Fijo): '{texto}'"
            
        if xpath_contenedor:
            desc += " [En Contenedor]"
        if paso.get("contenedores_visuales"):
            desc += f" [En {len(paso['contenedores_visuales'])} Contenedores]"

        paso.update({
            "accion": "click_texto", # Usamos la logica de texto pero con tag específico si se quiere
            "tipo_seleccion": "lista", # Marker para logica interna
            "valor": texto,     # Columna o Valor fijo
            "exacto": exacto,
            "ignore_case": ignore_case,
            "tag": tag, # Permitimos tag custom en lista
            "descripcion": desc
        })
        # Nota: antes usabamos 'seleccionar_opcion' como accion, pero 'click_texto' es más versátil en el loop principal 
        # si le pasamos los parametros correctos.
        # Ajustamos el loop principal para manejar 'tipo_seleccion' = 'lista' si es necesario, 
        # o simplemente lo tratamos como busqueda de texto con tag='option' (default)

    else:
        # Modo Texto (Original)
        if es_dinamico:
            desc = f"Click Texto Dinámico (Columna: {texto})"
            if xpath_contenedor:
                desc += " [En Contenedor]"
            if paso.get("contenedores_visuales"):
                desc += f" [En {len(paso['contenedores_visuales'])} Contenedores]"
            
            paso.update({
                "accion": "click_texto",
                "tipo_busqueda": "texto_dinamico",
                "columna": texto,
                "exacto": exacto,
                "tag": tag,
                "ignore_case": ignore_case,
                "descripcion": desc
            })
        else:
            # Si es fijo, generamos el XPath ahora
            # Si hay contenedor, lo preponemos
            xpath_final = _generar_xpath_texto(texto, exacto, tag, ignore_case)
            if xpath_contenedor:
                xpath_final = f"{xpath_contenedor}{xpath_final}"
            elif paso.get("contenedores_visuales"):
                # Generar OR de todos los contenedores
                parts = [f"{xp}{xpath_final}" for xp in paso["contenedores_visuales"]]
                xpath_final = " | ".join(parts)
                
            desc = f"Click Texto Fijo: '{texto}'"
            if xpath_contenedor:
                desc += " [En Contenedor]"
            if paso.get("contenedores_visuales"):
                desc += f" [En {len(paso['contenedores_visuales'])} Contenedores]"

            paso.update({
                "accion": "click", # Se convierte en click estático
                "xpath": xpath_final,
                "descripcion": desc
            })
    
    if saltar_al_final:
        paso["saltar_al_final"] = True
        paso["descripcion"] += " [⏩ SALTAR AL FINAL]"

    _insertar_paso(paso, indice_insercion)
    return True, f"Paso agregado: {desc}"

def smart_find_element(driver, xpath, timeout=5):
    """
    Busca un elemento en el contexto actual, y si falla, 
    busca recursivamente en todos los iframes/frames desde la raíz.
    Deja el driver en el contexto del elemento encontrado.
    Incluye reintentos por tiempo para elementos dinámicos.
    """
    start_time = time.time()
    
    while True:
        # 1. Intentar en el contexto actual
        try:
            return driver.find_element(By.XPATH, xpath)
        except:
            pass # Continuar a búsqueda profunda
        
        # 2. Búsqueda recursiva desde la raíz
        driver.switch_to.default_content()
        
        # Revisar raíz primero
        try:
            return driver.find_element(By.XPATH, xpath)
        except:
            pass

        def _recursive_search(drv):
            frames = drv.find_elements(By.TAG_NAME, "iframe") + drv.find_elements(By.TAG_NAME, "frame")
            for frame in frames:
                try:
                    drv.switch_to.frame(frame)
                    # Buscar en este frame
                    try:
                        el = drv.find_element(By.XPATH, xpath)
                        return el # ¡Encontrado!
                    except:
                        pass
                    
                    # No está aquí, buscar en hijos
                    found = _recursive_search(drv)
                    if found: return found
                    
                    # Si no está en hijos, volver al padre para seguir loop
                    drv.switch_to.parent_frame()
                except:
                    # Si falla el switch frame, intentar recuperar
                    try: drv.switch_to.parent_frame()
                    except: pass
            return None

        found = _recursive_search(driver)
        if found: return found
        
        # Verificar timeout
        if time.time() - start_time > timeout:
            raise Exception(f"Elemento {xpath} no encontrado en ningún frame tras {timeout}s.")
        
        time.sleep(0.5) # Esperar antes de reintentar


def ejecutar_secuencia(df, delay_pasos=0.5):
    """
    Ejecuta la secuencia de pasos memorizados para cada fila del DataFrame.
    Retorna un generador de logs.
    """
    global PASOS_MEMORIZADOS, EJECUCION_ACTIVA, ULTIMO_ERROR
    
    # Reset flags
    with EJECUCION_LOCK:
        EJECUCION_ACTIVA = True
    ULTIMO_ERROR = None

    driver = obtener_driver()
    
    if not driver:
        yield "❌ Error: El navegador no está abierto."
        return
    
    if not PASOS_MEMORIZADOS:
        yield "❌ Error: No hay pasos memorizados."
        return
    
    yield f"🚀 Iniciando secuencia con {len(df)} registros y {len(PASOS_MEMORIZADOS)} pasos por registro..."
    
    count_ok = 0
    count_err = 0
    failed_rows = []
    total = len(df)
    
    # Mapa de teclas Selenium
    TECLAS_MAP = {
        "ENTER": Keys.ENTER,
        "TAB": Keys.TAB,
        "ESCAPE": Keys.ESCAPE,
        "DOWN": Keys.ARROW_DOWN,
        "UP": Keys.ARROW_UP
    }

    for index, row in df.iterrows():
        try:
            yield f"--- 🏁 Iniciando Paciente {index + 1}/{total} 🏁 ---"
        
            # --- LÓGICA CONDICIONAL DE FLUJO (WATERFALL) ---
            # 1. Por defecto: Principal
            pasos_actuales = PASOS_MEMORIZADOS
            nombre_flujo = "PRINCIPAL"
            
            # 2. Evaluar Flujos Condicionales en orden de índice (0, 1, 2...)
            # Si FLUJOS_CONDICIONALES tiene datos, evaluamos.
            if FLUJOS_CONDICIONALES:
                flujo_seleccionado = None
                
                # Ordenar por índice para asegurar precedencia (0 -> 1 -> 2)
                for idx in sorted(FLUJOS_CONDICIONALES.keys()):
                    flujo_info = FLUJOS_CONDICIONALES[idx]
                    cond = flujo_info.get("condicion")
                    
                    if not cond:
                        continue # Sin condición, no se puede evaluar (o podría ser default?)
                        
                    try:
                        cond_cumplida = False
                        tipo_cond = cond.get("tipo", "texto")
                        valor_cond = cond.get("valor", "")
                        
                        if tipo_cond == "excel":
                            col_cond = cond.get("columna")
                            if col_cond and col_cond in df.columns:
                                val_celda = str(row[col_cond]).strip().lower()
                                vals_trigger = [v.strip().lower() for v in valor_cond.split("|")]
                                
                                if val_celda in vals_trigger:
                                    cond_cumplida = True
                                    yield f"   🧬 Condición {idx+1} ({flujo_info.get('nombre')}): Columna '{col_cond}'=='{val_celda}' -> CUMPLE"
                                else:
                                    # yield f"   🧬 Condición {idx+1}: No cumple ('{val_celda}' not in '{vals_trigger}')"
                                    pass
                            else:
                                yield f"   ⚠️ Condición {idx+1}: Columna '{col_cond}' no encontrada."

                        elif tipo_cond == "excel_multi":
                            reglas = cond.get("reglas", [])
                            # Lógica AND: todas las reglas deben cumplirse
                            cumple_todas = True
                            detalles = []
                            
                            if not reglas:
                                cumple_todas = False # Sin reglas no cumple
                            
                            for r in reglas:
                                r_col = r.get("columna")
                                r_val = r.get("valor", "")
                                if r_col and r_col in df.columns:
                                    val_celda = str(row[r_col]).strip().lower()
                                    vals_trigger = [v.strip().lower() for v in r_val.split("|")]
                                    
                                    if val_celda in vals_trigger:
                                        detalles.append(f"{r_col}='{val_celda}' (OK)")
                                    else:
                                        cumple_todas = False
                                        detalles.append(f"{r_col}='{val_celda}' (NO)")
                                        break # Short-circuit
                                else:
                                    cumple_todas = False
                                    detalles.append(f"Columna '{r_col}' faltante")
                                    break
                            
                            if cumple_todas:
                                cond_cumplida = True
                                yield f"   🧬 Condición {idx+1} (Multi): {', '.join(detalles)} -> CUMPLE"
                        
                        else: # Texto en pantalla
                            driver.switch_to.default_content()
                            if valor_cond in driver.page_source:
                                try:
                                    els = driver.find_elements(By.XPATH, f"//*[contains(text(), '{valor_cond}')]")
                                    if any(e.is_displayed() for e in els):
                                        cond_cumplida = True
                                        yield f"   👀 Condición {idx+1} ({flujo_info.get('nombre')}): Texto '{valor_cond}' DETECTADO -> CUMPLE"
                                except: pass
                        
                        if cond_cumplida:
                            flujo_seleccionado = flujo_info
                            break # WATERFALL: El primero que cumple gana
                            
                    except Exception as e_cond:
                        yield f"   ⚠️ Error evaluando condición {idx+1}: {e_cond}"

                if flujo_seleccionado:
                    pasos_actuales = flujo_seleccionado["pasos"]
                    nombre_flujo = flujo_seleccionado["nombre"]
                else:
                    yield "   ℹ️ Ninguna condición alternativa se cumplió -> Usando BOT PRINCIPAL"

            # Fallback Legacy directo (si no hay FLUJOS_CONDICIONALES pero sí vars legacy)
            elif PASOS_ALTERNATIVOS and CONDICION_EJECUCION:
                # ... (Lógica legacy si fuera necesaria, pero FLUJOS_CONDICIONALES[0] ya debería cubrirlo)
                pass

            yield f"   🤖 Ejecutando {len(pasos_actuales)} pasos del flujo {nombre_flujo}..."

            try:
                for i, paso in enumerate(pasos_actuales):
                    # Check Stop Flag per Step
                    if not EJECUCION_ACTIVA:
                        raise Exception("Ejecución detenida por usuario.")

                    # Bloque TRY/EXCEPT por paso para manejar "Opcional"
                    try:
                        for log_msg in _ejecutar_logica_paso(driver, paso, row, df, i, delay_pasos):
                            yield log_msg
                        
                        # --- SALTAR AL FINAL SI ÉXITO ---
                        if paso.get("saltar_al_final", False):
                            yield "   ⏩ CONDICIÓN CUMPLIDA: Saltando al último paso..."
                            if i < len(pasos_actuales) - 1:
                                # Ejecutar el último paso explícitamente antes de salir
                                try:
                                    ultimo_paso = pasos_actuales[-1]
                                    # Validar que el último paso no sea el mismo que acabamos de ejecutar
                                    # (Aunque el check i < len-1 ya cubre eso, pero por seguridad)
                                    yield "   🏁 Ejecutando paso final anticipadamente..."
                                    for log_msg in _ejecutar_logica_paso(driver, ultimo_paso, row, df, len(pasos_actuales)-1, delay_pasos):
                                        yield log_msg
                                except Exception as e_jump:
                                    yield f"   ⚠️ Error ejecutando paso final tras salto: {e_jump}"
                                    raise e_jump # Opcional: marcar como error o warning
                            
                            break # Salir del bucle de pasos (terminar paciente)

                    except Exception as e_paso:
                        es_opcional = paso.get("opcional", False)
                        if es_opcional:
                             yield f"   ⚠️ [OPCIONAL] Paso {i+1} omitido: {e_paso}"
                             continue
                        else:
                             raise e_paso # Re-raise to be caught by row handler

            except Exception as e_row_critico:
                yield f"   ❌ Error Crítico en Paso {i+1}: {e_row_critico}"
                count_err += 1
                failed_rows.append(index + 1)
                
                # --- LÓGICA DE RECUPERACIÓN (Ir al último paso) ---
                if pasos_actuales and i < len(pasos_actuales) - 1:
                    yield "   ⚠️ Intentando ejecutar el ÚLTIMO PASO para cerrar proceso..."
                    try:
                        ultimo_paso = pasos_actuales[-1]
                        for log_msg in _ejecutar_logica_paso(driver, ultimo_paso, row, df, len(pasos_actuales)-1, delay_pasos):
                            yield log_msg
                        yield "   ✅ Último paso ejecutado correctamente."
                    except Exception as e_rec:
                        yield f"   ☠️ Falló el paso de recuperación: {e_rec}"
                
                continue # Pasa al siguiente paciente

            # Si llega aquí, todo OK
            count_ok += 1
            yield "   ✅ Paciente completado exitosamente."

        except Exception as e_row_gen:
             yield f"❌ Error general en fila {index+1}: {e_row_gen}"
             count_err += 1
             failed_rows.append(index + 1)

    # --- REPORTE FINAL ---
    yield " "
    yield f"📊 REPORTE FINAL DE EJECUCIÓN:"
    yield f"✅ Exitosos: {count_ok}"
    yield f"❌ Fallidos: {count_err}"
    if failed_rows:
        yield f"📝 Filas con error: {failed_rows}"
    if count_err > 0:
        yield "⚠️ Revise los logs anteriores para ver detalles de los errores."

def _ejecutar_logica_paso(driver, paso, row, df, i, delay_pasos):
    """Lógica encapsulada de ejecución de un solo paso."""
    
    accion = paso["accion"]
    
    # --- GESTIÓN DE FRAMES Y BÚSQUEDA ROBUSTA ---
    if accion in ["click", "escribir", "click_texto", "limpiar_campo", "escribir_fecha"]:
        element_found = None
        target_xpath = paso.get("xpath")

        # Lógica especial para Click Dinámico
        if accion == "click_texto":
            if paso.get("tipo_seleccion") == "xpath":
                    # Si es XPath directo, usarlo tal cual
                    target_xpath = paso.get("xpath")
            
            elif paso.get("usar_indice_contenedor"):
                # --- NUEVO: SELECCIÓN POR ÍNDICE DE CONTENEDOR ---
                val_str = None
                col = paso.get("columna")
                if col:
                    if col in df.columns:
                        val_str = str(row[col]).strip()
                    else:
                        raise Exception(f"Columna '{col}' no encontrada.")
                else:
                    val_str = str(paso.get("valor", "")).strip()

                try:
                    # Convertir a int (ej: "1.0" -> 1)
                    if not val_str: raise ValueError("Vacío")
                    idx = int(float(val_str))
                
                    contenedores = paso.get("contenedores_visuales", [])
                    # Fallback legacy
                    if not contenedores and paso.get("xpath_contenedor"):
                        contenedores = [paso.get("xpath_contenedor")]
                    
                    if not contenedores:
                        raise Exception("No hay contenedores definidos para selección por índice.")
                    
                    if 1 <= idx <= len(contenedores):
                        target_xpath = contenedores[idx-1]
                        # Log handled by caller? No, we can't yield here easily without changing signature.
                        # For now, we assume success or raise exception.
                    else:
                        raise Exception(f"Índice '{idx}' fuera de rango (1-{len(contenedores)}).")
                    
                except ValueError:
                    raise Exception(f"Valor '{val_str}' no es un número válido para índice.")

            elif paso.get("es_dinamico"):
                col_dinamica = paso.get("columna") or paso.get("texto") or paso.get("valor") # Compatibilidad
                if col_dinamica and col_dinamica in df.columns:
                    val_dinamico = str(row[col_dinamica]).strip() # Strip para limpieza
                    if val_dinamico.lower() in ["nan", "nat", "none", ""]:
                        print(f"Valor vacío en columna '{col_dinamica}', saltando click.") # Print fallback
                        return # Skip silently?
                        
                    target_xpath = _generar_xpath_texto(val_dinamico, paso.get("exacto", False), paso.get("tag", "*"), paso.get("ignore_case", False))
                
                    # --- NUEVO: PREPEND CONTENEDOR SI EXISTE ---
                    xpath_cont = paso.get("xpath_contenedor")
                    contenedores_list = paso.get("contenedores_visuales", [])
                
                    if xpath_cont:
                        target_xpath = f"{xpath_cont}{target_xpath}"
                    elif contenedores_list:
                        # Generar OR de todos los contenedores
                        parts = [f"{xp}{target_xpath}" for xp in contenedores_list]
                        target_xpath = " | ".join(parts)
                    
                else:
                    raise Exception(f"Columna dinámica '{col_dinamica}' no encontrada.")
            else:
                # Caso Texto Fijo pero guardado como click_texto
                val_fijo = paso.get("valor")
                if val_fijo:
                    target_xpath = _generar_xpath_texto(val_fijo, paso.get("exacto", False), paso.get("tag", "*"), paso.get("ignore_case", False))
                    xpath_cont = paso.get("xpath_contenedor")
                    contenedores_list = paso.get("contenedores_visuales", [])
                
                    if xpath_cont:
                        target_xpath = f"{xpath_cont}{target_xpath}"
                    elif contenedores_list:
                        parts = [f"{xp}{target_xpath}" for xp in contenedores_list]
                        target_xpath = " | ".join(parts)


        # 1. Intentar ruta rápida (Frames grabados)
        try:
            driver.switch_to.default_content()
            frames = paso.get("frames", [])
            if frames:
                for f_xpath in frames:
                    f_el = driver.find_element(By.XPATH, f_xpath)
                    driver.switch_to.frame(f_el)
        
            # Buscar elemento en contexto final
            element_found = driver.find_element(By.XPATH, target_xpath)
    
        except Exception as e_fast:
            # 2. Si falla, usar SMART FIND (Búsqueda profunda en todos los frames)
            try:
                element_found = smart_find_element(driver, target_xpath)
            except Exception as e_smart:
                    raise Exception(f"Elemento no encontrado: {e_smart}")

        # --- ACCIÓN: CLICK O CLICK_TEXTO ---
        if accion == "click" or accion == "click_texto":
            try:
                # Highlight visual antes de click
                try:
                    driver.execute_script("arguments[0].style.border='3px solid red';", element_found)
                    time.sleep(0.2) # Pequeña pausa para efecto visual
                    driver.execute_script("arguments[0].style.border='';", element_found)
                except: pass

                # SOPORTE PARA DROPDOWNS NATIVOS (<select>)
                if element_found.tag_name.lower() == "option":
                    try:
                        from selenium.webdriver.support.ui import Select
                        parent_select = element_found.find_element(By.XPATH, "..")
                        if parent_select.tag_name.lower() == "optgroup":
                            parent_select = parent_select.find_element(By.XPATH, "..")
                        
                        if parent_select.tag_name.lower() == "select":
                            sel_obj = Select(parent_select)
                            sel_obj.select_by_visible_text(element_found.text)
                            time.sleep(delay_pasos)
                            return # Done
                    except Exception:
                        pass 

                try:
                    element_found.click()
                except Exception as e_click:
                        # Intentar JS click
                        msg = str(e_click).lower()
                        if "not interactable" in msg or "click intercepted" in msg:
                            driver.execute_script("arguments[0].click();", element_found)
                        else:
                            raise e_click
            except Exception as e:
                raise Exception(f"Fallo Click: {e}")

        # --- ACCIÓN: LIMPIAR CAMPO ---
        elif accion == "limpiar_campo":
            try:
                # Highlight visual antes de limpiar
                try:
                    driver.execute_script("arguments[0].style.border='3px solid orange';", element_found)
                    time.sleep(0.2)
                    driver.execute_script("arguments[0].style.border='';", element_found)
                except: pass

                # Estrategia 1: clear() nativo
                element_found.clear()
            
                # Verificar si queda contenido (para inputs complejos)
                val = element_found.get_attribute("value")
                if val and len(val) > 0:
                    # Estrategia 2: Ctrl+A + Delete
                    element_found.click()
                    element_found.send_keys(Keys.CONTROL + "a")
                    element_found.send_keys(Keys.DELETE)
            except Exception as e:
                # yield f"   ❌ Paso {i+1} (Limpiar): Falló - {e}"
                raise Exception(f"Fallo Limpiar: {e}")

        # --- ACCIÓN: ESCRIBIR FECHA ---
        elif accion == "escribir_fecha":
            col = paso.get("columna")
            fmt = paso.get("formato", "%d/%m/%Y")
        
            if not col or col not in df.columns:
                yield f"   ⚠️ Paso {i+1}: Columna '{col}' no encontrada."
                return
            
            raw_val = row[col]
        
            # Convertir a datetime si no lo es
            try:
                if pd.isna(raw_val) or str(raw_val).strip() == "":
                     # Valor vacío
                     valor_fecha = ""
                else:
                    # Intentar parsear con pandas
                    dt_val = pd.to_datetime(raw_val)
                    valor_fecha = dt_val.strftime(fmt)
            except Exception as e_date:
                yield f"   ⚠️ Paso {i+1}: Error formateando fecha '{raw_val}' -> {e_date}"
                # Fallback: usar string directo
                valor_fecha = str(raw_val)

            try:
                # Limpiar antes de escribir
                element_found.clear()
                # Escribir
                element_found.send_keys(valor_fecha)
            except Exception as e:
                raise Exception(f"Fallo Escribir Fecha: {e}")

        # --- ACCIÓN: ESCRIBIR ---
        elif accion == "escribir":
            col = paso.get("columna")
            if not col:
                yield f"   ⚠️ Paso {i+1}: No hay columna definida."
                return

            if col not in df.columns:
                yield f"   ⚠️ Paso {i+1}: Columna '{col}' no existe en Excel."
                return
            
            valor = str(row[col])
            if valor.lower() == "nan" or valor.lower() == "nat": valor = ""
        
            # Estrategia 1: Estándar
            try:
                element_found.clear()
                element_found.send_keys(valor)
            except Exception as e_std:
                # Si falla, intentar estrategias robustas
                try:
                    # Estrategia 2: Click + Keys (si clear falla o elemento no interactuable)
                    element_found.click()
                    element_found.send_keys(Keys.CONTROL + "a")
                    element_found.send_keys(Keys.DELETE)
                    element_found.send_keys(Keys.DELETE)
                    element_found.send_keys(valor)
                except:
                    # Estrategia 3: JavaScript Injection (Último recurso)
                    driver.execute_script("arguments[0].value = arguments[1];", element_found, valor)
                    # Disparar eventos de cambio para que la app detecte el valor
                    driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", element_found)
                    driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", element_found)

        # --- ACCIÓN: TECLA ---
        elif accion == "tecla":
            tecla_nombre = paso["tecla"]
            teclas_map = {
                "ENTER": Keys.ENTER,
                "TAB": Keys.TAB,
                "ESCAPE": Keys.ESCAPE,
                "DOWN": Keys.ARROW_DOWN,
                "UP": Keys.ARROW_UP
            }
            k = teclas_map.get(tecla_nombre)
            if k:
                try:
                    # Enviamos la tecla al elemento activo o al body
                    driver.switch_to.active_element.send_keys(k)
                except Exception as e:
                    raise Exception(f"Fallo Tecla: {e}")
            else:
                yield f"   ⚠️ Paso {i+1}: Tecla desconocida {tecla_nombre}"

        # --- ACCIÓN: ESPERA ---
        elif accion == "espera":
            time.sleep(paso["tiempo"])

        # --- ACCIÓN: CAMBIAR VENTANA ---
        elif accion == "cambiar_ventana":
            idx = paso["indice"]
            try:
                handles = driver.window_handles
                if idx == -1:
                    target = handles[-1]
                elif 0 <= idx < len(handles):
                    target = handles[idx]
                else:
                    yield f"   ⚠️ Paso {i+1}: Índice de ventana {idx} fuera de rango."
                    return
            
                driver.switch_to.window(target)
            except Exception as e:
                 raise Exception(f"Fallo Cambio Ventana: {e}")

        # --- ACCIÓN: ALERTA ---
        elif accion == "alerta":
            sub = paso.get("subaccion", "aceptar")
            try:
                # Esperar un poco a que aparezca la alerta
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                txt_alert = alert.text
            
                if sub == "aceptar":
                    alert.accept()
                else:
                    alert.dismiss()
                
            except Exception as e:
                 raise Exception(f"Fallo Alerta: {e}")

        # --- ACCIÓN: SCROLL ---
        elif accion == "scroll":
            direc = paso.get("direccion", "abajo")
            cant = paso.get("cantidad", 0)
        
            try:
                if direc == "inicio":
                    driver.execute_script("window.scrollTo(0, 0);")
                elif direc == "fin":
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                elif direc == "arriba":
                    driver.execute_script(f"window.scrollBy(0, -{cant});")
                elif direc == "abajo":
                    driver.execute_script(f"window.scrollBy(0, {cant});")
            except Exception as e:
                raise Exception(f"Fallo Scroll: {e}")

    # Pequeña pausa entre pasos para estabilidad
    if delay_pasos > 0:
        time.sleep(delay_pasos)

# --- PERSISTENCIA DE SESIÓN ---

SESSION_FILE = "bot_session.json"

def guardar_sesion():
    global PASOS_MEMORIZADOS, FLUJOS_CONDICIONALES
    try:
        # Convertir claves int a str para JSON
        flujos_export = {str(k): v for k, v in FLUJOS_CONDICIONALES.items()}
        
        state = {
            "pasos_principales": PASOS_MEMORIZADOS,
            "flujos_condicionales": flujos_export
        }
        with open(SESSION_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=4)
        return True, f"Sesión guardada en {SESSION_FILE}"
    except Exception as e:
        print(f"Error guardando sesión: {e}")
        return False, f"Error guardando sesión: {e}"

def cargar_sesion():
    global PASOS_MEMORIZADOS, FLUJOS_CONDICIONALES, PASOS_ALTERNATIVOS, CONDICION_EJECUCION
    try:
        if not os.path.exists(SESSION_FILE):
            return False, "No existe archivo de sesión anterior."
            
        with open(SESSION_FILE, "r", encoding="utf-8") as f:
            state = json.load(f)
            
        PASOS_MEMORIZADOS = state.get("pasos_principales", [])
        
        # Cargar flujos condicionales
        raw_flows = state.get("flujos_condicionales", {})
        # Convertir claves back to int
        FLUJOS_CONDICIONALES = {int(k): v for k, v in raw_flows.items()}
        
        # Sync Legacy (Slot 0)
        if 0 in FLUJOS_CONDICIONALES:
            PASOS_ALTERNATIVOS = FLUJOS_CONDICIONALES[0].get("pasos", [])
            CONDICION_EJECUCION = FLUJOS_CONDICIONALES[0].get("condicion")
        else:
            PASOS_ALTERNATIVOS = []
            CONDICION_EJECUCION = None
            
        return True, f"Sesión restaurada ({len(PASOS_MEMORIZADOS)} pasos, {len(FLUJOS_CONDICIONALES)} flujos)."
    except Exception as e:
        return False, f"Error cargando sesión: {e}"

