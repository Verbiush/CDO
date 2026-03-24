import requests
import sys

SERVER_IP = "18.118.37.215"
API_PORT = "8000"
WEB_PORT = "8501"

def check_connection():
    print("--------------------------------------------------")
    print("Verificando conexión con el servidor AWS CDO...")
    print("--------------------------------------------------")

    # 1. Check API Port (Correct for Agent)
    api_url = f"http://{SERVER_IP}:{API_PORT}/ping"
    print(f"\n[1] Probando conexión API (Agente): {api_url}")
    try:
        response = requests.get(api_url, timeout=5)
        if response.status_code == 200:
            print("✅ CONEXIÓN EXITOSA: El puerto API (8000) responde correctamente.")
            print("   -> Configure su agente con esta URL: http://18.118.37.215:8000")
        else:
            print(f"⚠️  ADVERTENCIA: El servidor respondió con código {response.status_code}")
    except Exception as e:
        print(f"❌ ERROR: No se pudo conectar al puerto API (8000).")
        print(f"   Detalle: {e}")

    # 2. Check Web Port (Incorrect for Agent)
    web_url = f"http://{SERVER_IP}:{WEB_PORT}"
    print(f"\n[2] Probando conexión Web (UI): {web_url}")
    try:
        response = requests.get(web_url, timeout=5)
        if response.status_code == 200:
            print("✅ CONEXIÓN EXITOSA: La página web está accesible.")
            print("   -> ¡IMPORTANTE! NO use esta URL para el agente. Use el puerto 8000.")
        else:
            print(f"⚠️  ADVERTENCIA: La página web respondió con código {response.status_code}")
    except Exception as e:
        print(f"❌ ERROR: No se pudo conectar al puerto Web (8501).")
        print(f"   Detalle: {e}")

    print("\n--------------------------------------------------")
    print("Resumen:")
    print(f"1. Para usar la App Web: http://{SERVER_IP}:{WEB_PORT}")
    print(f"2. Para configurar el Agente Local: http://{SERVER_IP}:{API_PORT}")
    print("--------------------------------------------------")

if __name__ == "__main__":
    check_connection()
    input("\nPresione Enter para salir...")
