"""
Script para crear el ejecutable de RDA Pólizas D&A
Ejecutar desde la carpeta donde están los .py: python build_exe.py
"""

import os
import subprocess
import sys

NOMBRE_APP = "RDA_Polizas_DA"

ARCHIVOS_REQUERIDOS = [
    "main.py",
    "interfaz_GUI.py",
    "controller.py",
    "Consolidacion.py",
    "utils.py",
]


def install_pyinstaller():
    """Instalar PyInstaller si no está disponible"""
    try:
        import PyInstaller  # noqa: F401
        print("✅ PyInstaller ya está instalado")
    except ImportError:
        print("📦 Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✅ PyInstaller instalado correctamente")


def create_executable():
    """Crear el ejecutable"""
    print("🚀 Creando ejecutable...")

    cmd = [
        "pyinstaller",
        # --onedir: una CARPETA con el .exe adentro. Arranca más rápido que
        # --onefile y el antivirus corporativo la marca con menos frecuencia.
        # Para enviarla al stakeholder, se comprime la carpeta en un .zip.
        "--onedir",
        "--windowed",              # Sin ventana de consola (app con GUI)
        f"--name={NOMBRE_APP}",
        "--icon=NONE",             # Sin icono personalizado (cambiar si hay un .ico)
        "--clean",                 # Limpia caché de builds anteriores
        "--noconfirm",             # No pregunta al sobrescribir dist/ y build/

        # Módulos que PyInstaller no detecta solo.
        # Nota: todavía NO se incluye win32com porque esta versión escribe el
        # Excel con openpyxl. Cuando el escritor cambie a Excel COM, agregar:
        #   --hidden-import=win32com  --hidden-import=win32com.client
        #   --hidden-import=pywintypes
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.cell._writer",
        "--hidden-import=pandas",

        # Punto de entrada
        "main.py",
    ]

    try:
        subprocess.run(cmd, check=True)

        carpeta = os.path.join("dist", NOMBRE_APP)
        exe_path = os.path.join(carpeta, f"{NOMBRE_APP}.exe")

        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            total_mb = sum(
                os.path.getsize(os.path.join(raiz, f))
                for raiz, _, archivos in os.walk(carpeta)
                for f in archivos
            ) / (1024 * 1024)

            print("✅ Ejecutable creado exitosamente!")
            print(f"📁 Ubicación: {exe_path}")
            print(f"📊 Tamaño del .exe: {size_mb:.1f} MB "
                  f"| carpeta completa: {total_mb:.1f} MB")

            print("\n📋 INSTRUCCIONES PARA EL USUARIO:")
            print(f"1. Comprimir la CARPETA dist/{NOMBRE_APP} en un .zip y enviarla")
            print("2. El usuario descomprime y ejecuta el .exe de adentro")
            print("   (el .exe NO funciona si se saca solo de la carpeta)")
            print("3. No requiere tener Python instalado")
            print("4. Botón 'Examinar...' → elegir el archivo de Creación de Reservas")
            print("5. El resultado se guarda en una subcarpeta 'Output'")
            print("   junto al archivo elegido. El original NO se modifica.")
        else:
            print("❌ Error: No se encontró el ejecutable después de la compilación")

    except subprocess.CalledProcessError as e:
        print(f"❌ Error creando ejecutable: {e}")
        print("\n💡 Sugerencias:")
        print("   - Verifica que estés en la carpeta donde están los archivos .py")
        print("   - Ejecuta: pip install pyinstaller")
        print("   - Si el error menciona 'openpyxl', ejecuta: pip install openpyxl")
        print("   - Si el error menciona 'pandas', ejecuta: pip install pandas")


def check_source_files():
    """Verificar que todos los archivos fuente necesarios existen"""
    print("🔍 Verificando archivos fuente...")
    todos_ok = True
    for archivo in ARCHIVOS_REQUERIDOS:
        if os.path.exists(archivo):
            print(f"   ✅ {archivo}")
        else:
            print(f"   ❌ {archivo} — NO ENCONTRADO")
            todos_ok = False
    return todos_ok


def main():
    print("🔨 CONSTRUCCIÓN DE EJECUTABLE - RDA PÓLIZAS D&A")
    print("=" * 60)

    if not check_source_files():
        print("\n❌ Faltan archivos fuente. Asegúrate de ejecutar este script")
        print("   desde la carpeta donde están todos los archivos .py del proyecto.")
        return

    print()
    install_pyinstaller()
    print()
    create_executable()
    print("\n🎉 ¡Proceso completado!")


if __name__ == "__main__":
    main()
