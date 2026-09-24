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
    "Comparativos.py",
    "Poliza_SAP.py",
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


def check_pywin32():
    """
    pywin32 es lo que le permite a Poliza_SAP.py manejar Excel (el archivo
    WHSL es .xlsb y sólo Excel lo abre de forma nativa). Sin esto, el botón
    de 'PÓLIZA SAP' truena en tiempo de ejecución, no al compilar.
    """
    try:
        import win32com.client  # noqa: F401
        print("✅ pywin32 ya está instalado")
        return True
    except ImportError:
        print("⚠️  pywin32 NO está instalado — el proceso de 'PÓLIZA SAP' no funcionará")
        print("    Instálalo con:  pip install pywin32")
        return False


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
        # Proceso 1 ('Variaciones') — escribe el .xlsx con openpyxl:
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.cell._writer",
        "--hidden-import=pandas",

        # Proceso 2 ('PÓLIZA SAP') y las tablas dinámicas comparativas — los dos
        # manejan Excel por COM. PyInstaller no rastrea estos imports porque
        # utils.abrir_excel() los hace DENTRO de la función (a propósito: así el
        # módulo se puede importar en una máquina sin Excel y sólo truena si de
        # verdad se usa el botón).
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=pywintypes",

        # Consolidacion.py importa Comparativos DENTRO de la función, por la
        # misma razón. Se declara por si el análisis de PyInstaller no lo ve.
        "--hidden-import=Comparativos",

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
            print("4. SÍ requiere tener Excel instalado para el botón de 'PÓLIZA SAP'")
            print("5. Botón 'Examinar...' de cada bloque → elegir su archivo")
            print("6. Los resultados se guardan en una subcarpeta 'Output'")
            print("   junto a cada archivo elegido. Los originales NO se modifican.")
        else:
            print("❌ Error: No se encontró el ejecutable después de la compilación")

    except subprocess.CalledProcessError as e:
        print(f"❌ Error creando ejecutable: {e}")
        print("\n💡 Sugerencias:")
        print("   - Verifica que estés en la carpeta donde están los archivos .py")
        print("   - Ejecuta: pip install pyinstaller")
        print("   - Si el error menciona 'openpyxl', ejecuta: pip install openpyxl")
        print("   - Si el error menciona 'pandas', ejecuta: pip install pandas")
        print("   - Si el error menciona 'win32com', ejecuta: pip install pywin32")


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
    check_pywin32()
    print()
    create_executable()
    print("\n🎉 ¡Proceso completado!")


if __name__ == "__main__":
    main()
