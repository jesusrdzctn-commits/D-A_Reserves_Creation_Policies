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
    "Reservas_Manuales.py",
    "Create_Poliza_SAP.py",
    "Create_Poliza_Batch.py",
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
    """pywin32 lets Create_Poliza_SAP/Batch drive Excel; without it the 'PÓLIZA SAP' button fails at run time, not at build time."""
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
        # --onedir: a FOLDER with the .exe inside; starts faster than --onefile and trips corporate antivirus less (zip it to send).
        "--onedir",
        "--windowed",
        f"--name={NOMBRE_APP}",
        "--icon=NONE",
        "--clean",
        "--noconfirm",

        # Modules PyInstaller does not detect on its own.
        # Process 1 ('Variaciones') — writes the .xlsx with openpyxl:
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.cell._writer",
        "--hidden-import=pandas",

        # Process 2 ('PÓLIZA SAP') and the comparison pivots drive Excel through COM; utils.abrir_excel() imports
        # win32com INSIDE the function (so the module still imports without Excel), which PyInstaller cannot trace.
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=pywintypes",

        # Consolidacion.py imports Comparativos inside the function for the same reason; declared in case the analysis misses it.
        "--hidden-import=Comparativos",

        # Consolidacion.py also imports Reservas_Manuales inside the function (a top-level import would be circular);
        # declared here so the build does not depend on the controller importing it.
        "--hidden-import=Reservas_Manuales",
        "--hidden-import=openpyxl.formula.translate",

        # Entry point
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
