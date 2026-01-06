#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
LibreOffice Automatic Installer (Debian/Ubuntu)
------------------------------------------------
Compatible con PyInstaller
Rutas dinámicas (script/ejecutable)
Logging completo
Manejo robusto de errores
"""

import os
import re
import sys
import tarfile
import subprocess
import requests
import logging
from tqdm import tqdm
from typing import Optional
from pathlib import Path

# =========================
# RUTAS DINÁMICAS (PyInstaller Compatible)
# =========================

def get_base_path() -> Path:
    """
    Retorna la ruta base donde se ejecuta el script/binario.
    - En PyInstaller: directorio del ejecutable
    - En Python normal: directorio del script .py
    """
    if getattr(sys, 'frozen', False):
        # Ejecutable compilado con PyInstaller
        return Path(sys.executable).parent.resolve()
    else:
        # Script Python normal
        return Path(__file__).parent.resolve()

# =========================
# CONFIGURACIÓN
# =========================

BASE_PATH = get_base_path()
BASE_URL = "https://download.documentfoundation.org/libreoffice/stable/"
DOWNLOAD_DIR = BASE_PATH / "LibreOffice"
LOG_FILE = BASE_PATH / "libreoffice_installer.log"
HTTP_TIMEOUT = 30

# Crear carpeta de descargas
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# =========================
# COLORES
# =========================

class colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'

# =========================
# LOGGING
# =========================

logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def log_and_print(message: str, level: str = "info") -> None:
    """Registra en log y muestra en consola"""
    if level == "info":
        logging.info(message)
    elif level == "warning":
        logging.warning(message)
    elif level == "error":
        logging.error(message)
    elif level == "critical":
        logging.critical(message)

# =========================
# VALIDACIONES DEL SISTEMA
# =========================

def check_system_dependencies() -> None:
    """Verifica que el sistema sea compatible (Debian/Ubuntu)"""
    if not os.path.exists("/usr/bin/dpkg"):
        print(f"{colors.FAIL}❌ Este script requiere un sistema Debian/Ubuntu con dpkg{colors.ENDC}")
        log_and_print("Sistema no compatible: dpkg no encontrado", "critical")
        raise SystemExit(1)
    log_and_print("✅ Sistema compatible verificado")

def check_sudo_available() -> None:
    """Informa sobre la necesidad de permisos sudo"""
    if os.geteuid() != 0:
        print(f"{colors.WARNING}⚠️  Se solicitará sudo durante la instalación de paquetes{colors.ENDC}")
        log_and_print("Usuario sin privilegios root, se requerirá sudo", "warning")

# =========================
# 🔍 DETECCIÓN DE VERSIÓN
# =========================

def get_latest_version() -> str:
    """
    Obtiene la última versión estable de LibreOffice desde el servidor oficial.
    Usa regex para extraer solo directorios de versión válidos.
    """
    log_and_print(f"Consultando versiones en: {BASE_URL}")
    
    try:
        response = requests.get(BASE_URL, timeout=HTTP_TIMEOUT)
        response.raise_for_status()
    except requests.RequestException as e:
        log_and_print(f"Error al acceder a {BASE_URL}: {e}", "error")
        raise RuntimeError(f"No se pudo acceder al servidor: {e}")

    # Buscar solo enlaces válidos de versión (ej: href="7.6.4/")
    versions = re.findall(r'href="(\d+\.\d+\.\d+)/"', response.text)
    
    if not versions:
        log_and_print("No se encontraron versiones válidas en la página", "error")
        raise RuntimeError("No se encontraron versiones disponibles")

    # Ordenar por versión semántica (ej: 7.6.4 > 7.6.3)
    versions.sort(key=lambda s: tuple(map(int, s.split("."))))
    latest = versions[-1]
    
    log_and_print(f"Última versión detectada: {latest}")
    return latest

# =========================
# ⬇️  DESCARGA DE ARCHIVOS
# =========================

def download_file(url: str, dest: Path) -> str:
    """
    Descarga un archivo desde una URL con barra de progreso.
    Si el archivo ya existe, omite la descarga.
    """
    filename = dest / url.split("/")[-1]

    if filename.exists():
        print(f"{colors.WARNING}⏭️  Ya existe (omitiendo):{colors.ENDC} {filename.name}")
        log_and_print(f"Archivo ya existe: {filename}")
        return str(filename)

    log_and_print(f"Descargando: {url}")
    
    try:
        response = requests.get(url, stream=True, timeout=HTTP_TIMEOUT)
        response.raise_for_status()
    except requests.RequestException as e:
        log_and_print(f"Error al descargar {url}: {e}", "error")
        raise RuntimeError(f"Descarga fallida: {e}")

    total_size = int(response.headers.get("content-length", 0))
    chunk_size = 8192  # 8KB chunks para mejor rendimiento

    with open(filename, "wb") as f, tqdm(
        total=total_size,
        unit="B",
        unit_scale=True,
        unit_divisor=1024,
        desc=filename.name,
        ncols=80,
        colour="cyan"
    ) as bar:
        for chunk in response.iter_content(chunk_size):
            f.write(chunk)
            bar.update(len(chunk))

    print(f"{colors.OKGREEN}✅ Descargado:{colors.ENDC} {filename.name}")
    log_and_print(f"Descarga completada: {filename}")
    return str(filename)

# =========================
# 📦 EXTRACCIÓN DE ARCHIVOS
# =========================

def extract_tar_gz(file_path: str, dest: Path) -> Optional[Path]:
    """
    Extrae un archivo .tar.gz y retorna la ruta del directorio raíz extraído.
    Usa el listado del tarball para encontrar el directorio base correctamente.
    """
    file_path_obj = Path(file_path)
    
    if not tarfile.is_tarfile(file_path):
        log_and_print(f"No es un archivo tar válido: {file_path}", "error")
        raise RuntimeError(f"Archivo tar inválido: {file_path}")

    log_and_print(f"Extrayendo: {file_path_obj.name}")

    try:
        with tarfile.open(file_path, "r:gz") as tar:
            # Obtener el directorio raíz del tarball
            members = tar.getmembers()
            if not members:
                raise RuntimeError("El archivo tar está vacío")
            
            # El primer miembro debería contener el directorio raíz
            root_dir = members[0].name.split("/")[0]
            
            # Extraer todo
            tar.extractall(path=dest)
    except tarfile.TarError as e:
        log_and_print(f"Error al extraer {file_path}: {e}", "error")
        raise RuntimeError(f"Extracción fallida: {e}")

    extracted_path = dest / root_dir
    
    if not extracted_path.exists():
        log_and_print(f"Directorio extraído no encontrado: {extracted_path}", "error")
        raise RuntimeError(f"No se pudo localizar el directorio extraído: {extracted_path}")

    print(f"{colors.OKGREEN}📂 Descomprimido:{colors.ENDC} {file_path_obj.name} → {root_dir}/")
    log_and_print(f"Extracción completada: {extracted_path}")
    return extracted_path

# =========================
# 💿 INSTALACIÓN DE PAQUETES
# =========================

def install_debs(package_dir: Path) -> None:
    """
    Instala todos los archivos .deb dentro de la carpeta DEBS.
    Maneja automáticamente dependencias faltantes con apt-get.
    """
    debs_path = package_dir / "DEBS"

    if not debs_path.is_dir():
        log_and_print(f"Carpeta DEBS no encontrada en: {package_dir}", "error")
        raise RuntimeError(f"No existe el directorio DEBS en {package_dir}")

    # Buscar todos los .deb y ordenarlos
    debs = sorted(debs_path.glob("*.deb"))

    if not debs:
        log_and_print(f"No hay paquetes .deb en: {debs_path}", "warning")
        raise RuntimeError(f"No se encontraron archivos .deb en {debs_path}")

    print(f"\n{colors.OKBLUE}📦 Instalando {len(debs)} paquetes desde:{colors.ENDC} {debs_path.name}")
    log_and_print(f"Iniciando instalación de {len(debs)} paquetes")

    # Convertir paths a strings para subprocess
    deb_files = [str(deb) for deb in debs]

    try:
        # Intento 1: Instalación directa
        subprocess.run(
            ["sudo", "dpkg", "-i"] + deb_files,
            check=True,
            capture_output=True,
            text=True
        )
        print(f"{colors.OKGREEN}✅ Instalación completada exitosamente{colors.ENDC}")
        log_and_print("Instalación completada sin errores")
        
    except subprocess.CalledProcessError as e:
        # Si falló por dependencias, corregir con apt-get
        print(f"{colors.WARNING}⚠️  Corrigiendo dependencias faltantes...{colors.ENDC}")
        log_and_print("Dependencias faltantes, ejecutando apt-get -f install", "warning")
        
        try:
            # Corregir dependencias
            subprocess.run(
                ["sudo", "apt-get", "-f", "install", "-y"],
                check=True,
                capture_output=True,
                text=True
            )
            
            # Intento 2: Reinstalar después de corregir dependencias
            subprocess.run(
                ["sudo", "dpkg", "-i"] + deb_files,
                check=True,
                capture_output=True,
                text=True
            )
            
            print(f"{colors.OKGREEN}✅ Instalación completada (dependencias corregidas){colors.ENDC}")
            log_and_print("Instalación completada tras corrección de dependencias")
            
        except subprocess.CalledProcessError as e2:
            log_and_print(f"Error crítico durante instalación: {e2.stderr}", "critical")
            raise RuntimeError(f"Instalación fallida: {e2.stderr}")

# =========================
# 🚀 FUNCIÓN PRINCIPAL
# =========================

def main() -> None:
    """
    Función principal que orquesta todo el proceso:
    1. Validaciones del sistema
    2. Detección de versión
    3. Descarga de paquetes
    4. Extracción
    5. Instalación
    """
    print(f"{colors.HEADER}")
    print("=" * 70)
    print("  LibreOffice Automatic Installer for Debian/Ubuntu")
    print("=" * 70)
    print(f"{colors.ENDC}\n")
    
    log_and_print("=" * 50)
    log_and_print("Iniciando LibreOffice Installer")
    log_and_print(f"Directorio base: {BASE_PATH}")
    log_and_print(f"Directorio de descargas: {DOWNLOAD_DIR}")
    log_and_print("=" * 50)

    # 1. Validaciones
    try:
        check_system_dependencies()
        check_sudo_available()
    except SystemExit:
        return

    # 2. Obtener última versión
    print(f"{colors.HEADER}🔍 Buscando última versión estable...{colors.ENDC}")
    try:
        version = get_latest_version()
        print(f"{colors.OKBLUE}📌 Versión detectada:{colors.ENDC} {version}\n")
    except RuntimeError as e:
        print(f"{colors.FAIL}❌ Error:{colors.ENDC} {e}")
        return

    # 3. Confirmación del usuario
    print(f"{colors.WARNING}Este proceso descargará e instalará LibreOffice {version}{colors.ENDC}")
    print(f"Carpeta de destino: {DOWNLOAD_DIR}")
    confirm = input(
        f"\n{colors.HEADER}¿Continuar? [Y/S/Yes/Si]:{colors.ENDC} "
    ).strip().lower()

    if confirm not in {"y", "s", "si", "yes"}:
        print(f"{colors.WARNING}❌ Operación cancelada por el usuario{colors.ENDC}")
        log_and_print("Usuario canceló la operación")
        return

    # 4. URLs de descarga (orden específico)
    urls = [
        # 1️⃣ Paquete BASE (obligatorio)
        f"{BASE_URL}{version}/deb/x86_64/LibreOffice_{version}_Linux_x86-64_deb.tar.gz",
        
        # 2️⃣ HELPPACK en español (ayuda integrada)
        f"{BASE_URL}{version}/deb/x86_64/LibreOffice_{version}_Linux_x86-64_deb_helppack_es.tar.gz",
        
        # 3️⃣ LANGPACK en español (interfaz traducida)
        f"{BASE_URL}{version}/deb/x86_64/LibreOffice_{version}_Linux_x86-64_deb_langpack_es.tar.gz",
    ]

    package_names = ["BASE", "HELPPACK (ES)", "LANGPACK (ES)"]

    # 5. Proceso de descarga e instalación
    print(f"\n{colors.OKBLUE}{'=' * 70}{colors.ENDC}")
    print(f"{colors.OKBLUE}Iniciando descarga e instalación de 3 paquetes{colors.ENDC}")
    print(f"{colors.OKBLUE}{'=' * 70}{colors.ENDC}\n")

    for idx, (url, pkg_name) in enumerate(zip(urls, package_names), 1):
        print(f"\n{colors.HEADER}[{idx}/3] Procesando: {pkg_name}{colors.ENDC}")
        log_and_print(f"Procesando paquete {idx}/3: {pkg_name}")
        
        try:
            # Descargar
            tarball = download_file(url, DOWNLOAD_DIR)
            
            # Extraer
            extracted = extract_tar_gz(tarball, DOWNLOAD_DIR)
            
            # Instalar
            install_debs(extracted)
            
            print(f"{colors.OKGREEN}✅ {pkg_name} instalado correctamente{colors.ENDC}")
            
        except Exception as e:
            print(f"\n{colors.FAIL}❌ ERROR en {pkg_name}:{colors.ENDC} {e}")
            log_and_print(f"Error crítico durante {pkg_name}: {e}", "critical")
            print(f"\n{colors.WARNING}La instalación se detuvo. Revisa el log:{colors.ENDC} {LOG_FILE}")
            return

    # 6. Finalización exitosa
    print(f"\n{colors.OKGREEN}{'=' * 70}{colors.ENDC}")
    print(f"{colors.OKGREEN}✅ LibreOffice {version} instalado correctamente{colors.ENDC}")
    print(f"{colors.OKGREEN}{'=' * 70}{colors.ENDC}\n")
    
    print(f"📁 Archivos descargados en: {DOWNLOAD_DIR}")
    print(f"📝 Log disponible en: {LOG_FILE}")
    print(f"\n{colors.OKBLUE}Para ejecutar LibreOffice, usa:{colors.ENDC} libreoffice")
    
    log_and_print("=" * 50)
    log_and_print(f"Instalación de LibreOffice {version} completada exitosamente")
    log_and_print("=" * 50)

# =========================
# 🎬 PUNTO DE ENTRADA
# =========================

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n{colors.WARNING}⚠️  Instalación interrumpida por el usuario{colors.ENDC}")
        log_and_print("Instalación interrumpida por KeyboardInterrupt", "warning")
        sys.exit(1)
    except Exception as e:
        print(f"\n{colors.FAIL}❌ Error inesperado:{colors.ENDC} {e}")
        log_and_print(f"Error inesperado no capturado: {e}", "critical")
        sys.exit(1)
