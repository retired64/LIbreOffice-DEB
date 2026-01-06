#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
LibreOffice Automatic Installer (Debian/Ubuntu)
------------------------------------------------
- Detecta la última versión estable
- Descarga paquetes oficiales
- Instala en orden correcto:
  1) Base
  2) HelpPack (ES)
  3) LangPack (ES)
- Manejo robusto de errores
- Logging a archivo
"""

import os
import re
import tarfile
import subprocess
import requests
import logging
from tqdm import tqdm
from typing import Optional

# =========================
# CONFIGURACIÓN
# =========================

BASE_URL = "https://download.documentfoundation.org/libreoffice/stable/"
DOWNLOAD_DIR = "LibreOffice"
LOG_FILE = "libreoffice_installer.log"
HTTP_TIMEOUT = 30

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
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# =========================
# VALIDACIONES
# =========================

def check_system_dependencies() -> None:
    if not os.path.exists("/usr/bin/dpkg"):
        print(f"{colors.FAIL}Este script requiere un sistema Debian/Ubuntu{colors.ENDC}")
        raise SystemExit(1)

def check_sudo_available() -> None:
    if os.geteuid() != 0:
        print(f"{colors.WARNING}Se solicitará sudo durante la instalación{colors.ENDC}")

# =========================
# VERSIONES
# =========================

def get_latest_version() -> str:
    response = requests.get(BASE_URL, timeout=HTTP_TIMEOUT)
    response.raise_for_status()

    versions = re.findall(r'href="(\d+\.\d+\.\d+)/"', response.text)
    if not versions:
        raise RuntimeError("No se encontraron versiones válidas")

    versions.sort(key=lambda s: tuple(map(int, s.split("."))))
    return versions[-1]

# =========================
# DESCARGA
# =========================

def download_file(url: str, dest: str) -> str:
    filename = os.path.join(dest, url.split("/")[-1])

    if os.path.exists(filename):
        print(f"{colors.WARNING}Ya existe:{colors.ENDC} {filename}")
        return filename

    logging.info(f"Descargando {url}")

    response = requests.get(url, stream=True, timeout=HTTP_TIMEOUT)
    response.raise_for_status()

    total_size = int(response.headers.get("content-length", 0))
    chunk_size = 1024

    with open(filename, "wb") as f, tqdm(
        total=total_size,
        unit="B",
        unit_scale=True,
        desc=os.path.basename(filename),
        ncols=80
    ) as bar:
        for chunk in response.iter_content(chunk_size):
            f.write(chunk)
            bar.update(len(chunk))

    print(f"{colors.OKGREEN}Descargado:{colors.ENDC} {filename}")
    return filename

# =========================
# EXTRACCIÓN SEGURA
# =========================

def extract_tar_gz(file_path: str, dest: str) -> Optional[str]:
    if not tarfile.is_tarfile(file_path):
        raise RuntimeError(f"No es un tar válido: {file_path}")

    with tarfile.open(file_path, "r:gz") as tar:
        members = tar.getmembers()
        root_dir = members[0].name.split("/")[0] if members else None
        tar.extractall(path=dest)

    if not root_dir:
        raise RuntimeError("No se pudo determinar el directorio raíz")

    extracted_path = os.path.join(dest, root_dir)
    logging.info(f"Extraído {file_path} → {extracted_path}")

    print(f"{colors.OKGREEN}Descomprimido:{colors.ENDC} {file_path}")
    return extracted_path

# =========================
# INSTALACIÓN
# =========================

def install_debs(package_dir: str) -> None:
    debs_path = os.path.join(package_dir, "DEBS")

    if not os.path.isdir(debs_path):
        raise RuntimeError(f"No existe DEBS en {package_dir}")

    debs = sorted(
        os.path.join(debs_path, f)
        for f in os.listdir(debs_path)
        if f.endswith(".deb")
    )

    if not debs:
        raise RuntimeError(f"No hay paquetes .deb en {debs_path}")

    print(f"{colors.OKBLUE}Instalando desde {debs_path}{colors.ENDC}")
    logging.info(f"Instalando paquetes en {debs_path}")

    try:
        subprocess.run(["sudo", "dpkg", "-i"] + debs, check=True)
    except subprocess.CalledProcessError:
        print(f"{colors.WARNING}Corrigiendo dependencias...{colors.ENDC}")
        subprocess.run(["sudo", "apt-get", "-f", "install", "-y"], check=True)
        subprocess.run(["sudo", "dpkg", "-i"] + debs, check=True)

    print(f"{colors.OKGREEN}Instalación completada{colors.ENDC}")

# =========================
# MAIN
# =========================

def main() -> None:
    check_system_dependencies()
    check_sudo_available()

    print(f"{colors.HEADER}Buscando última versión estable de LibreOffice...{colors.ENDC}")
    version = get_latest_version()
    print(f"{colors.OKBLUE}Versión detectada:{colors.ENDC} {version}\n")

    confirm = input(
        f"{colors.HEADER}¿Instalar LibreOffice {version}? [Y/S/Yes/Si]: {colors.ENDC}"
    ).strip().lower()

    if confirm not in {"y", "s", "si", "yes"}:
        print(f"{colors.WARNING}Operación cancelada{colors.ENDC}")
        return

    urls = [
        # 1️⃣ BASE
        f"{BASE_URL}{version}/deb/x86_64/LibreOffice_{version}_Linux_x86-64_deb.tar.gz",
        # 2️⃣ HELPPACK (orden idéntico a tu historial)
        f"{BASE_URL}{version}/deb/x86_64/LibreOffice_{version}_Linux_x86-64_deb_helppack_es.tar.gz",
        # 3️⃣ LANGPACK
        f"{BASE_URL}{version}/deb/x86_64/LibreOffice_{version}_Linux_x86-64_deb_langpack_es.tar.gz",
    ]

    for url in urls:
        try:
            tarball = download_file(url, DOWNLOAD_DIR)
            extracted = extract_tar_gz(tarball, DOWNLOAD_DIR)
            install_debs(extracted)
        except Exception as e:
            logging.exception("Error durante instalación")
            print(f"{colors.FAIL}Error:{colors.ENDC} {e}")
            return

    print(f"\n{colors.OKGREEN}LibreOffice instalado correctamente ✔{colors.ENDC}")
    print(f"Log disponible en: {LOG_FILE}")

# =========================

if __name__ == "__main__":
    main()
