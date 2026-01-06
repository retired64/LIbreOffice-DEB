#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
import re
import os
import tarfile
import subprocess
from tqdm import tqdm

# Colores para terminal
class colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'

# Carpeta donde se guardarán los archivos
DOWNLOAD_DIR = "LibreOffice"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# URL base de descargas
BASE_URL = "https://download.documentfoundation.org/libreoffice/stable/"

def get_latest_version():
    """Obtiene la última versión estable de LibreOffice desde la web."""
    response = requests.get(BASE_URL)
    if response.status_code != 200:
        raise Exception("No se pudo acceder a la página de versiones")

    versions = re.findall(r'\d+\.\d+\.\d+/', response.text)
    versions = [v.strip("/") for v in versions]
    versions.sort(key=lambda s: list(map(int, s.split('.'))))
    return versions[-1]

def download_file(url, folder):
    """Descarga un archivo mostrando barra de progreso."""
    local_filename = os.path.join(folder, url.split('/')[-1])

    if os.path.exists(local_filename):
        print(f"{colors.WARNING}Archivo ya existe, se omitirá:{colors.ENDC} {local_filename}")
        return local_filename

    response = requests.get(url, stream=True)
    total_size = int(response.headers.get('content-length', 0))
    chunk_size = 1024

    with open(local_filename, 'wb') as f, tqdm(
        total=total_size,
        unit='B',
        unit_scale=True,
        desc=os.path.basename(local_filename),
        ncols=80
    ) as bar:
        for chunk in response.iter_content(chunk_size=chunk_size):
            f.write(chunk)
            bar.update(len(chunk))

    print(f"{colors.OKGREEN}Descargado:{colors.ENDC} {local_filename}")
    return local_filename

def extract_tar_gz(file_path, folder):
    """Extrae archivos .tar.gz en la carpeta indicada."""
    if not tarfile.is_tarfile(file_path):
        print(f"{colors.FAIL}No es un archivo tar válido:{colors.ENDC} {file_path}")
        return None

    with tarfile.open(file_path, "r:gz") as tar:
        tar.extractall(path=folder)

    base_name = os.path.basename(file_path).replace(".tar.gz", "")
    extracted_folder = os.path.join(folder, base_name)

    for item in os.listdir(folder):
        if item.startswith(base_name.split("_")[0]):
            extracted_folder = os.path.join(folder, item)
            break

    print(f"{colors.OKGREEN}Descomprimido:{colors.ENDC} {file_path}")
    return extracted_folder

def install_debs(debs_folder):
    """Instala todos los .deb dentro de la carpeta DEBS."""
    debs_path = os.path.join(debs_folder, "DEBS")

    if not os.path.exists(debs_path):
        print(f"{colors.FAIL}No se encontró carpeta DEBS en {debs_folder}{colors.ENDC}")
        return

    deb_files = sorted(
        os.path.join(debs_path, f)
        for f in os.listdir(debs_path)
        if f.endswith(".deb")
    )

    if not deb_files:
        print(f"{colors.WARNING}No hay archivos .deb para instalar en {debs_path}{colors.ENDC}")
        return

    print(f"{colors.OKBLUE}Instalando paquetes desde {debs_path}...{colors.ENDC}")

    try:
        subprocess.run(["sudo", "dpkg", "-i"] + deb_files, check=True)
        print(f"{colors.OKGREEN}Instalación completada.{colors.ENDC}")
    except subprocess.CalledProcessError:
        print(f"{colors.WARNING}Dependencias faltantes, corrigiendo...{colors.ENDC}")
        subprocess.run(["sudo", "apt-get", "-f", "install", "-y"], check=True)
        subprocess.run(["sudo", "dpkg", "-i"] + deb_files, check=True)
        print(f"{colors.OKGREEN}Instalación completada.{colors.ENDC}")

def main():
    print(f"{colors.HEADER}Buscando la última versión estable de LibreOffice...{colors.ENDC}")
    version = get_latest_version()
    print(f"{colors.OKBLUE}Última versión detectada:{colors.ENDC} {version}\n")

    confirm = input(
        f"{colors.HEADER}¿Deseas descargar, descomprimir e instalar LibreOffice {version}? [Y/S/Yes/Si]: {colors.ENDC}"
    ).strip().lower()

    if confirm not in ["y", "s", "si", "yes"]:
        print(f"{colors.WARNING}Operación cancelada por el usuario.{colors.ENDC}")
        return

    urls_ordered = [
        # 1️⃣ BASE
        f"{BASE_URL}{version}/deb/x86_64/LibreOffice_{version}_Linux_x86-64_deb.tar.gz",

        # 2️⃣ HELP PACK (igual que tu historial)
        f"{BASE_URL}{version}/deb/x86_64/LibreOffice_{version}_Linux_x86-64_deb_helppack_es.tar.gz",

        # 3️⃣ LANGUAGE PACK
        f"{BASE_URL}{version}/deb/x86_64/LibreOffice_{version}_Linux_x86-64_deb_langpack_es.tar.gz",
    ]

    for url in urls_ordered:
        try:
            downloaded = download_file(url, DOWNLOAD_DIR)
            extracted = extract_tar_gz(downloaded, DOWNLOAD_DIR)
            if extracted:
                install_debs(extracted)
        except Exception as e:
            print(f"{colors.FAIL}Error con {url}: {e}{colors.ENDC}")

    print(f"\n{colors.OKGREEN}LibreOffice instalado siguiendo el orden correcto ✔{colors.ENDC}")

if __name__ == "__main__":
    main()
