import csv
import os
import json
import subprocess
import threading
import sys
import time
from tqdm import tqdm
import itertools


def ensure_absolute_path(path):
    """
    Ensure the given path is an absolute path.

    Parameters:
        path (str): The file path to check.

    Returns:
        str: Absolute path of the given file.
    """
    if not os.path.isabs(path):
        return os.path.abspath(path)
    return path

def txt_to_csv(txt_path, csv_path):
    """
    Convert a TXT file to a CSV file.

    Parameters:
        txt_path (str): Path to the input TXT file.
        csv_path (str): Path to the output CSV file.
    """
    ensure_absolute_path(txt_path)
    ensure_absolute_path(csv_path)

    with open(txt_path, "r") as txt_file:
        lines = txt_file.readlines()

    data = []
    for line in tqdm(lines, desc="Converting TXT to CSV", unit="rows"):
        row = line.strip().split(",")
        data.append(row)

    with open(csv_path, "w", newline="") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerows(data)

    print(f"Converted TXT to CSV: {csv_path}")

def csv_to_txt(csv_path, txt_path):
    """
    Convert a CSV file to a TXT file.

    Parameters:
        csv_path (str): Path to the input CSV file.
        txt_path (str): Path to the output TXT file.
    """
    ensure_absolute_path(csv_path)
    ensure_absolute_path(txt_path)

    with open(csv_path, "r") as csv_file:
        reader = csv.reader(csv_file)
        data = list(reader)

    with open(txt_path, "w") as txt_file:
        for row in tqdm(data, desc="Converting CSV to TXT", unit="rows"):
            txt_file.write(",".join(row) + "\n")

    print(f"Converted CSV to TXT: {txt_path}")

def csv_to_json(csv_path, json_path):
    """
    Convert a CSV file to a TXT file.

    Parameters:
        csv_path (str): Path to the input CSV file.
        json_path (str): Path to the output TXT file.
    """
    ensure_absolute_path(csv_path)
    ensure_absolute_path(json_path)

    with open(csv_path, "r") as csv_file:
        reader = csv.DictReader(csv_file)
        data = []
        for row in tqdm(reader, desc="Converting CSV to JSON", unit="rows"):
            data.append(row)
        
    with open(json_path, "w") as json_file:
        json.dump(data, json_file, indent=4)

    print(f"Converted CSV to JSON: {json_path}")

def json_to_csv(json_path, csv_path):
    """
    Convert a JSON file to a CSV file.

    Parameters:
        json_path (str): Path to the input JSON file.
        csv_path (str): Path to the output CSV file.
    """
    ensure_absolute_path(json_path)
    ensure_absolute_path(csv_path)

    with open(json_path, "r") as json_file:
        data = json.load(json_file)

    with open(csv_path, "w", newline="") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=data[0].keys())
        writer.writeheader()
        for row in tqdm(data, desc="Converting JSON to CSV", unit="rows"):
            writer.writerow(row)

def is_dism_available():
    try:
        subprocess.run(
            ["dism", "/?"], check=True, text=True, capture_output=True
        )
        return True
    except subprocess.CalledProcessError as e:
        print(f"DISM returned an error: {e}")
        return False
    except FileNotFoundError as e:
        print(f"DISM not found: {e}")
        return False

def wimlib_esd_to_wim(esd_path, wim_path, source_index=1):
    try:
        subprocess.run(["wimlib-imagex", "export", esd_path, str(source_index), wim_path, "--compress=max"], check=True)
        print(f"Successfully converted ESD to WIM using wimlib: {wim_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error: {e.stderr}")

def wimlib_wim_to_esd(wim_path, esd_path, source_index=1):
    try:
        # Menggunakan wimlib-imagex untuk mengekspor image dari WIM menjadi ESD
        subprocess.run(["wimlib-imagex", "export", wim_path, str(source_index), esd_path, "--solid"], check=True)
        print(f"Successfully converted WIM to ESD: {esd_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error: {e.stderr}")

def esd_to_wim(esd_path, wim_path, source_index=1):
    """
    Convert ESD file to WIM file using DISM (with admin rights required).
    If DISM is unavailable, use wimlib as an alternative.
    """
    ensure_absolute_path(esd_path)
    ensure_absolute_path(wim_path)

    if is_dism_available():
        dism_command = [
            "dism",
            "/export-image",
            f"/sourceimagefile:{esd_path}",
            f"/sourceindex:{source_index}",
            f"/destinationimagefile:{wim_path}",
            "/compress:max",
            "/checkintegrity",
        ]
        print("Starting DISM process...")

        try:
            result = subprocess.run(
                dism_command, check=True, text=True, capture_output=True
            )
            print(f"\nSuccess: {result.stdout}")
        except subprocess.CalledProcessError as e:
            print(f"\nError: {e.stderr}")
            print("DISM failed, switching to wimlib...")
            wimlib_esd_to_wim(esd_path, wim_path, source_index)
    else:
        print("DISM is not available. Trying alternative method with wimlib...")
        wimlib_esd_to_wim(esd_path, wim_path, source_index)

def wim_to_esd(wim_path, esd_path, source_index=1):
    """
    Convert WIM file to ESD file using DISM (with admin rights required).
    If DISM is unavailable, use wimlib as an alternative.
    """
    ensure_absolute_path(wim_path)
    ensure_absolute_path(esd_path)

    if is_dism_available():
        dism_command = [
            "dism",
            "/export-image",
            f"/sourceimagefile:{wim_path}",
            f"/sourceindex:{source_index}",
            f"/destinationimagefile:{esd_path}",
            "/compress:recovery",
            "/checkintegrity",
        ]
        print("Starting DISM process...")

        try:
            result = subprocess.run(
                dism_command, check=True, text=True, capture_output=True
            )
            print(f"\nSuccess: {result.stdout}")
        except subprocess.CalledProcessError as e:
            print(f"\nError: {e.stderr}")
            print("DISM failed, switching to wimlib...")
            wimlib_wim_to_esd(wim_path, esd_path, source_index)
    else:
        print("DISM is not available. Trying alternative method with wimlib...")
        wimlib_wim_to_esd(wim_path, esd_path, source_index)