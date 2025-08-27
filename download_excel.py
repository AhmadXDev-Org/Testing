#!/usr/bin/env python3
import requests
import sys

def download_file(url, filename):
    """Download a file from URL and save it locally"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, stream=True)
        response.raise_for_status()
        
        with open(filename, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        print(f"Successfully downloaded {filename}")
        return True
    except Exception as e:
        print(f"Error downloading file: {e}")
        return False

if __name__ == "__main__":
    url = "https://github.com/user-attachments/files/22005679/Fruits.xlsx"
    filename = "Fruits.xlsx"
    download_file(url, filename)