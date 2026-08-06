import os
import time
import json
import hashlib
import requests
import tempfile
from pathlib import Path
from pdf2docx import Converter
import sys

API_BASE = "https://api.opencaselist.com/v1"
SESSION = requests.Session()
SESSION.cookies.set("caselist_token", "7197e95921e0982ac01651ae3045ff26", domain=".opencaselist.com")
CACHE_DIR = Path("caselist_output/cache")

def run():
    metas = json.loads(Path("caselist_output/last_metas.json").read_text())
    
    missing = []
    for meta in metas:
        path = str(meta.get("opensource") or "")
        if not path:
            continue
        is_pdf = path.lower().endswith(".pdf")
            
        key = hashlib.md5(path.encode()).hexdigest()
        cached = CACHE_DIR / f"{key}.docx"
        if not cached.exists() or cached.stat().st_size == 0:
            missing.append(path)
            
    print(f"Found {len(missing)} missing files to download from the unreached schools.")
    
    for path in missing:
        is_pdf = path.lower().endswith(".pdf")
        key = hashlib.md5(path.encode()).hexdigest()
        cached = CACHE_DIR / f"{key}.docx"
        
        print(f"Downloading: {Path(path).name}")
        
        success = False
        for attempt in range(6):
            try:
                r = SESSION.get(f"{API_BASE}/download", params={"path": path}, timeout=30)
                if r.status_code == 429:
                    wait = (2 ** attempt) + 2
                    print(f"  [rate limit] waiting {wait}s...")
                    time.sleep(wait)
                    continue
                if r.status_code == 200:
                    if is_pdf:
                        try:
                            old_stdout = sys.stdout
                            sys.stdout = open(os.devnull, 'w')
                            
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                                tmp_pdf.write(r.content)
                                tmp_pdf_path = tmp_pdf.name
                                
                            tmp_docx_path = tmp_pdf_path + ".docx"
                            
                            try:
                                cv = Converter(tmp_pdf_path)
                                cv.convert(tmp_docx_path, start=0, end=None)
                                cv.close()
                            finally:
                                sys.stdout.close()
                                sys.stdout = old_stdout
                                
                            converted_bytes = Path(tmp_docx_path).read_bytes()
                            cached.write_bytes(converted_bytes)
                            
                            Path(tmp_pdf_path).unlink(missing_ok=True)
                            Path(tmp_docx_path).unlink(missing_ok=True)
                            
                            print(f"  [✓] Converted PDF: {Path(path).name}")
                            success = True
                            time.sleep(1)
                            break
                        except Exception as e:
                            if sys.stdout != old_stdout:
                                sys.stdout.close()
                                sys.stdout = old_stdout
                            print(f"  [!] PDF Error: {e}")
                            break
                    else:
                        if r.content[:4] == b'PK\x03\x04':
                            cached.write_bytes(r.content)
                            print(f"  [✓] Downloaded DOCX: {Path(path).name}")
                            success = True
                            time.sleep(1)
                            break
                        else:
                            print("  [!] Invalid DOCX bytes")
                            break
                else:
                    print(f"  [!] HTTP {r.status_code}")
                    break
            except Exception as e:
                print(f"  [!] Network error: {e}")
                time.sleep(2)
        if not success:
            print(f"  [X] Failed completely: {Path(path).name}")

if __name__ == "__main__":
    run()
