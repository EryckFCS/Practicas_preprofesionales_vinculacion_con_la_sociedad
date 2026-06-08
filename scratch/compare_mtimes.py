import os
import time

paths = [
    "/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/25.  INFORME SG DERECHO 2023.docx",
    "/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/pregrado/25.  INFORME SG DERECHO 2023.docx"
]

for p in paths:
    if os.path.exists(p):
        mtime = os.path.getmtime(p)
        print(f"{p} -> Last Modified: {time.ctime(mtime)}, Size: {os.path.getsize(p)}")
    else:
        print(f"{p} -> Not found")
