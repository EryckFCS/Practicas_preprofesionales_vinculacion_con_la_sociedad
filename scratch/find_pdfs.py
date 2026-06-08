import os

for root, dirs, files in os.walk("/home/erick-fcs/Descargas/Practicas_preprofesionales"):
    for file in files:
        if file.lower().endswith(".pdf"):
            print(os.path.join(root, file))
