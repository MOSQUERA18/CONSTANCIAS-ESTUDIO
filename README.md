PASO A PASO PARA INSTALAR DEPENDENCIAS Y PONER COMO .EXE

pip install pandas python-docx


pyinstaller --onefile --windowed --add-data "C:\\CONSTANCIAS\\plantilla_constancias.docx;." --add-data "C:\\CONSTANCIAS\\plantilla_constancias.xlsx;." app.py
