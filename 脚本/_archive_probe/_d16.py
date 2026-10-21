import glob, os, shutil
for p in [r"C:\Program Files\LibreOffice\program\soffice.exe",
          r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
    print(p, os.path.exists(p))
