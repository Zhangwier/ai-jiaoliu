import zipfile, sys, shutil

src = sys.argv[1]
dst = sys.argv[2]
zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == "ppt/presentation.xml":
        d = d.decode("utf-8").replace('type="screen4x3"', 'type="screen16x9"').encode("utf-8")
    zout.writestr(it, d)
zout.close()
zin.close()
print("ok")
