from pptx import Presentation
for f in ["_v35/build.pptx","_v36/build.pptx"]:
    try:
        p=Presentation(f); print(f, "slides:", len(p.slides))
    except Exception as e: print(f,"ERR",e)
