from PIL import Image
im=Image.open(r"r1\page-0003.png")
im.crop((90,85,760,150)).resize((1340,130),Image.LANCZOS).save("crop_title.png")
im.crop((700,280,760,660)).resize((240,1520),Image.LANCZOS).save("crop_vr.png")
