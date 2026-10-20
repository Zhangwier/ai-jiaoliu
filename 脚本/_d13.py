from PIL import Image
im = Image.open(r"D:\MG\program\ai交流\成品\_work29\png_orig\page-0001.png")
print(im.size)
im.resize((im.width//2, im.height//2)).save(r"D:\MG\program\ai交流\成品\_work29\prev_half.png")
