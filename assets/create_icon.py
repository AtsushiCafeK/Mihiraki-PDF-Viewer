from PIL import Image

sizes = [(256,256), (128,128), (64,64), (32,32)]
img = Image.open("image.png")
img.save("mihiraki.ico", sizes=sizes)
