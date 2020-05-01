from wand.image import Image

f = "OpTransactionHistory.pdf"
with(Image(filename=f, resolution=120)) as source: 
    for i, image in enumerate(source.sequence):
        newfilename = f[:-4] + str(i + 1) + '.jpeg'
        Image(image).save(filename=newfilename)