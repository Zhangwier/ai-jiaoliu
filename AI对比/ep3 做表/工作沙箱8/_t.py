from paddleocr import TextRecognition
print([x for x in dir(TextRecognition) if not x.startswith("_")])
