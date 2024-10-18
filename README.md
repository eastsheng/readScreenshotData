This code can take a screenshot to quickly read the data in the paper table.

### Dependencies

```
pip install pyautogui pillow pytesseract pandas openpyxl
```

- Install  Tesseract OCR（Windows）。
  - [Tesseract download](https://github.com/UB-Mannheim/tesseract/wiki)

- After installation, you need to specify the path to Tesseract. Configure the path in your code, for example on Windows:

```
pytesseract.pytesseract.tesseract_cmd = r'D:\Program Files\Tesseract-OCR\tesseract.exe'
```

