import os
from docx2pdf import convert

os.makedirs("./pdf", exist_ok=True)
convert("./ProcessedWordFile", "./pdf")
