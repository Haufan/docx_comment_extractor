# Projekt Parteirogramme - get_docx_comments.py

Das Programm extrahiert alle Kommentare aus docx-Dateien und deren Seitenzahlen aus dem Original pdf-Dokument.
Die Daten werden in einer csv-Datei ausgegeben (Der Separator ist aber '|').
  
## Requirements

Python 3.10 oder höher.

Die folgenden Packages müssen installiert sein. In Klammern stehen die benutzen Versionen. 
- alive_progress   (3.1.5)
- argparse         (1.4.0)
- lxml             (5.2.2)
- os
- pandas           (2.2.2)
- PyPDF2           (3.0.1)
- re
- zipfile

Es liegt eine requirements.txt bei.
```
pip install -r requirements.txt
```

  
## How to Run

Das Programm kann über die Kommandozeile gestartet werden.

Variante Daten in Programmordner data
```
python get_docx_comments.py
```

Variant Daten in einem anderen Ordner
```
python get_docx_comments.py <path_of_folder>
```

  
  
Dietmar Benndorf  
26.07.2024

