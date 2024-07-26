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


## Further Remarks

(!) Es ist darauf zu achten, dass die Dateien folgende Form haben:<br><br>
*Annotationsdateien*: z.B. FDP-Bbg-2024.docx 
```
<dreistelliges_Kürzel_großgeschrieben><nicht_relevant>.docx
```

*Parteiprogrammsdateien*: z.B. FDP-Bbg-2024.pdf 
```
<dreistelliges_Kürzel_großgeschrieben><nicht_relevant>.pdf
```

(!) Die Annotationen innerhalb der docx müssen immer mit einem Kürzel wie z.B. P5 (Kombination aus P und einer Zahl).
Dahinter können frei Kommentare stehen.

Es befindet sich ein Beispiel der Dateien im data-Ordner und eine Beispielausgabe unter BEISPIEL_data.csv.
<br><br><br><br>  
  
Dietmar Benndorf  
26.07.2024

