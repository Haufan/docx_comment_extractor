# docx_comment_extractor

Programm extrahiert aus docx-Dateien alle Kommentare und die dazugehörigen Textstellen. Zusätzlich extrahiert es aus dem Original PDF-Dokument (docx-Datei aus diesem erstellt) die Seitenzahl (das funktioniert leider nicht 100%).
Die Daten werden in einer csv-Datei ausgegeben, aber haben "|" als Separator (siehe BEISPIEL_data.csv).
Es liegen drei Dateien im data-Ordner, so dass das Programm getestest werden kann.

Um das Programm zu nutzen, muss der data-Ordner geleert werden und mit den zu durchsuchenden Datei gefüllt werden.
   - (!) Alle Dateien müssen mit einem dreistelligen und großgeschriebenen Parteikürzel beginnen (z.B. FDP-Bbg-2024.docx, FDP-Bbg-2024.pdf)
   - (!) Die Annotationsdateien müssen im docx- und das Parteiprogramm im PDF-Format sein.
   - (!) Es ist nicht relevant, was zwischen PArteikürzel und Dateiformat steht.
