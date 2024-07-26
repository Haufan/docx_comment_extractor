# Projekt Parteiprogramme - docx-Kommentar-Extrahierer
# Dietmar Benndorf
# 2024


from alive_progress import alive_bar
import argparse
from lxml import etree
import os
import pandas
import PyPDF2
import re
import zipfile


def get_file_list(*args):
    """
    Extracts a list of files with docx formats and
    creates a dictionary of parties and their program-file
    """

    annot_file_list = []
    party_file_dict = {}
    print(args)
    _path = os.listdir(args[0])
    print(_path)

    # Ordnet Dateien in Liste mit docx und Dict nach Partei: pdf
    for file in _path:
        if file.endswith(".docx"):
            annot_file_list.append(file)
        elif file.endswith(".pdf"):
            party_file_dict[file[:3]] = file

    # Checkt, ob Annotationsdateien existieren
    if annot_file_list == []:
        print("There are no annotation files. "
              "Check the folder or if they are named correctly (e.g. SPD***.docx).")

    # Checkt, ob jede Annotation einem Parteiprogramm zugeordnet werden kann
    for file in annot_file_list:
        if file[:3] not in party_file_dict.keys():
            print(f"The following party program is missing: {file[:3]}")
            quit()
        else:
            pass

    return [annot_file_list, party_file_dict]


def get_all_data(files):
    """Loops through all files and extract comments+information and saves it as a csv."""

    annot_file_list = files[0]
    party_file_dict = files[1]
    all_data_df = pandas.DataFrame(columns=['File',
                                            'Annotation',
                                            'Comment',
                                            'Text',
                                            'Page'])

    with alive_bar(len(annot_file_list), force_tty=True) as bar:

        for file in annot_file_list:
            _temp_annot_file_path = "data/" + file
            _temp_party_file_path = "data/" + party_file_dict[file[:3]]
            file_data_df = get_docx_comments(_temp_annot_file_path, _temp_party_file_path)

            all_data_df = pandas.concat([all_data_df, file_data_df], ignore_index=True)

        bar()

    all_data_df.to_csv('data.csv', sep='|', mode='a', encoding='utf-8')
    print('Data can be found in data.csv')


def get_docx_comments(filename, party_file):
    """
    Extracts all comments and commented on text passages from a docx.
    In addition, it finds the page number of that passage in the original PDF document
    (through calling get_page_pdf()).
    """

    comments_page = {}

    # Code von https://stackoverflow.com/questions/47390928/extract-docx-comments
    ooXMLns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    comments_dict = {}
    comments_of_dict = {}
    docx_zip = zipfile.ZipFile(filename)
    comments_xml = docx_zip.read('word/comments.xml')
    comments_of_xml = docx_zip.read('word/document.xml')
    et_comments = etree.XML(comments_xml)
    et_comments_of = etree.XML(comments_of_xml)
    comments = et_comments.xpath('//w:comment', namespaces=ooXMLns)
    comments_of = et_comments_of.xpath('//w:commentRangeStart', namespaces=ooXMLns)

    # Kommentare
    for c in comments:
        comment = c.xpath('string(.)', namespaces=ooXMLns)
        comment_id = c.xpath('@w:id', namespaces=ooXMLns)[0]

        _temp_annot = re.findall('P[\d]*', comment)
        _temp_comm = re.sub('P[\d]*', '', comment)

        comments_dict[comment_id] = [_temp_annot[0], _temp_comm]

    # markierter Text
    for c in comments_of:
        comments_of_id = c.xpath('@w:id', namespaces=ooXMLns)[0]
        parts = et_comments_of.xpath(
            "//w:r[preceding-sibling::w:commentRangeStart[@w:id=" +
            comments_of_id + "] and following-sibling::w:commentRangeEnd[@w:id="
            + comments_of_id + "]]", namespaces=ooXMLns)
        comment_of = ''

        if parts != []:
            for part in parts:
                comment_of += part.xpath('string(.)', namespaces=ooXMLns)
                comments_of_dict[comments_of_id] = comment_of
        # weil nicht alle Kommentare erfasst wurden
        else:
            _temp_text = comments_of_xml.decode("utf-8")
            # sucht volle Kommentarzeilen
            _temp_comment = re.findall(
                rf'<w:commentRangeStart w:id="{comments_of_id}".*?{comments_of_id}"/>',
                _temp_text)
            # löscht Metainformation
            comment_of = re.sub(r'<.*?>', '', _temp_comment[0])
            comments_of_dict[comments_of_id] = comment_of

        # Seite
        comments_page[comments_of_id] = get_page_pdf(party_file, comment_of)

    combined_data = []
    for x in comments_dict:
        if x in comments_of_dict.keys():
            _temp_comment_of = comments_of_dict[x]
            _temp_comment_page = comments_page[x]
        else:
            _temp_comment_of = 'FEHLER - Eintrag verloren'
            _temp_comment_page = 'FEHLER - Eintrag verloren'

        combined_data.append([filename,
                              comments_dict[x][0],
                              comments_dict[x][1],
                              _temp_comment_of,
                              _temp_comment_page])
    file_data_df = pandas.DataFrame(combined_data,
                                    columns=['File',
                                             'Annotation',
                                             'Comment',
                                             'Text',
                                             'Page'])

    return file_data_df


def get_page_pdf(pdf_path, sentence):
    """Extracts a page number of a text passage in a pdf document."""

    ### Es werden leider nicht alle Kommentare gefunden. ###

    # für Erkennung notwendig - entfernt Leerzeichen aus Eingabe
    sentence = re.sub(r' ', '', sentence)

    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)

        for page_number in range(len(reader.pages)):
            page = reader.pages[page_number]
            page_text = page.extract_text()

            # für Erkennung notwendig - Entfernen aller versteckten Sonderfunktionen
            mapping = dict.fromkeys(range(32))
            page_text = page_text.translate(mapping)

            innit = re.search(sentence, page_text)
            if innit:
                return page_number + 1

    return 'FEHLER - kein Eintrag gefunden'


if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog='Projekt Parteirogramme - get_docx_comments.py',
                                     description='Das Programm extrahiert alle Kommentare aus docx-Dateien und '
                                                 'deren Seitenzahlen aus dem Original pdf-Dokument. '
                                                 'Die Daten werden in einer csv-Datei ausgegeben '
                                                 '(Der Separator ist aber \'|\').\n')
    parser.add_argument('path', nargs='?', type=str, default='data',
                        help='Ordnerpfad mit Annotationsdateien (.docx) und Parteiprogrammen (.pdf)')
    args = parser.parse_args()

    get_all_data(get_file_list(args.path))
