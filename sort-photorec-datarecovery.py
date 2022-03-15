import os
import shutil
import subprocess
import sys
import xml.dom.minidom
import zipfile
from configparser import ConfigParser
from datetime import datetime
from struct import *

import pikepdf
import win32com.client as win32
from PIL import Image
from PIL.ExifTags import TAGS
from PyPDF2 import PdfFileReader


def get_images_meta_data(file):
    try:
        try:
            ifd_offset = 0x10
            f = open(file, "rb")
            buffer = f.read(1024)
            (num_of_entries,) = unpack_from('H', buffer, ifd_offset)

            datetime_offset = -1
            for entry_num in range(0, num_of_entries - 1):
                (tag_id, tag_type, num_of_value, value) = unpack_from('HHLL', buffer, ifd_offset + 2 + entry_num * 12)
                if tag_id == 0x0132:
                    datetime_offset = value

            datetime_offset = "".join(
                [a.decode('ascii') for a in unpack_from(20 * 's', buffer, datetime_offset) if a not in ['', '\x00', None]])
            datetime1 = datetime_offset[:-1]
            date = datetime.strptime(datetime1.strip(' '), '%Y:%m:%d %H:%M:%S')
            return date
        except:
            pass



        try:
            image = Image.open(file)
        except Exception as e:
            exeProcess = "hachoir-metadata"
            process = subprocess.Popen([exeProcess, file],
                                       stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                       universal_newlines=True)

            for tag in process.stdout:
                if 'creation date' in str(tag).lower():
                    date=datetime.strptime(tag.strip().strip('\n'),'- Creation date: %Y-%m-%d %H:%M:%S')
                    return date
            return False

        exifdata = image.getexif()
        for tag_id in exifdata:
            tag = TAGS.get(tag_id, tag_id)
            if tag=='DateTime':
                date = exifdata.get(tag_id)
                if isinstance(date, bytes):
                    date = date.decode()
                try:
                    date = datetime.strptime(date, '%Y:%m:%d %H:%M:%S')
                except :
                    date= datetime.strptime(date[:24], '%a %b %d %H:%M:%S %Y')  #Update format
                return date
        for tag_id in exifdata:
            tag = TAGS.get(tag_id, tag_id)
            if tag=='DateTimeOriginal':
                date = exifdata.get(tag_id)
                if isinstance(date, bytes):
                    date = date.decode()
                date = datetime.strptime(date, '%Y:%m:%d %H:%M:%S')
                return date




        return False
    except :
        return False

def get_pdf_metadata(file):
    print(file)
    with open(file, 'rb') as pdf:
        pdfFile = PdfFileReader(pdf)

        if pdfFile.isEncrypted:
            try:
                pdfFile.decrypt('')
                print('File Decrypted (PyPDF2)')
            except:
                try:
                    command = ("cp " + file +
                               " temp.pdf; qpdf --password='' --decrypt temp.pdf " + file
                               + "; rm temp.pdf")
                    os.system(command)
                    print('File Decrypted (qpdf)')
                    with open(file) as fp:
                        pdfFile = PdfFileReader(fp)
                except:
                    try:
                        return datetime.strptime(pikepdf.open(file).docinfo['/ModDate'].split(':')[1].replace('+', '-').split('-')[0],"%Y%m%d%H%M%S")
                    except:
                        return False


        else:
            print('File Not Encrypted')

        pdf_info = pdfFile.getDocumentInfo()
        if '/ModDate' in pdf_info:
            return datetime.strptime(pdf_info['/ModDate'].split(':')[1].replace('+','-').replace('Z','-').split('-')[0], "%Y%m%d%H%M%S")
            del pdfFile
        else:
            return False

def get_office_files_meta_data(file):


    if file.endswith('xls'):#,'ppt','doc')):
        xlApp = win32.DispatchEx('Excel.Application')
        wb = xlApp.Workbooks.Open(file)
        xlApp.Visible = False
        xlApp.DisplayAlerts = False
        wb.DoNotPromptForConvert = True
        wb.CheckCompatibility = False
        date=wb.BuiltinDocumentProperties[10].value
        wb.Close()
        xlApp.Quit()
        return date

    elif file.endswith('doc'):
        wordApp=win32.DispatchEx("Word.Application")
        document = wordApp.Documents.Open(file)
        wordApp.Visible = False
        wordApp.DisplayAlerts = False
        date = document.BuiltinDocumentProperties[10].value
        document.Close()
        wordApp.Quit()
        return date
    
    elif file.endswith('ppt'):
        pptApp=win32.DispatchEx("Powerpoint.Application")
        ppt =  pptApp.Presentations.open(file, True, False, False)
        date = ppt.BuiltinDocumentProperties[10].value
        ppt.Close()
        pptApp.Quit()
        return date


    try:
        myFile = zipfile.ZipFile(file, 'r')
    except :
        return False

    #for old doc files , it will return false
    try:
        doc = xml.dom.minidom.parseString(myFile.read('docProps/core.xml'))
        xml.dom.minidom.parseString(myFile.read('docProps/core.xml')).toprettyxml()
    except:
        return False

    try:
        date= doc.getElementsByTagName('dcterms:modified')[0].childNodes[0].data
        try:
            date=datetime.strptime(date, '%Y:%m:%d %H:%M:%S')
            return date
        except:
            date=datetime.strptime(date, '%Y-%m-%dT%H:%M:%SZ')
            return date
        return False
    except:
        return False


def main():
    config=ConfigParser()
    config.read('config.ini')

    try:
        input_path=sys.argv[1]
        output_path=sys.argv[2]
    except:
        try:
            input_path = config.get('PATHS', 'input_directory')
            output_path = config.get('PATHS', 'output_directory')
        except:
            print('Error in Path arguments')
            return

    if not os.path.exists(input_path):
        raise FileNotFoundError('ERROR : Invalid Input Path')

    if not os.path.exists(output_path):
        os.mkdir(output_path)

    file_list = list()
    for (dirpath, dirnames, filenames) in os.walk(input_path):
        file_list += [os.path.join(dirpath, file) for file in filenames]

    for file in file_list:
        ext=os.path.splitext(file)[1].lstrip('.').upper()

        if config.get('EXTENSIONS','ALL')=='False':
            if ext not in [e.strip().upper() for e in config.get('EXTENSIONS','required_extensions').split(',')]:
                continue
        print(ext)
        target_path = os.path.join(output_path, str(ext))

        OfficeFiles=config.get('OfficeFiles','extension_list')
        ImageFiles=config.get('ImageFiles','extension_list')

        if len(ext)>=3 and ext in OfficeFiles:
            if ext in ['PDF']:
                date=get_pdf_metadata(file)
            else:
                date=get_office_files_meta_data(file)


            if date is not False:
                month=str(date.month) if date.month>9 else '0'+str(date.month)
                target_path = os.path.join(target_path, str(date.year),month)


        elif ext in ImageFiles:
            date=get_images_meta_data(file)
            if date is not False:
                month = str(date.month) if date.month > 9 else '0' + str(date.month)
                target_path = os.path.join(target_path, str(date.year), month)

        os.makedirs(target_path,exist_ok=True)
        print('moving',file ,'       to       ',target_path)
        shutil.copy(file,target_path)
        try:
            os.remove(file)
        except:
            pass
if __name__=='__main__':
    main()