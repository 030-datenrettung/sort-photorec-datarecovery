Sort PhotoRec files from raw data recovery
======================================================

A python script to sort results of a [data recovery](https://www.030-datenrettung.de) with tools like PhotoRec, DiskDrill and others where one only gets RAW results after deep scan of media like hdd, RAID, pendrive or memory-card. The script reads metadata from recovered files and moves the files in a folder-structure by extension/year/month. 

The script uses date-taken for images and date last changed (modern office-formats) / created (older office-formats) and creates a foler-structure like EXT/YYYY/MM.  

## Installation
    Install Phyton
    pip install pikepdf
    pip install pywin32
    pip install pypdf2

## Usage

In config.ini specify file extensions to move and sort as well as source and destination folders. Source & destination folders can also be used on command-line.

Run script:

```commandline
sort-photorec-datarecovery.py SOURCE-FOLDER DESTINATION-FOLDER
```

### Supported fileTypes

Pictures: RGB, GIF, PBM, PGM, PPM, TIFF, RAST, XBM, JPEG, JPG, BMP, RAW, PNG, WEBP, EXR, ARW, CR2, DNG, NEF, ORF

Office: PDF, DOCX, PPTX, XLSX, XLS, DOC, PPT

### License

MIT License

Copyright (c) 2022 030 Datenrettung Berlin GmbH

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
