# risarchiver
**RIS Archiver: Bulk Process Research from RIS files with associated PDFs or TXT files**

Designed as a way to download a set of research sources quickly from a database to process at home. Useful for scholars who may have to travel to access databases and need to catalog items quickly for later reference, as long as full text (PDF or TXT) is available. 

_Tip:_
Download the .RIS file from your archive first, followed by the full text in PDF or TXT or both. The program will associate them by timestamp, so the order you download is important. The file names are then recorded in the .xlsx spreadsheet for you to locate or access later.

By default, the .xlsx sheet is created in the same folder as your downloaded items.


**Command Line Usage:**

usage: ris-archiver [-h] [--folder FOLDER | --ris RIS [RIS ...]] [--xlsx XLSX]
                    [--outputfolder OUTPUTFOLDER] [--sheet SHEET] [--auto-pdf]
                    [--no-auto-pdf] [--flag-duplicates] [--no-flag-duplicates]
                    [--skip-processed] [--no-skip-processed]
                    [--columns COLUMNS]

Archive RIS citations into an Excel workbook (append-only), with PDF/TXT
association, duplicate flagging, and skip-processed.

options:
  -h, --help            show this help message and exit
  --folder FOLDER       Folder containing .ris and .pdf/.txt files.
  --ris RIS [RIS ...]   One or more .ris files to process.
  --xlsx XLSX           Output .xlsx path OR filename (default: archive.xlsx
                        in outputfolder/input folder).
  --outputfolder OUTPUTFOLDER
                        Output folder for .xlsx when --xlsx is a filename or
                        omitted.
  --sheet SHEET         Sheet name (default: Archive).
  --auto-pdf            Auto-associate PDFs by timestamp (soonest AFTER RIS
                        time).
  --no-auto-pdf         Disable auto PDF association.
  --flag-duplicates     Flag duplicates (append anyway; mark Duplicate? and
                        Duplicate Key).
  --no-flag-duplicates  Disable duplicate flagging.
  --skip-processed      Skip RIS files already logged as processed in workbook
                        metadata sheet.
  --no-skip-processed   Do not skip processed RIS files.
  --columns COLUMNS     Comma-separated column list.
