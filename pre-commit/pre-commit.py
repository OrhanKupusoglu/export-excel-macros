import os
import shutil
from oletools.olevba3 import VBA_Parser
import xml.etree.ElementTree as ET
import sys
import zipfile

EXACT_MATCH = False
VBA_EXT = 'bas'
KEEP_NAME = False  # set this to True if you would like to keep "Attribute VB_Name"

def get_work_sheets(workbook_path):
    zf = zipfile.ZipFile(workbook_path)
    f = zf.open(r'xl/workbook.xml')
    l = f.readline()
    l = f.readline()
    root = ET.fromstring(l)
    sheets=[]
    for c in root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets/*'):
        sheets.append(c.attrib['name'])
    return sheets

# worksheet's name doesn't need to match exactly
def check_worksheets(worksheet_name, workbook_path):
    ok = True
    work_sheets = get_work_sheets(workbook_path)
    if EXACT_MATCH:
        if worksheet_name in work_sheets:
            ok = False
    else:
        if any(worksheet_name in w for w in work_sheets):
            ok = False
    return ok

def parse(vba_dir, workbook_path):
    vba_parser = VBA_Parser(workbook_path)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

    for _, _, filename, content in vba_modules:
        lines = []
        if '\r\n' in content:
            lines = content.split('\r\n')
        else:
            lines = content.split('\n')
        if lines:
            content = []
            for line in lines:
                if line.startswith('Attribute') and 'VB_' in line:
                    if 'VB_Name' in line and KEEP_NAME:
                        content.append(line)
                else:
                    content.append(line)
            if content and content[-1] == '':
                content.pop(len(content)-1)
                non_empty_lines_of_code = len([c for c in content if c])
                if non_empty_lines_of_code > 0:
                    if not os.path.exists(os.path.join(vba_dir)):
                        os.makedirs(vba_dir)
                    with open(os.path.join(vba_dir, f"{filename}.{VBA_EXT}"), 'w', encoding='utf-8') as f:
                        f.write('\n'.join(content))

def check_staged_files(args):
    vba_dir = args[0]
    excel_exts = tuple(args[1].split(':'))
    worksheet_name = args[2]
    staged_files = args[3:]
    deleted = False

    # unwanted worksheets?
    if worksheet_name.strip() != '':
        for f in staged_files:
            if f.endswith(excel_exts):
                if not check_worksheets(worksheet_name, f):
                    sys.exit(1)

    # export vba
    for f in staged_files:
        if f.endswith(excel_exts):
            if not deleted:
                if os.path.exists(vba_dir):
                    deleted = True
                    shutil.rmtree(vba_dir)

            ff = os.path.basename(f).split('/')[-1]
            if not ff.startswith('~') and ff.endswith(excel_exts):
                parse(vba_dir, f)

if __name__ == '__main__':
    check_staged_files(sys.argv[1:])
