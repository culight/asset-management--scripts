import xml.etree.ElementTree as ET
import xlsxwriter as XW
import getopt
import sys
import os.path
from os.path import expanduser

input = ''
excel_file = ''
multi_dirs = True

home_path = expanduser("~") + '/'
project_path = ''
directory_path = ''
strings_path = ''

excel_file_name = ''
excel_file_path = ''


def main(argv):
    global input
    global excel_file
    global multi_dirs

    try:
        opts, args = getopt.getopt(argv, "hi:e:s", ["ifile=", "efile=", "dir="])
    except getopt.GetoptError:
        print '\n' + 'string_extraction.py -i <asset directory> -e <excel document name>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print '\n' + 'string_extraction.py -i <asset directory> -e <excel document name>'
            sys.exit()
        elif opt in ("-i", "--ifile"):
            input = arg
        elif opt in ("-e", "--efile"):
            excel_file = arg
        elif opt in ("-s", "--dir"):
            multi_dirs = False


def initial_setup():
    global project_path
    global directory_path
    global strings_path
    global excel_file_name
    global excel_file_path

    project_path = get_project_path()
    directory_path = get_directory_path()
    strings_path = project_path + 'asset-management/strings/'

    # Create stringz directory
    if not os.path.exists(strings_path):
        os.makedirs(strings_path)

    print '\n' + 'Input:', directory_path

    # Create project level directory
    if len(excel_file) < 1:
        output_name = get_output_name(input)
        print 'Output:', strings_path + output_name + '/' + '\n'
    else:
        output_name = excel_file

        print 'Output:', strings_path + output_name + '/' + '\n'

    working_path = strings_path + output_name + '/'

    if not os.path.exists(working_path):
        os.makedirs(working_path)

    # Create excel file in local project path
    excel_file_name = output_name + '.xlsx'
    excel_file_path = working_path + excel_file_name


def read_strings_dir():
    print 'Reading value directories...' + '\n'
    workbook = XW.Workbook(excel_file_path)

    for directory, dirnames, filenames in os.walk(directory_path, topdown=True):
        # if this directory does not have low-level directories
        if not dirnames:
            print 'From ' + get_sheet_name(directory) + ':'
            if not filenames:
                continue
            else:
                read_strings(workbook, directory, filenames)
                print '\n'
        else:
            if filenames:
                read_strings(workbook, directory, filenames)

    # Clear out empty sheets
    workbook.close()


def read_strings(wb, directory, names):

    # Get acceptable sheet name
    sheet_name = get_sheet_name(directory)
    if len(sheet_name) > 31:
        sheet_name = sheet_name[:31]

    try:
        ws = wb.add_worksheet(sheet_name)
    except:
        return

    # Widen the columns being used
    ws.set_column('A:A', 20)
    ws.set_column('B:B', 50)
    ws.set_column('C:C', 50)
    ws.set_column('D:D', 50)

    # Set the font formats
    label_format = wb.add_format({'font_size': 18})
    label_format.set_align('top')
    path_format = wb.add_format({'font_size': 18})
    path_format.set_align('top')
    path_format.set_text_wrap()

    # Add headers
    ws.write('A1', 'Tag', label_format)
    ws.write('B1', 'Name', label_format)
    ws.write('C1', 'Parent', label_format)
    ws.write('D1', 'Value', label_format)

    row = 2
    entry = {}
    entries = []
    cell_written = False

    for xmlfile in names:
        if not str(xmlfile).endswith('.xml'):
            continue

        element_tree = ET.parse(directory + '/' + xmlfile)
        root = element_tree.getroot()

        for resource in root:
            entry['tag'] = resource.tag
            ws.write('A' + str(row), entry['tag'], path_format)
            entry['name'] = resource.get('name')
            ws.write('B' + str(row), entry['name'], path_format)
            entry['parent'] = 'root'
            ws.write('C' + str(row), entry['parent'], path_format)
            entry['value'] = resource.text
            ws.write('D' + str(row), entry['value'], path_format)
            entries.append(entry)
            row += 1
            entry = {}
            for item in resource:
                entry['tag'] = item.tag
                ws.write('A' + str(row), entry['tag'], path_format)
                entry['name'] = item.get('name')
                ws.write('B' + str(row), entry['name'], path_format)
                entry['parent'] = resource.get('name')
                ws.write('C' + str(row), entry['parent'], path_format)
                entry['value'] = item.text
                ws.write('D' + str(row), entry['value'], path_format)
                entries.append(entry)
                row += 1
                entry = {}
                cell_written = True

    if cell_written is False:
        ws.hide()


def home_path_present():
    if home_path not in input:
        return False
    return True


def get_project_path():
    if home_path_present() is not True:
        abs_path = home_path + input
    else:
        abs_path = input
    segments = abs_path.split("/")
    del segments[-2:]
    str = '/'
    project_path = str.join(segments) + '/'
    return project_path


def get_directory_path():
    dir_path = input
    if home_path_present() is not True:
        dir_path = home_path + input
    return dir_path


def get_output_name(input_path):
    segments = input_path.split("/")
    if (len(segments) < 3):
        output_name = segments[0]
    else:
        output_name = segments[-2]
    return output_name


def get_sheet_name(input_path):
    segments = input_path.split("/")
    if (len(segments) < 3):
        sheet_name = segments[0]
    else:
        sheet_name = segments[-1]
    return sheet_name

if __name__ == "__main__":
    main(sys.argv[1:])

initial_setup()

if multi_dirs is True:
    read_strings_dir()

