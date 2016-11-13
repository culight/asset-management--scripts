import xlsxwriter as XW
import os.path
import sys
import getopt
import time
from PIL import Image
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
        print '\n' + 'image_extraction.py -i <asset directory> -e <excel document name>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print '\n' + 'image_extraction.py -i <asset directory> -e <excel document name>'
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
    images_path = project_path + 'asset-management/images/'

    # Create images directory
    if not os.path.exists(images_path):
        os.makedirs(images_path)

    print '\n' + 'Input:', directory_path

    # Create project level directory
    if len(excel_file) < 1:
        output_name = get_output_name(input)
        print 'Output:', images_path + output_name + '/' + '\n'
    else:
        output_name = excel_file

        print 'Output:', images_path + output_name + '/' + '\n'

    working_path = images_path + output_name + '/'

    if not os.path.exists(working_path):
        os.makedirs(working_path)

    # Create excel file in local project path
    excel_file_name = output_name + '.xlsx'
    excel_file_path = working_path + excel_file_name

    # Create source folder in project path
    if not os.path.exists(working_path + 'source'):
        os.makedirs(working_path + 'source')

    # Create replaced folder in project path
    if not os.path.exists(working_path + 'replaced'):
        os.makedirs(working_path + 'replaced')


def read_image_dirs():
    print 'Reading image directories...' + '\n'
    workbook = XW.Workbook(excel_file_path)

    for directory, dirnames, filenames in os.walk(directory_path, topdown=True):
        # if this directory does not have low-level directories
        if not dirnames:
            print 'From ' + get_sheet_name(directory) + ':'
            if not filenames:
                continue
            else:
                read_images(workbook, directory, filenames)
                print '\n'
        else:
            if filenames:
                read_images(workbook, directory, filenames)

    # Clear out empty sheets
    workbook.close()

def read_images(wb, directory, names):
    sheet_name = get_sheet_name(directory)
    if len(sheet_name) > 31:
        sheet_name = sheet_name[:31]

    try:
        worksheet = wb.add_worksheet(sheet_name)
    except:
        return

    # Widen the columns being used
    worksheet.set_column('A:A', 50)
    worksheet.set_column('B:B', 100)
    worksheet.set_column('C:C', 30)
    worksheet.set_column('D:D', 50)
    worksheet.set_column('E:E', 100)
    worksheet.set_column('F:F', 100)

    # Set the font formats
    label_format = wb.add_format({'font_size': 18})
    label_format.set_align('top')
    path_format = wb.add_format({'font_size': 12})
    path_format.set_align('top')

    # Add headers
    worksheet.write('A1', 'Image Name', label_format)
    worksheet.write('B1', 'Image Preview', label_format)
    worksheet.write('C1', 'Image Dimensions', label_format)
    worksheet.write('D1', 'Last Updated', label_format)
    worksheet.write('E1', 'Current File Path', label_format)
    worksheet.write('F1', 'Source File', label_format)

    row = 2
    cell_written = False

    for name in names:
        # png
        if name.lower().endswith('.png'):
            # Write image name
            worksheet.write('A' + str(row), name, label_format)
            # Plop image preview
            worksheet.insert_image('B' + str(row), directory + '/' + name, {'x_scale': 0.5, 'y_scale': 0.5})
            # Retrieve dimensions of image
            dimen = image_size(directory + '/' + name)
            dimen_string = '(' + str(dimen[0]) + ',' + str(dimen[1]) + ')'
            worksheet.write('C' + str(row), dimen_string, label_format)
            # Write "last updated" statistic
            timestamp = image_last_updated(directory + '/' + name)
            worksheet.write('D' + str(row), timestamp, label_format)
            # Write the image's current file path
            worksheet.write('E' + str(row), directory + '/' + name, path_format)
            cell_written = True
            print(os.path.join(directory, name))
        # jpg
        elif name.lower().endswith('.jpg'):
            # Write image name
            worksheet.write('A' + str(row), name, label_format)
            # Plop image preview
            worksheet.insert_image('B' + str(row), directory + '/' + name, {'x_scale': 0.5, 'y_scale': 0.5})
            # Retrieve dimensions of image
            dimen = image_size(directory + '/' + name)
            dimen_string = '(' + str(dimen[0]) + ',' + str(dimen[1]) + ')'
            worksheet.write('C' + str(row), dimen_string, label_format)
            # Write "last updated" statistic
            timestamp = image_last_updated(directory + '/' + name)
            worksheet.write('D' + str(row), timestamp, label_format)
            # Write the image's file path
            worksheet.write('E' + str(row), directory + '/' + name, path_format)
            cell_written = True
            print(os.path.join(directory, name))
        # gif
        elif name.lower().endswith('.gif'):
            # Write image name
            worksheet.write('A' + str(row), name, label_format)
            # Plop image preview
            worksheet.insert_image('B' + str(row), directory + '/' + name, {'x_scale': 0.5, 'y_scale': 0.5})
            # Retrieve dimensions of image
            dimen = image_size(directory + '/' + name)
            dimen_string = '(' + str(dimen[0]) + ',' + str(dimen[1]) + ')'
            worksheet.write('C' + str(row), dimen_string, label_format)
            # Write "last updated" statistic
            timestamp = image_last_updated(directory + '/' + name)
            worksheet.write('D' + str(row), timestamp, label_format)
            # Write the image's file path
            worksheet.write('E' + str(row), directory + '/' + name, path_format)
            cell_written = True
            print(os.path.join(directory, name))
        # bmp
        elif name.lower().endswith('.bmp'):
            # Write image name
            worksheet.write('A' + str(row), name, label_format)
            # Plop image preview
            worksheet.insert_image('B' + str(row), directory + '/' + name, {'x_scale': 0.5, 'y_scale': 0.5})
            # Retrieve dimensions of image
            dimen = image_size(directory + '/' + name)
            dimen_string = '(' + str(dimen[0]) + ',' + str(dimen[1]) + ')'
            worksheet.write('C' + str(row), dimen_string, label_format)
            # Write "last updated" statistic
            timestamp = image_last_updated(directory + '/' + name)
            worksheet.write('D' + str(row), timestamp, label_format)
            # Write the image's file path
            worksheet.write('E' + str(row), directory + '/' + name, path_format)
            cell_written = True
            print(os.path.join(directory, name))
        # ico
        elif name.lower().endswith('.ico'):
            # Write image name
            worksheet.write('A' + str(row), name, label_format)
            # Plop image preview
            worksheet.insert_image('B' + str(row), directory + '/' + name, {'x_scale': 0.5, 'y_scale': 0.5})
            # Retrieve dimensions of image
            dimen = image_size(directory + '/' + name)
            dimen_string = '(' + str(dimen[0]) + ',' + str(dimen[1]) + ')'
            worksheet.write('C' + str(row), dimen_string, label_format)
            # Write "last updated" statistic
            timestamp = image_last_updated(directory + '/' + name)
            worksheet.write('D' + str(row), timestamp, label_format)
            # Write the image's file path
            worksheet.write('E' + str(row), directory + '/' + name, path_format)
            cell_written = True
            print(os.path.join(directory, name))
        row += 1
    if cell_written is False:
        worksheet.hide()


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


def image_size(image_file):
    im = Image.open(image_file)
    return im.size


def image_last_updated(image_file):
    mod_time = time.ctime(os.path.getmtime(image_file))
    return mod_time


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
    read_image_dirs()
