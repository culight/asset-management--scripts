import os.path
import getopt
import sys
from wand.image import Image
from wand.color import Color
from os.path import expanduser

gray = '#CCCCCC'
home_path = expanduser("~") + '/'

input = ''
exten = '.png'

def main(argv):
    global input

    try:
        opts, args = getopt.getopt(argv, "hi:o:", ["ifile=", "ofile="])
    except getopt.GetoptError:
        print '\n' + 'image_neutraliaztion.py -i <drawable folder>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print '\n' + 'image_extraction.py -i <drawable folder> '
            sys.exit()
        elif opt in ("-i", "--ifile"):
            input = arg


def neutralize(ext, dirname, names):
    for name in names:
        base_path = home_path + input
        image_path = base_path + name
        neutral_path = base_path + 'neutral/'

        if not os.path.exists(neutral_path):
            os.makedirs(neutral_path)

        output_path = neutral_path + name
        if name.endswith('.png'):
            with Image(filename=image_path) as img:
                img.blank(width=img.width, height=img.height, background=Color(gray))
                img.save(filename=output_path)
                segments = name.split(".")
                image_name = segments[0]
                image_path = neutral_path + image_name
                if os.path.isfile(image_path + '-0' + '.png'):
                    os.remove(image_path + '-0' + '.png')
                if os.path.isfile(image_path + '-1' + '.png'):
                    os.rename(image_path + '-1' + '.png', image_path + '.png')

            print name
        elif name.endswith('.jpg'):
            with Image(filename=image_path) as img:
                img.blank(width=img.width, height=img.height, background=Color(gray))
                img.save(filename=output_path)
                segments = name.split(".")
                image_name = segments[0]
                image_path = neutral_path + image_name
                if os.path.isfile(image_path + '-0' + '.jpg'):
                    os.remove(image_path + '-0' + '.jpg')
                if os.path.isfile(image_path + '-1' + '.jpg'):
                    os.rename(image_path + '-1' + '.jpg', image_path + '.jpg')
            print name
        elif name.endswith('.gif'):
            with Image(filename=image_path) as img:
                img.blank(width=img.width, height=img.height, background=Color(gray))
                img.save(filename=output_path)
                segments = name.split(".")
                image_name = segments[0]
                image_path = neutral_path + image_name
                if os.path.isfile(image_path + '-0' + '.gif'):
                    os.remove(image_path + '-0' + '.gif')
                if os.path.isfile(image_path + '-1' + '.gif'):
                    os.rename(image_path + '-1' + '.gif', image_path + '.gif')
            print name

if __name__ == "__main__":
    main(sys.argv[1:])

os.path.walk(input, neutralize, exten)