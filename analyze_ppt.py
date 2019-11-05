"""
See http://pbpython.com/creating-powerpoint.html for details on this script
Requires https://python-pptx.readthedocs.org/en/latest/index.html

Program takes a PowerPoint input file and generates a marked up version that
shows the various layouts and placeholders in the template.
"""

from __future__ import print_function
from pptx import Presentation
import collections
import json
import argparse
import imageio


def parse_args():
    """ Setup the input and output arguments for the script
    Return the parsed input and output files
    """
    parser = argparse.ArgumentParser(description='Analyze powerpoint file structure')
    parser.add_argument('infile',
                        type=argparse.FileType('r'),
                        help='Powerpoint file to be analyzed')
    parser.add_argument('outfile',
                        type=argparse.FileType('w'),
                        help='Output powerpoint')
    return parser.parse_args()


def analyze_ppt(input, output):
    """ Take the input file and analyze the structure.
    The output file contains marked up information to make it easier
    for generating future powerpoint templates.
    """
    # JSON String
    data = """
    {
    	"Common Python data types": {
            "There are many data types in python but the most common ones are": {
                "item1": "foo",
                "item2": "bar"
            }
        },
        "Strings": {
            "In python a string is most easily identified by the use of double quotes": {
                "item1": "bar",
                "item2": "foo"
            }
        }
    }
    """

    titles = json.loads(data, object_pairs_hook=collections.OrderedDict)

    prs = Presentation(input)

    picture_left  = int(prs.slide_width * 0.15)
    picture_top   = int(prs.slide_height * 0.1)
    picture_width = int(prs.slide_width * 0.7)

    for index, title_item in enumerate(titles):
        slide = prs.slides.add_slide(prs.slide_layouts[0])

        img = imageio.imread('meme.jpg')
        picture_height = int(picture_width * img.shape[0] / img.shape[1])
        picture = slide.shapes.add_picture('meme.jpg', picture_left, picture_top, picture_width, picture_height)

        title = slide.shapes.title
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame
        #p = tf.add_paragraph()

        title.text = '{}'.format(title_item)
        tf.text = '{}'.format(titles.get(title_item).keys()[0])
        #p.text = '{}'.format(titles.get(title_item))
        #p.level = 1

    prs.save(output)


if __name__ == "__main__":
    args = parse_args()
    analyze_ppt(args.infile.name, args.outfile.name)
