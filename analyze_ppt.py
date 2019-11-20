from __future__ import print_function
from pptx import Presentation
import collections
import json
import argparse
import imageio
import os


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
                "item2": "foo",
                "image": {
                    "meme.jpg": "top_right"
                }
            }
        },
        "Strings2": {
            "In python a string is most easily identified by the use of double quotes": {
                "item1": "bar",
                "item2": "foo",
                "image": {
                    "meme.jpg": "top_right"
                }
            }
        },
        "Strings3": {
            "In python a string is most easily identified by the use of double quotes": {
                "item1": "bar",
                "item2": "foo",
                "image": {
                    "meme.jpg": "top_right"
                }
            }
        },
        "Strings4": {
            "In python a string is most easily identified by the use of double quotes": {
                "item1": "bar",
                "item2": "foo",
                "image": {
                    "meme.jpg": "top_right"
                }
            }
        },
        "Strings5": {
            "In python a string is most easily identified by the use of double quotes": {
                "item1": "bar",
                "item2": "foo",
                "image": {
                    "meme.jpg": "top_right"
                }
            }
        },
        "Strings6": {
            "In python a string is most easily identified by the use of double quotes": {
                "item1": "bar",
                "item2": "foo",
                "image": {
                    "meme.jpg": "top_right"
                }
            }
        },
        "Strings7": {
            "In python a string is most easily identified by the use of double quotes": {
                "item1": "bar",
                "item2": "foo",
                "image": {
                    "meme.jpg": "top_right"
                }
            }
        }
    }
    """

    titles = json.loads(data, object_pairs_hook=collections.OrderedDict)

    prs = Presentation(input)

    picture_position = {
        "top_right": {
            "left": 0.3,
            "top": 0.01,
            "width": 0.7
        },
        "bottom_right": {
            "left": 0.3,
            "top": 0.4,
            "width": 0.7
        },
        "top_left": {
            "left": 0.01,
            "top": 0.01,
            "width": 0.7
        },
        "bottom_left": {
            "left": 0.01,
            "top": 0.4,
            "width": 0.7
        },
        "center": {
            "left": 0.15,
            "top": 0.1,
            "width": 0.7
        }
    }

    # Create slides
    for index, (title_key, title_value) in enumerate(titles.iteritems()):
        slide = prs.slides.add_slide(prs.slide_layouts[0])

        title = slide.shapes.title
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame

        title.text = '{}'.format(title_key)
        tf.text = '{}'.format(title_value)

        # Add title to slide
        for slide_title in title_value.keys():
            tf.text = '{}'.format(slide_title)

        # Add content to slide
        for slide_index, slide_value in enumerate(title_value.values()):

            # Add paragraphs to slide
            for content_index, (content_key, content_value) in enumerate(slide_value.items()):
                if "image" not in content_key:
                    p = tf.add_paragraph()
                    p.text = '{}'.format(content_value)
                if "image" in content_key:
                    for image, position in content_value.items():
                        img = imageio.imread(image)
                        picture_left = int(prs.slide_width * picture_position.get(position).get('left'))
                        picture_top = int(prs.slide_height * picture_position.get(position).get('top'))
                        picture_width = int(prs.slide_width * picture_position.get(position).get('width'))

                        picture_height = int(picture_width * img.shape[0] / img.shape[1])
                        picture = slide.shapes.add_picture(image, picture_left, picture_top, picture_width, picture_height)

    prs.save(output)


if __name__ == "__main__":
    args = parse_args()
    analyze_ppt(args.infile.name, args.outfile.name)
