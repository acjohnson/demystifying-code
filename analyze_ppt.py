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
        "Programming languages": {
            "": {
                "item1": "Can be thought of as simply a program that you feed commands to",
                "item2": "You can write scripts or programs that make the computer do whatever you can dream up",
                "item3": "There are many programming languages at least 700 or so according to Wikipedia",
                "image1": {
                    "slide1.jpg": "bottom"
                }
            }
        },
        "Brief history of programming languages": {
            "": {
                "item1": "1957 - FORTRAN - Compiled",
                "item2": "1964 - BASIC - Interpreted",
                "item3": "1970 - Pascal - Compiled",
                "item4": "1972 - C - Compiled",
                "item5": "1980 - C++ - Compiled",
                "item7": "1991 - Python - Interpreted",
                "item8": "1991 - Visual Basic - Compiled",
                "item10": "1995 - Ruby - Interpreted",
                "item11": "1995 - Java - Compiled (JVM)",
                "item12": "1995 - JavaScript - Interpreted",
                "item13": "1995 - PHP - Interpreted",
                "item14": "2001 - C# - Compiled (CLR)",
                "item15": "2009 - Go - Compiled (Google - produces statically linked native binaries without external dependencies.)",
                "item16": "2011 - Dart Compiled/Interpreted (Google - AOT-compiled to JavaScript)",
                "image1": {
                    "slide2.jpg": "right"
                }
            }
        },
    	"Interpreted vs Compiled": {
            "": {
                "item1": "Compiled languages - converted directly into machine code that the processor can execute. As a result, they tend to be faster and more efficient to execute than interpreted languages.",
                "item2": "Interpreted languages - the source code is not directly translated by the target machine. Instead, a different program, aka the interpreter, reads and executes the code.",
                "image1": {
                    "slide3.jpg": "bottom"
                }
            }
        },
        "Scripts vs Programs": {
            "": {
                "item1": "Scripts are usually interpreted (but not always, such as with golang scripts)",
                "item2": "Programs can be either compiled (C++, C#, Java) or interpreted (Python, Ruby, PHP)",
                "item3": "The biggest difference is that scripts are written to control an existing program",
                "item4": "Scripts often automate manual tasks to make work easier to accomplish",
                "item5": "Scripts can accomplish many important tasks and are often written by a single person",
                "item6": "Programs usually have very ambitious goals and often take a large amount of time and money to create",
                "image1": {
                    "slide4.jpg": "bottom_right"
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

    picture_position = {
        "top": {
            "left": 0.15,
            "top": 0.01,
            "width": 0.7
        },
        "bottom": {
            "left": 0.32,
            "top": 0.57,
            "width": 0.35
        },
        "right": {
            "left": 0.55,
            "top": 0.3,
            "width": 0.4
        },
        "top_right": {
            "left": 0.3,
            "top": 0.01,
            "width": 0.7
        },
        "bottom_right": {
            "left": 0.6,
            "top": 0.7,
            "width": 0.3
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

    titles = json.loads(data, object_pairs_hook=collections.OrderedDict)
    prs = Presentation(input)

    # Create slides
    for index, (title_key, title_value) in enumerate(titles.iteritems()):
        # Choose the slide layout
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        title = slide.shapes.title
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame

        title.text = '{}'.format(title_key)

        # Add sub-title to slide
        for slide_title in title_value.keys():
            tf.text = '{}'.format(slide_title)

        # Add content to slide
        for slide_index, slide_value in enumerate(title_value.values()):

            # Add paragraphs and images to slide
            for content_index, (content_key, content_value) in enumerate(slide_value.items()):

                # Add bullets
                if "item" in content_key:
                    p = tf.add_paragraph()
                    p.text = '{}'.format(content_value)

                # Add images
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
