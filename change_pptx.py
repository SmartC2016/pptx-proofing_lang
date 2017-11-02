# -*- coding: utf-8 -*-
"""
This program is exchanging / setting the proofing language for any text-containing object
in a given Powerpoint Presentation.

My usecase:

I often review / add to existing Powerpoint presentations coming from colleagues all over the world.
Our company language is british english. However, very often the people do not change the proofing language
in the text elements, so that every english text is underlined and assumed wrong, if you for example have
swedish as proofing language. Often our presentations are 50 and more pages with title elements, sub-title
and text. To change all the elements to english is very cumbersome and for sure annoying.

Then there was the idea, to 'fix' this issue using the awesome Python programming language! :-) Yeah!

I use the great python-pptx module from Steve Canny - a million thanks to Steve!

This program allows you to select your pptx, select a language that should be set to all text-containing objects and
save your pptx under a new name.

if you have any questions, please don't hesitate to come back to me.
"""

__author__ = "Christian Hetmann"

#todo create a tkinter window with the following features
#todo -- open file dialog: open a pptx file
#todo -- select a language from a dropdown menu that you want to set in the complete presentation
#todo ---- (default ENGLISH_UK, because that is my favourite :-) )
#todo -- create a little window for "logging" / showing what the program has found in the given pptx
#todo -- have a start button for execution
#todo -- have a save button to save the presentation
#todo create list with all existing languages in order to populate the dropdown


from pptx import Presentation
from pptx.enum.lang import MSO_LANGUAGE_ID

# select ENGLISH_UK as the new language to be set - this should be changed in the future to pick any language
new_language = MSO_LANGUAGE_ID.ENGLISH_UK

input_file = 'test_pptx.pptx'
output_file = input_file[:-5] + '_modified.pptx'

# Open the presentation
prs = Presentation(input_file)

# iterate through all slides
for slide_no, slide in enumerate(prs.slides):
    # iterate through all shapes/objects on one slide
    for shape in slide.shapes:
        # check if the shape/object has text (pictures e.g. don't have text)
        if shape.has_text_frame:
            # print some output to the console for now
            print('SLIDE NO# ', slide_no + 1)
            print('Object-Name: ', shape.name)
            print('Text -->', shape.text)
            # check for each paragraph of text for the actual shape/object
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    # display the current language
                    print('Actual set language: ', run.font.language_id)
                    # set the 'new_language'
                    run.font.language_id = new_language
        else:
            print('SLIDE NO# ', slide_no + 1, ': This object "', shape.name, '" has no text.')
        print(' +++++ next element +++++ ')
    print('--------- next slide ---------')

# save pptx with new filename
prs.save(output_file)
