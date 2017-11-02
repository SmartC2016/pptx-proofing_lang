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

