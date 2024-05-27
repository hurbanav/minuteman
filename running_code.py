#todo The idea here is to make an app that transforms pdfs into apps of the prime calculator, heres what we need:

"""
    - tkinter utils class to make the app (we'll need a select file button, pages button (create a function to select
      the pages after converting by deleting the files not wanted), app name, black and white and resize dimensions.
    - I'll also need to create a function that duplicates the files for the program to work, and rename the directory
      to the programs' name
    - Need a function to pack the final directory into a single zip file and erase the rest
    - don't forget to print the file size (function to tell me the file size) as a label as the end of the program running
"""

import pdf_manager

pdf_manager.pdf_to_png(r'C:\Users\henri\Downloads\testinhos sanches gestão ambiental 4.pdf')

