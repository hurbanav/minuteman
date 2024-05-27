from pdf2image import convert_from_path
import os
import fitz  # PyMuPDF


def pdf_to_png(pdf_filepath):
    """
    Converts each page of a PDF file to a PNG image.

    :param pdf_filepath: Path to the PDF file.
    """
    # Extract directory and base for naming
    output_dir = os.path.dirname(pdf_filepath + 'pdf_to_png')

    # Convert PDF to a list of images
    images = convert_from_path(pdf_filepath)

    # Loop through images and save them to the same directory as the PDF
    for i, image in enumerate(images):
        # Format page number with leading zeros
        filename = os.path.join(output_dir, f"{i + 1:03}.png")
        image.save(filename, 'PNG')
        print(f"Saved: {filename}")

def pdf_to_png(pdf_filepath):
    """
    Converts each page of a PDF file to a PNG image using PyMuPDF and saves it in a subdirectory.
    :param pdf_filepath: Path to the PDF file.
    :return output directory to be used if needed.
    """
    # Open the provided PDF file
    doc = fitz.open(pdf_filepath)

    # Calculate output directory path
    base_dir = os.path.dirname(pdf_filepath)
    output_dir = os.path.join(base_dir,
                              os.path.basename(pdf_filepath) + ' - pdf_to_png')

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Iterate through each page
    for page_number in range(doc.page_count):
        page = doc.load_page(page_number)  # number of page
        pix = page.get_pixmap()  # render page to an image
        image_filename = os.path.join(output_dir, f"{page_number + 1:03}.png")
        pix.save(image_filename)  # save image as a PNG
        print(f"Saved: {image_filename}")

    # Close the document
    doc.close()
    return output_dir