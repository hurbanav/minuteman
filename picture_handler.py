from PIL import Image
from datetime import datetime


def resize_image(input_path, output_path, new_size):
    with Image.open(input_path) as img:
        # Resize the image
        img = img.resize(new_size, Image.Resampling.LANCZOS,)
        # Save the image back to disk
        img.save(output_path, "PNG")
        print(f'{datetime.now()}' + f'Image {input_path} was resized.')


def convert_to_black_and_white(input_path, output_path):
    """
    Converts an image to black and white (grayscale).

    :param input_path: Path to the input image.
    :param output_path: Path where the black and white image will be saved.
    """
    with Image.open(input_path) as img:
        # Convert image to grayscale
        bw_img = img.convert('L')
        # Save the converted image
        bw_img.save(output_path, "PNG")
        print(f'{datetime.now()}' + f'Image {input_path} was converted to B&W.')
