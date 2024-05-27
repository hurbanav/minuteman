import picture_handler
import file_manager

image_dir_list = file_manager.list_files(directory=r'C:\Users\henri\OneDrive\Documents\HP Connectivity Kit\Conteúdo\Silos\Sales\Sales.hpappdir',
                                         extension='.png')

for image in image_dir_list:
    if 'icon' in image:
        picture_handler.resize_image(input_path=image,
                                     output_path=image,
                                     new_size=(38, 38))
    else:
        picture_handler.resize_image(input_path=image,
                                     output_path=image,
                                     new_size=(800, 600))
    picture_handler.convert_to_black_and_white(input_path=image,
                                               output_path=image)