"""
Contains a SheetImagesLoader class that allow you to loadimages from a sheet
This class can get ALL images from a cell containing multiple images
"""

import io
import string

from PIL import Image


class SheetImagesLoader:
    """Loads all images in a sheet"""

    def __init__(self, sheet):
        """Loads all sheet images"""
        
        self._images = {}
        
        sheet_images = sheet._images
        for image in sheet_images:
            row = image.anchor._from.row + 1
            col = string.ascii_uppercase[image.anchor._from.col]
            if f'{col}{row}' not in self._images:
                self._images[f'{col}{row}'] = []
            self._images[f'{col}{row}'].append(image._data)

    def image_in(self, cell):
        """Checks if there's an image in specified cell"""
        return cell in self._images

    def get(self, cell):
        """Retrieves images data from a cell"""
        if cell not in self._images:
            raise ValueError("Cell {} doesn't contain any image".format(cell))
        else:
            images = []
            for img_data in self._images[cell]:
                img_bytes = io.BytesIO(img_data())
                images.append(Image.open(img_bytes))
            return images