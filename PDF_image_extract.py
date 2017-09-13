# coding: utf-8

"""
Script to extract images from all PDFs in a directory

The Python PDF Toolkit
Copyright ©2016 Ronan Paixão
Licensed under the terms of the MIT License.

Links:
PDF format: http://www.adobe.com/content/dam/Adobe/en/devnet/acrobat/pdfs/pdf_reference_1-7.pdf
CCITT Group 4: https://www.itu.int/rec/dologin_pub.asp?lang=e&id=T-REC-T.6-198811-I!!PDF-E&type=items
Extract images from pdf: http://stackoverflow.com/questions/2693820/extract-images-from-pdf-without-resampling-in-python
Extract images coded with CCITTFaxDecode in .net: http://stackoverflow.com/questions/2641770/extracting-image-from-pdf-with-ccittfaxdecode-filter
TIFF format and tags: http://www.awaresystems.be/imaging/tiff/faq.html

originally authored by: Ronan Paixão, with some code from
    http://stackoverflow.com/questions/2693820/extract-images-from-pdf-without-resampling-in-python

Adapted by Alexander Tomlinson with several modifications to traverse folders.
See original at: https://github.com/ronanpaixao/PyPDFTK/blob/master/pdf_images.py
"""

# python 2/3 compatibility
from __future__ import print_function
import future
from builtins import input

import struct
from pathlib import Path

import PyPDF2
from PIL import Image

__author__ = 'Alexander Tomlinson'
__email__ = 'tomlinsa@ohsu.edu'

img_modes = {'/DeviceRGB': 'RGB', '/DefaultRGB': 'RGB',
             '/DeviceCMYK': 'CMYK', '/DefaultCMYK': 'CMYK',
             '/DeviceGray': 'L', '/DefaultGray': 'L',
             '/Indexed': 'P'}


def tiff_header_for_CCITT(width, height, img_size, CCITT_group=4):
    tiff_header_struct = '<' + '2s' + 'h' + 'l' + 'h' + 'hhll' * 8 + 'h'
    return struct.pack(tiff_header_struct,
                       b'II',
                       42,                                              # Version number (always 42)
                       8,                                               # Offset to first IFD
                       8,                                               # Number of tags in IFD
                       256, 4, 1, width,                                # ImageWidth, LONG, 1, width
                       257, 4, 1, height,                               # ImageLength, LONG, 1, lenght
                       258, 3, 1, 1,                                    # BitsPerSample, SHORT, 1, 1
                       259, 3, 1, CCITT_group,                          # Compression, SHORT, 1, 4 = CCITT Group 4 fax encoding
                       262, 3, 1, 0,                                    # Threshholding, SHORT, 1, 0 = WhiteIsZero
                       273, 4, 1, struct.calcsize(tiff_header_struct),  # StripOffsets, LONG, 1, len of header
                       278, 4, 1, height,                               # RowsPerStrip, LONG, 1, lenght
                       279, 4, 1, img_size,                             # StripByteCounts, LONG, 1, size of image
                       0)                                               # last IFD


def extract_images(pdf_path):
    """
    Extracts the images from a page of a PDF and saves them.

    :param pdf_path: path of PDF
    """
    pdf_file = PyPDF2.PdfFileReader(open(str(pdf_path), 'rb'))
    pdf_name = pdf_path.stem
    filename_prefix = pdf_name + '_IMG_'

    pdf_folder = pdf_path.parent
    img_folder = Path.joinpath(pdf_folder, 'extracted_images')
    if not Path.exists(img_folder):
        Path.mkdir(img_folder)

    print('\n{}'.format(pdf_name))

    for i in range(pdf_file.getNumPages()):
        print('\tpage {}'.format(i+1))
        page = pdf_file.getPage(i)

        try:
            # if page has no images, skip it
            xObject = page['/Resources']['/XObject'].getObject()
        except KeyError:
            continue

        j = 1  # image number on page

        for obj in xObject:

            if xObject[obj]['/Subtype'] == '/Image':

                size = (xObject[obj]['/Width'], xObject[obj]['/Height'])
                color_space = xObject[obj]['/ColorSpace']

                # colorspace fixes
                if isinstance(color_space, PyPDF2.generic.ArrayObject) and color_space[0] == '/Indexed':
                    color_space, base, hival, lookup = [v.getObject() for v in color_space]  # pg 262
                if isinstance(color_space, PyPDF2.generic.ArrayObject) and color_space[0] == '/ICCBased':
                    a = [v.getObject() for v in color_space]
                    # use alternate for simplicity
                    if '/Alternate' in a[1].keys():
                        color_space = a[1]['/Alternate']
                    else:
                        raise IOError('color space non existent: {}'.format(a))

                # lookup color space
                mode = img_modes[color_space]

                img_name = '{}page{:02}_{:04}.png'.format(filename_prefix, i+1, j)
                print('\t\t{}'.format(img_name))
                img_name = str(Path.joinpath(img_folder, img_name))

                if xObject[obj]['/Filter'] == '/FlateDecode':
                    data = xObject[obj].getData()
                    img = Image.frombytes(mode, size, data)
                    if color_space == '/Indexed':
                        img.putpalette(lookup.getData())
                        img = img.convert('RGB')
                    img.save(img_name)

                elif xObject[obj]['/Filter'] == '/DCTDecode':
                    data = xObject[obj]._data
                    img = open(img_name, "wb")
                    img.write(data)
                    img.close()

                elif xObject[obj]['/Filter'] == '/JPXDecode':
                    data = xObject[obj]._data
                    img = open(img_name, "wb")
                    img.write(data)
                    img.close()

                # The  CCITTFaxDecode filter decodes image data that has been encoded using
                # either Group 3 or Group 4 CCITT facsimile (fax) encoding. CCITT encoding is
                # designed to achieve efficient compression of monochrome (1 bit per pixel) image
                # data at relatively low resolutions, and so is useful only for bitmap image data, not
                # for color images, grayscale images, or general data.
                #
                # K < 0 --- Pure two-dimensional encoding (Group 4)
                # K = 0 --- Pure one-dimensional encoding (Group 3, 1-D)
                # K > 0 --- Mixed one- and two-dimensional encoding (Group 3, 2-D)

                elif xObject[obj]['/Filter'] == '/CCITTFaxDecode':
                    if xObject[obj]['/DecodeParms']['/K'] == -1:
                        CCITT_group = 4
                    else:
                        CCITT_group = 3

                    width = xObject[obj]['/Width']
                    height = xObject[obj]['/Height']

                    data = xObject[obj]._data  # getData() does not work for CCITTFaxDecode
                    img_size = len(data)

                    tiff_header = tiff_header_for_CCITT(width, height, img_size, CCITT_group)
                    with open(img_name, 'wb') as img_file:
                        img_file.write(tiff_header + data)

                else:
                    print('Unable to save image; unrecognized format!')

                j += 1


if __name__ == '__main__':
    while True:
        try:
            pdf_dir = input('\nEnter directory of PDFs, or path of single PDF: ')
            assert isinstance(pdf_dir, str)
            pdf_dir = Path(pdf_dir)

            if Path.exists(pdf_dir):
                break
            else:
                print('File or directory does not exist. Try again. (ctrl-c to quit)')

        except:
            print('File or directory not recognized. Try again. (ctrl-c to quit)')
            raise

    if pdf_dir.suffix == '.pdf':
        extract_images(pdf_dir)
    else:
        pdfs = list(pdf_dir.glob('**/*.pdf'))
        for pdf in pdfs:
            extract_images(pdf)