# doctools

Usage:
$ python plugin_extract_img.py -f ../test_docs/image_in_doc.doc -s test
['image sha256 hash is: 5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf']

$ file test/5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf 
test/5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf: JPEG image data, JFIF standard 1.02, resolution (DPI), density 300x300, segment length 16, Exif Standard: [TIFF image data, big-endian, direntries=7, orientation=upper-left, xresolution=98, yresolution=106, resolutionunit=2, software=Adobe Photoshop CS3 Windows, datetime=2008:07:01 09:49:29], baseline, precision 8, 2170x1560, components 3

$ sha256sum test/5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf 
5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf  test/5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf
