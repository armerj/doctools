# doctools

extract_img usage:
```
$ python plugin_extract_img.py -f ../test_docs/image_in_doc.doc -s test
['image sha256 hash is: 5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf']

$ file test/5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf 
test/5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf: JPEG image data, JFIF standard 1.02, resolution (DPI), density 300x300, segment length 16, Exif Standard: [TIFF image data, big-endian, direntries=7, orientation=upper-left, xresolution=98, yresolution=106, resolutionunit=2, software=Adobe Photoshop CS3 Windows, datetime=2008:07:01 09:49:29], baseline, precision 8, 2170x1560, components 3

$ sha256sum test/5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf 
5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf  test/5737761889ed2d709d00a65d84cfe4dee120c8c2d98054e5fff073652021aaaf
```

inkedit_parser usage: 
```
$ python oleform_patched.py -f ../doc_inkedit/ink_default.doc 
{'rtf_data': '{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1033{\\fonttbl{\\f0\\fnil MS Sans Serif;}}\r\n{\\*\\generator Riched20 10.0.18362}\\viewkind4\\uc1 \r\n\\pard\\f0\\fs16 InkEdit1\\par\r\n}\r\n', 'height': 1040, 'RecognTimeOut': 2000, 'backColor': '0x80000005', 'fontname': 'MS Sans Serif', 'cbClassTable': 0, 'mouseIcon': None, 'InkInsertMode': '0 - IEM_InsertText', 'width': 3900, 'version': 2, 'PropMask': 0, 'data_size': 505, 'UseMouseForInput': 0, 'factorid': 'DEFAULT', 'Locked': False, 'font_data': '\x01\x00\x00\x00\x90\x01\xf8$\x01\x00\rMS Sans Serif\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00', 'Enabled': -1, 'ScrollBars': '0 - rtfNone', 'apperance': '1 - rtfThreeD', 'disableNoScroll': False, 'InkMode': '2 - IEM_InkAndGesture', 'MultiLine': False, 'MaxLength': 0, 'borderStyle': '1 - rtfFixedSingle', 'MousePointer': '0 - IMP_Default'}
```
