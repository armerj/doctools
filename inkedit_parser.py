#!/usr/bin/env python

__description__ = 'Parse Form Control data using oletools.oleform and print properties in /f for inkedit form controls'
__author__ = 'Jon Armer'
__version__ = '0.0.2'
__date__ = '2020/02/08'

"""

Source code put in public domain by Jon Armer, no Copyright
Use at your own risk

This plugin will attempt to extract image and hash it

Usage:
$ python oleform_patched.py -f ../doc_inkedit/ink_default.doc 
{'rtf_data': '{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1033{\\fonttbl{\\f0\\fnil MS Sans Serif;}}\r\n{\\*\\generator Riched20 10.0.18362}\\viewkind4\\uc1 \r\n\\pard\\f0\\fs16 InkEdit1\\par\r\n}\r\n', 'height': 1040, 'RecognTimeOut': 2000, 'backColor': '0x80000005', 'fontname': 'MS Sans Serif', 'cbClassTable': 0, 'mouseIcon': None, 'InkInsertMode': '0 - IEM_InsertText', 'width': 3900, 'version': 2, 'PropMask': 0, 'data_size': 505, 'UseMouseForInput': 0, 'factorid': 'DEFAULT', 'Locked': False, 'font_data': '\x01\x00\x00\x00\x90\x01\xf8$\x01\x00\rMS Sans Serif\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00', 'Enabled': -1, 'ScrollBars': '0 - rtfNone', 'apperance': '1 - rtfThreeD', 'disableNoScroll': False, 'InkMode': '2 - IEM_InkAndGesture', 'MultiLine': False, 'MaxLength': 0, 'borderStyle': '1 - rtfFixedSingle', 'MousePointer': '0 - IMP_Default'}


History:
  2020/02/08: start

Todo:
    - Figureout how to parse cbClassTable
    - Make PR into oletools repo
    - Add option to use actual RTF parser
"""

import olefile
from oletools.oleform import *
import argparse

def consume_inkeditControl(stream):
    PROPERTY_LIST = {"apperance" : ["0 - rtfFlat", "1 - rtfThreeD"], 
                     "borderStyle" : ["0 - rtfNoBorder", "1 - rtfFixedSingle"],
                     "InkInsertMode" : ["0 - IEM_InsertText", "1 - IEM_InsertInk"],
                     "InkMode" : ["0 - IEM_Disabled", "1 - IEM_Ink", "2 - IEM_InkAndGesture"],
                     "MousePointer" : ["0 - IMP_Default", "1 - IMP_Arrow", "2 - IMP_Crosshair", "3 - IMP_Ibeam", 
                                       "4 - IMP_SizeNESW", "5 - IMP_SizeNS", "6 - IMP_SizeNWSE", "7 - IMP_SizeWE", 
                                       "8 - IMP_UpArrow", "9 - IMP_Hourglass", "10 - IMP_NoDrop", "11 - IMP_ArrowHourglass", 
                                       "12 - IMP_ArrowQuestion", "13 - IMP_SizeAll", "14 - IMP_Hand", "99 - IMP_Custom"],
                     "ScrollBars" : ["0 - rtfNone", "1 - rtfHorizontal", "2 - rtfVertical", "3 - rtfBoth"]}
    
    # stream.check_values('LabelControl (versions)', '<BB', 2, (0, 2))
    inkedit_data = {}

    inkedit_data["version"] = stream.unpack("<h", 2)
    inkedit_data["cbClassTable"] = stream.unpack("<h", 2)
    inkedit_data["PropMask"] = stream.unpack("<i", 4)
        
    inkedit_data["data_size"] = stream.unpack("<i", 4)
    unkown_1 = stream.unpack("<i", 4)
    unkown_2 = stream.unpack("<i", 4)
    inkedit_data["width"] = stream.unpack("<i", 4)
    inkedit_data["height"] = stream.unpack("<i", 4)
    inkedit_data["backColor"] = hex(stream.unpack("<I", 4))
    inkedit_data["apperance"] = PROPERTY_LIST["apperance"][stream.unpack("<i", 4)] 
    inkedit_data["borderStyle"] = PROPERTY_LIST["borderStyle"][stream.unpack("<i", 4)]
    inkedit_data["MousePointer"] = PROPERTY_LIST["MousePointer"][stream.unpack("<i", 4)]
    inkedit_data["InkMode"] = PROPERTY_LIST["InkMode"][stream.unpack("<i", 4)]
    inkedit_data["InkInsertMode"] = PROPERTY_LIST["InkInsertMode"][stream.unpack("<i", 4)]
    inkedit_data["RecognTimeOut"] = stream.unpack("<i", 4)
    inkedit_data["Locked"] = stream.unpack("<h", 2) > 0
    inkedit_data["MultiLine"] = stream.unpack("<h", 2) > 0
    inkedit_data["disableNoScroll"] = stream.unpack("<h", 2) > 0
    poss_padding = stream.unpack("<h", 2)
    inkedit_data["ScrollBars"] = PROPERTY_LIST["ScrollBars"][stream.unpack("<i", 4)]
    inkedit_data["Enabled"] = stream.unpack("<h", 2)
    poss_padding = stream.unpack("<h", 2)
    inkedit_data["MaxLength"] = stream.unpack("<i", 4)
    inkedit_data["UseMouseForInput"] = stream.unpack("<h", 2)
    poss_padding = stream.read(6)
    factorid_size = stream.unpack("<i", 4)
    inkedit_data["factorid"] = stream.read(factorid_size).replace("\x00", "") # remove null chars
    mouseIcon_size = stream.unpack("<i", 4)
    if mouseIcon_size > 0:
        inkedit_data["mouseIcon"] = stream.read(mouseIcon_size)
    else:
        inkedit_data["mouseIcon"] = None
    font_data_size = stream.unpack("<i", 4)
    inkedit_data["font_data"] = stream.read(font_data_size)
    fontname_size = struct.unpack(">h", inkedit_data["font_data"][9:11])[0]
    inkedit_data["fontname"] = inkedit_data["font_data"][11:11+fontname_size]
    rtf_data_size = stream.unpack("<i", 4)
    inkedit_data["rtf_data"] = stream.read(rtf_data_size).replace("\x00", "") # remove null chars

    beginning_text = inkedit_data["rtf_data"].find("\\fs16 ") + 6 # TODO add option to use actual RTF parser
    end_text = inkedit_data["rtf_data"][beginning_text:].find("\\par")
    inkedit_data["text"] = inkedit_data["rtf_data"][beginning_text:beginning_text + end_text]

    print inkedit_data


# Functions from oletools.oleform that I patched to handle non-standard form controls;
def extract_OleFormVariables_PATCHED(ole_file, stream_dir):
    control = ExtendedStream.open(ole_file, '/'.join(stream_dir + ['f']))
    variables = list(consume_FormControl(control))
    data = ExtendedStream.open(ole_file, '/'.join(stream_dir + ['o']))
    for var in variables:
        # See FormEmbeddedActiveXControlCached for type definition: [MS-OFORMS] 2.4.5
        if var['ClsidCacheIndex'] == 7:
            consume_FormControl(data)
        elif var['ClsidCacheIndex'] == 12:
            consume_ImageControl(data)
        elif var['ClsidCacheIndex'] == 14:
            consume_FormControl(data)
        elif var['ClsidCacheIndex'] in [15, 23, 24, 25, 26, 27, 28]:
            var['value'], var['caption'], var['group_name'] = consume_MorphDataControl(data)
        elif var['ClsidCacheIndex'] == 16:
            consume_SpinButtonControl(data)
        elif var['ClsidCacheIndex'] == 17:
            consume_CommandButtonControl(data)
        elif var['ClsidCacheIndex'] == 18:
            consume_TabStripControl(data)
        elif var['ClsidCacheIndex'] == 21:
            var['caption'] = consume_LabelControl(data)
        elif var['ClsidCacheIndex'] == 47:
            consume_ScrollBarControl(data)
        elif var['ClsidCacheIndex'] == 57:
            consume_FormControl(data)
        elif var['ClsidCacheIndex'] > 0x7FFF:
            if control.classTable[var['ClsidCacheIndex'] - 0x8000] == "\xf5\x59\xca\xe5\xc4\x57\xd8\x4d\x9b\xd6\x1d\xee\xed\xd2\x7a\xf4":
                consume_inkeditControl(data)
        else:
            # TODO: use logging instead of print
            print('ERROR: Unsupported stored type in user form: {0}'.format(str(var['ClsidCacheIndex'])))
            break
    return variables

# Need to save classTable to determine what control non-standard control is
def ExtendedStream__init__PATCHED(self, stream, path):
    self._pos = 0
    self._jumps = []
    self._stream = stream
    self._path = path
    self._padding = False
    self._pad_start = 0
    self.classTable = []


def consume_SiteClassInfo_PATCHED(stream):
   # SiteClassInfo: [MS-OFORMS] 2.2.10.10.1
   stream.check_value('SiteClassInfo (version)', '<H', 2, 0)
   cbClassTable = stream.unpack('<H', 2)
   stream.classTable.append(stream.read(cbClassTable)[0x18:0x28]) # TODO learn how to parse cbClassTabe


# Replace orginal functions with patched functions
consume_SiteClassInfo = consume_SiteClassInfo_PATCHED
extract_OleFormVariables = extract_OleFormVariables_PATCHED
ExtendedStream.__init__ = ExtendedStream__init__PATCHED

my_argparser = argparse.ArgumentParser()
my_argparser.add_argument("-f", "--file", type=str, help="Document to extract files from")

args = my_argparser.parse_args()

ole = olefile.OleFileIO(args.file)

# Call parser
extract_OleFormVariables(ole, ['Macros','UserForm1'])

