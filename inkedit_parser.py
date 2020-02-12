#!/usr/bin/env python

__description__ = 'Parse Form Control data using oletools.oleform and print properties in /f for inkedit form controls'
__author__ = 'Jon Armer'
__version__ = '0.0.4'
__date__ = '2020/02/09'

"""

Source code put in public domain by Jon Armer, no Copyright
Use at your own risk

This plugin will attempt to extract image and hash it

Usage:
$ python oleform_patched.py -f ../doc_inkedit/ink_default.doc 
{'rtf_data': '{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1033{\\fonttbl{\\f0\\fnil MS Sans Serif;}}\r\n{\\*\\generator Riched20 10.0.18362}\\viewkind4\\uc1 \r\n\\pard\\f0\\fs16 InkEdit1\\par\r\n}\r\n', 'height': 1040, 'RecognTimeOut': 2000, 'backColor': '0x80000005', 'fontname': 'MS Sans Serif', 'cbClassTable': 0, 'mouseIcon': None, 'InkInsertMode': '0 - IEM_InsertText', 'width': 3900, 'version': 2, 'PropMask': 0, 'data_size': 505, 'UseMouseForInput': 0, 'factorid': 'DEFAULT', 'Locked': False, 'font_data': '\x01\x00\x00\x00\x90\x01\xf8$\x01\x00\rMS Sans Serif\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00', 'Enabled': -1, 'ScrollBars': '0 - rtfNone', 'apperance': '1 - rtfThreeD', 'disableNoScroll': False, 'InkMode': '2 - IEM_InkAndGesture', 'MultiLine': False, 'MaxLength': 0, 'borderStyle': '1 - rtfFixedSingle', 'MousePointer': '0 - IMP_Default'}


History:
  2020/02/08: start
  2020/02/12: Added cbClassTable Parser, and put inkedit parsing into class 

Todo:
    - Make PR into oletools repo
    - Add option to use actual RTF parser
"""

import olefile
from oletools.oleform import *
import argparse
from pprint import pprint

class inkeditControl():
    PROPERTY_LIST = {"apperance" : ["0 - rtfFlat", "1 - rtfThreeD"], 
                     "borderStyle" : ["0 - rtfNoBorder", "1 - rtfFixedSingle"],
                     "InkInsertMode" : ["0 - IEM_InsertText", "1 - IEM_InsertInk"],
                     "InkMode" : ["0 - IEM_Disabled", "1 - IEM_Ink", "2 - IEM_InkAndGesture"],
                     "MousePointer" : ["0 - IMP_Default", "1 - IMP_Arrow", "2 - IMP_Crosshair", "3 - IMP_Ibeam", 
                                       "4 - IMP_SizeNESW", "5 - IMP_SizeNS", "6 - IMP_SizeNWSE", "7 - IMP_SizeWE", 
                                       "8 - IMP_UpArrow", "9 - IMP_Hourglass", "10 - IMP_NoDrop", "11 - IMP_ArrowHourglass", 
                                       "12 - IMP_ArrowQuestion", "13 - IMP_SizeAll", "14 - IMP_Hand", "99 - IMP_Custom"],
                     "ScrollBars" : ["0 - rtfNone", "1 - rtfHorizontal", "2 - rtfVertical", "3 - rtfBoth"]}
    
    inkedit_fields = [("version", "<h", 2), ("cbClassTable", "<h", 2), ("PropMask", "<i", 4), ("data_size", "<i", 4), 
                                         ("unkown_1", "<I", 4, "hex"), ("unkown_2", "<I", 4, "hex"), ("width", "<i", 4), ("height", "<i", 4), 
                                         ("backColor", "<I", 4, "hex"), ("apperance", "<i", 4, "prop"), ("borderStyle", "<i", 4, "prop"), 
                                         ("MousePointer", "<i", 4, "prop"), ("InkMode", "<i", 4, "prop"), ("InkInsertMode", "<i", 4, "prop"),
                                         ("RecognTimeOut", "<i", 4), ("Locked", "<h", 2, "bool"), ("MultiLine", "<h", 2, "bool"), 
                                         ("disableNoScroll", "<h", 2, "bool"), ("padding", "<bb", 2), ("ScrollBars", "<i", 4, "prop"),
                                         ("Enabled", "<h", 2, "bool"), ("padding", "<bb", 2), ("MaxLength", "<i", 4), ("UseMouseForInput", "<h", 2, "bool"), 
                                         ("padding", "<bbbbbb", 6), ("factorid", "<i", 4, "unicode"), ("mouseIcon", "<i", 4, "image"),
                                         ("font", "<i", 4, "font"), ("rtf_data", "<i", 4, "rtf")]
    # Note: Padding is likely simliar to that in other MS-OFORMS objects, where data of 4 bytes needs to be align to offsets that are multiple of 4
    # Though it looks like all fields are required here, so having hardcoded padding sizes should be ok 

    property_values = {}
    

    def __init__(self, stream):
        for field in self.inkedit_fields:
            field_value = stream.unpack(field[1], field[2])
            if len(field) == 3:
                self.property_values[field[0]] = field_value
            else:
                if field[3] == "bool":
                    self.property_values[field[0]] = field_value > 0
                if field[3] == "unicode":
                    unicodedata = stream.read(field_value)
                    self.property_values[field[0]] = unicodedata.replace("\x00", "")
                elif field[3] == "prop":
                    self.property_values[field[0]] = self.PROPERTY_LIST[field[0]][field_value]
                elif field[3] == "font":
                    fontdata = stream.read(field_value)
                    self.property_values["fontdata"] = fontdata
                    fontname_size = struct.unpack(">h", fontdata[9:11])[0]  # TODO add in font parsing
                    self.property_values["fontname"] = fontdata[11:11+fontname_size]
                elif field[3] == "rtf":
                    self.property_values["rtf_data"] = stream.read(field_value).replace("\x00", "") # remove null chars
                    beginning_text = self.property_values["rtf_data"].find("\\fs16 ") + 6 # TODO add option to use actual RTF parser
                    end_text = self.property_values["rtf_data"][beginning_text:].find("\\par")
                    self.property_values["text"] = self.property_values["rtf_data"][beginning_text:beginning_text + end_text]
                elif field[3] == "image":
                    if field_value > 0:
                        self.property_values[field[0]] = stream.read(field_value)
                elif field[3] == "hex":
                    self.property_values[field[0]] = hex(field_value)



class ClassInfoPropMask(Mask):
    """ClassInfoPropMask: [MS-OFORMS] 2.2.10.10.2"""
    _size = 15
    _read_size = 0
    _names = ['fClsID', 'fDispEvent', 'Unused1', 'fDefaultProg',
              'fClassFlags', 'fCountOfMethods', 'fDispidBind', 'fGetBindIndex', 'fPutBindIndex',
              'fBindType', 'fGetValueIndex', 'fPutValueIndex', 'fValueType', 'fDispidRowset', 'fSetRowset']


    def reset_read(self):
        temp_read_size = self._read_size
        self._read_size = 0
        return temp_read_size

    def consume(self, stream, props):
        for (name, size) in props:
            if self[name]:
                stream.read(size)
                self._read_size += size


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
            #print control.classTable
            if control.classTable[var['ClsidCacheIndex'] - 0x8000] == "\xf5\x59\xca\xe5\xc4\x57\xd8\x4d\x9b\xd6\x1d\xee\xed\xd2\x7a\xf4":
                control_data = inkeditControl(data).property_values
                var['raw_data'] = control_data
                var['factorid'] = control_data['factorid']
                var['rtf_data'] = control_data['rtf_data']
                var['text'] = control_data['text']
                var['non_ms_type'] = "InkEdit"
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

    with stream.will_jump_to(cbClassTable):
        temp = stream.unpack("<L", 4)
        propmask = ClassInfoPropMask(temp)

        # ClassInfoDataBlock: [MS-OFORMS] 2.2.10.10.3
        propmask.consume(stream, [('fClassFlags', 2), ('fClassFlags', 2), # If fClassFlags is set then ClassTableFlags and VarFlags exists
                                  ('fCountOfMethods', 4), ('fDispidBind', 4), ('fGetBindIndex', 2), 
                                  ('fPutBindIndex', 2), ('fBindType', 2), ('fGetValueIndex', 2), ('fPutValueIndex', 2), 
                                  ('fValueType', 2)])
        
        padding1_expected = propmask.reset_read() % 4
        stream.read(padding1_expected)
        propmask.consume(stream, [('fDispidRowset', 4), ('fSetRowset', 2)])
        padding2_expected = propmask.reset_read() % 4
        stream.read(padding2_expected)

        # ClassInfoExtraDataBlock: [MS-OFORMS] 2.2.10.10.5
        if propmask['fClsID']:
            clsid = stream.read(16)
            stream.classTable.append(clsid)
        propmask.consume(stream, [('fDispEvent', 16), ('fDefaultProg', 16)])


# Replace orginal functions with patched functions
consume_FormControl.func_globals['consume_SiteClassInfo'] = consume_SiteClassInfo_PATCHED
extract_OleFormVariables = extract_OleFormVariables_PATCHED
ExtendedStream.__init__ = ExtendedStream__init__PATCHED


if __name__ == "__main__":
    my_argparser = argparse.ArgumentParser()
    
    my_argparser.add_argument("-f", "--file", type=str, help="Document to extract files from")

    args = my_argparser.parse_args()

    ole = olefile.OleFileIO(args.file)
    dirs = ole.listdir()

    # Call parser
    for dir in dirs:
        if dir[-1] == "f" and (dir[:2] + ["o"]) in dirs:
            print dir
            controls = extract_OleFormVariables(ole, dir[:2])
            for control in controls:
                if "non_ms_type" in control and control["non_ms_type"] == "InkEdit":
                    pprint(control)
