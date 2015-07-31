#!/usr/local/bin/python
# -*- coding: latin-1 -*-
"""
ExtractMsg:
    Extracts emails and attachments saved in Microsoft Outlook's .msg files

https://github.com/mattgwwalker/msg-extractor
"""

__author__ = "Matthew Walker"
__date__ = "2013-11-19"
__version__ = '0.2'

# --- LICENSE -----------------------------------------------------------------
#
#    Copyright 2013 Matthew Walker, 2015 Accuvant, Inc. (bspengler@accuvant.com)
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
import os
import sys
import glob
import traceback
from email.parser import Parser as EmailParser
import email.utils
import olefile as OleFile
from lib.cuckoo.common.utils import store_temp_file
def windowsUnicode(string):
    if string is None:
        return None
    if sys.version_info[0] >= 3:  # Python 3
        return str(string, 'utf_16_le')
    else:  # Python 2
        return unicode(string, 'utf_16_le')
INTERESTING_FILE_EXTENSIONS = [
    '', '.bat', '.bin', '.cmd', '.com', '.cpl', '.dll', '.doc', '.docb', '.docm', '.docx', '.dot', '.dotm', '.dotx', '.exe', '.hta', '.htm', '.html',
    '.jar', '.msc', '.msi', '.msp', '.mst', '.pdf', '.pif', '.pot', '.potm', '.potx', '.ppam', '.pps', '.ppsm', '.ppsx', '.ppt', '.pptm', '.pptx',
    '.ps1', '.ps1xml', '.ps2', '.ps2xml', '.psc1', '.psc2', '.reg', '.rgs', '.scr', '.sct', '.shb', '.shs', '.sldm', '.sldx', '.vb', '.vba', '.vbe',
    '.vbs', '.vbscript', '.ws', '.wsh', '.xla', '.xlam', '.xll', '.xlm', '.xls', '.xlsb', '.xlsm', '.xlsx', '.xlt', '.xltm', '.xltx', '.xlw', '.zip'
]

class Attachment:
    def __init__(self, msg, dir_):

        # Get long filename
        self.longFilename = msg._getStringStream(dir_ + ['__substg1.0_3707'])

        # Get short filename
        self.shortFilename = msg._getStringStream(dir_ + ['__substg1.0_3704'])

        # Get attachment data
        self.data = msg._getStream(dir_ + ['__substg1.0_37010102'])

        print "logfilename %s"%self.longFilename
    def save(self):
        # Use long filename as first preference
        filename = self.longFilename
        # Otherwise use the short filename
        if filename is None:
            filename = self.shortFilename
        # Otherwise just make something up!
        if filename is None:
            import random
            import string
            filename = 'UnknownAttachment' + \
                ''.join(random.choice(string.ascii_uppercase + string.digits)
                        for _ in range(5)) + ".bin"

        base, ext = os.path.splitext(filename)
        basename = os.path.basename(filename)
        ext = ext.lower()
        if ext == "" and len(basename) and basename[0] == ".":
            return None
        extensions = INTERESTING_FILE_EXTENSIONS
        foundext = False
        for theext in extensions:
            if ext == theext:
                foundext = True
                break

        if not foundext:
            return None

        return store_temp_file(self.data, filename)


class Message(OleFile.OleFileIO):
    def __init__(self, filename):
        OleFile.OleFileIO.__init__(self, filename)

    def _getStream(self, filename):
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def _getStringStream(self, filename, prefer='unicode'):
        """Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then the parameter /prefer/ specifies which will be
        returned.
        """

        if isinstance(filename, list):
            # Join with slashes to make it easier to append the type
            filename = "/".join(filename)

        asciiVersion = self._getStream(filename + '001E')
        unicodeVersion = windowsUnicode(self._getStream(filename + '001F'))
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion

    @property
    def body(self):
        # Get the message body
        return self._getStringStream('__substg1.0_1000')

    @property
    def attachments(self):
        try:
            return self._attachments
        except Exception:
            # Get the attachments
            attachmentDirs = []
            dirList = self.listdir()
            """
            The purpose is to get the most nested attachment dir in order to handle msgs inside msgs
            In this case, the second message will be in an __attach in "__substg1.0_3701000D"
                cf fig 3. https://msdn.microsoft.com/en-us/library/ee217698%28v=exchg.80%29.aspx
            """
            dirList = sorted(dirList,key=lambda dir: len(dir), reverse=True) # Used to gets the most nested attachment
            for dir_ in dirList:
                if dir_[0].startswith('__attach'):
                    result = [dir_[0]]
                    for d in dir_[1:]:
                        if (d == "__substg1.0_3701000D") or (d.startswith('__attach')):
                           result.append(d)
                        else:
                            break
                    if len([d for d in [''.join(d) for d in attachmentDirs] if d.startswith("".join(result))]) == 0:
                        # Reject if a more specific dir is known
                        attachmentDirs.append(result)

            self._attachments = []

            for attachmentDir in attachmentDirs:
                self._attachments.append(Attachment(self, attachmentDir))

            return self._attachments

    def get_extracted_attachments(self):
        retlist = []
        # Save the attachments
        print "get_extracted_attachments"
        for attachment in self.attachments:
            print attachment
            saved = attachment.save()
            if saved:
                retlist.append(saved)
        return retlist
