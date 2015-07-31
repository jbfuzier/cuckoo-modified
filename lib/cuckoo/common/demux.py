# Copyright (C) 2015 Accuvant, Inc. (bspengler@accuvant.com)
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

import os
import tempfile
from zipfile import ZipFile
try:
    from rarfile import RarFile
    HAS_RARFILE = True
except ImportError:
    HAS_RARFILE = False

from lib.cuckoo.common.config import Config
from lib.cuckoo.common.objects import File
from lib.cuckoo.common.email_utils import find_attachments_in_email
from lib.cuckoo.common.office.msgextract import Message
import logging
log = logging.getLogger(__name__)
INTERESTING_FILE_EXTENSIONS = [
    '', '.bat', '.bin', '.cmd', '.com', '.cpl', '.dll', '.doc', '.docb', '.docm', '.docx', '.dot', '.dotm', '.dotx', '.exe', '.hta', '.htm', '.html',
    '.jar', '.msc', '.msi', '.msp', '.mst', '.pdf', '.pif', '.pot', '.potm', '.potx', '.ppam', '.pps', '.ppsm', '.ppsx', '.ppt', '.pptm', '.pptx',
    '.ps1', '.ps1xml', '.ps2', '.ps2xml', '.psc1', '.psc2', '.reg', '.rgs', '.scr', '.sct', '.shb', '.shs', '.sldm', '.sldx', '.vb', '.vba', '.vbe',
    '.vbs', '.vbscript', '.ws', '.wsh', '.xla', '.xlam', '.xll', '.xlm', '.xls', '.xlsb', '.xlsm', '.xlsx', '.xlt', '.xltm', '.xltx', '.xlw', '.zip', '.msg'
]

KEEP_INTERMEDIATE_FILES = False # Move as a setting in config ?

def demux_zip(filename, options):
    retlist = []

    log.error("test")
    try:
        # don't try to extract from office docs
        magic = File(filename).get_type()
        if "Microsoft" in magic or "Java Jar" in magic or "Composite Document File" in magic:
            return retlist

        extracted = []
        password="infected"
        fields = options.split(",")
        for field in fields:
            try:
                key, value = field.split("=", 1)
                if key == "password":
                    password = value
                    break
            except:
                pass

        with ZipFile(filename, "r") as archive:
            infolist = archive.infolist()
            for info in infolist:
                # avoid obvious bombs
                if info.file_size > 100 * 1024 * 1024 or not info.file_size:
                    continue
                # ignore directories
                if info.filename.endswith("/"):
                    continue
                print info.filename
                base, ext = os.path.splitext(info.filename)
                basename = os.path.basename(info.filename)
                ext = ext.lower()
                if ext == "" and len(basename) and basename[0] == ".":
                    continue
                if ext in INTERESTING_FILE_EXTENSIONS:
                    extracted.append(info.filename)

            options = Config()
            tmp_path = options.cuckoo.get("tmppath", "/tmp")
            target_path = os.path.join(tmp_path, "cuckoo-zip-tmp")
            if not os.path.exists(target_path):
                os.mkdir(target_path)
            tmp_dir = tempfile.mkdtemp(prefix='cuckoozip_',dir=target_path)

            for extfile in extracted:
                try:
                    retlist.append(archive.extract(extfile, path=tmp_dir, pwd=password))
                except:
                    retlist.append(archive.extract(extfile, path=tmp_dir))
    except:
        pass

    return retlist

def demux_rar(filename, options):
    retlist = []

    if not HAS_RARFILE or not filename.endswith(".rar"):
        return retlist

    try:
        # don't try to auto-extract RAR SFXes
        magic = File(filename).get_type()
        if "PE32" in magic or "MS-DOS executable" in magic:
            return retlist

        extracted = []
        password="infected"
        fields = options.split(",")
        for field in fields:
            try:
                key, value = field.split("=", 1)
                if key == "password":
                    password = value
                    break
            except:
                pass

        with RarFile(filename, "r") as archive:
            infolist = archive.infolist()
            for info in infolist:
                # avoid obvious bombs
                if info.file_size > 100 * 1024 * 1024 or not info.file_size:
                    continue
                # ignore directories
                if info.filename.endswith("\\"):
                    continue
                # add some more sanity checking since RarFile invokes an external handler
                if "..\\" in info.filename:
                    continue
                base, ext = os.path.splitext(info.filename)
                basename = os.path.basename(info.filename)
                ext = ext.lower()
                if ext == "" and len(basename) and basename[0] == ".":
                    continue
                if ext in INTERESTING_FILE_EXTENSIONS:
                    extracted.append(info.filename)
            options = Config()
            tmp_path = options.cuckoo.get("tmppath", "/tmp")
            target_path = os.path.join(tmp_path, "cuckoo-rar-tmp")
            if not os.path.exists(target_path):
                os.mkdir(target_path)
            tmp_dir = tempfile.mkdtemp(prefix='cuckoorar_',dir=target_path)

            for extfile in extracted:
                # RarFile differs from ZipFile in that extract() doesn't return the path of the extracted file
                # so we have to make it up ourselves
                try:
                    archive.extract(extfile, path=tmp_dir, pwd=password)
                    retlist.append(os.path.join(tmp_dir, extfile.replace("\\", "/")))
                except:
                    archive.extract(extfile, path=tmp_dir)
                    retlist.append(os.path.join(tmp_dir, extfile.replace("\\", "/")))
    except:
        pass

    return retlist

def demux_email(filename, options):
    retlist = []

    if not filename.endswith(".rar"):
        return retlist

    try:
        with open(filename, "rb") as openfile:
            buf = openfile.read()
            atts = find_attachments_in_email(buf, True)
            if atts and len(atts):
                for att in atts:
                    retlist.append(att[0])
    except:
        pass

    return retlist

def demux_msg(filename, options):
    retlist = []
    if not filename.endswith(".msg"):
        print "%s not endswith msg"%filename
        return None
    try:
        retlist = Message(filename).get_extracted_attachments()
    except:
        pass

    return retlist

def demux_sample(filename, package, options):
    """
    If file is a ZIP, extract its included files and return their file paths
    If file is an email, extracts its attachments and return their file paths (later we'll also extract URLs)
    """

    # if a package was specified, then don't do anything special
    # this will allow for the ZIP package to be used to analyze binaries with included DLL dependencies
    DEMUXERS = [demux_zip, demux_rar, demux_email, demux_msg] # demux functions are in charge of filtering files
    if package:
        return [ filename ]

    files = [{'filename': filename, 'options': options, 'origin': 'sample', 'has_children': False, 'processed': False}] # File submited
    tmp_files = []
    if not KEEP_INTERMEDIATE_FILES:
        for f in files:
            if f['has_children']:
                files.remove(f)


    filenames = [f['filename'] for f in files]
    print filenames
    return filenames

