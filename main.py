
#from zipfile import ZipFile as Zf
import datetime
import zipfile
import re
from xml.etree.ElementTree import ElementTree as ET, Element, SubElement, dump
from xml.etree.ElementTree import fromstring as ETfromstring, tostring as ETtostring

def print_content_file_info(archive_name):
    zf = zipfile.ZipFile(archive_name)
    for info in zf.infolist():
        print (info.filename)
        print ('\tComment:\t', info.comment)
        print ('\tModified:\t', datetime.datetime(*info.date_time))
        print ('\tSystem:\t\t', info.create_system, '(0 = Windows, 3 = Unix)')
        print ('\tZIP version:\t', info.create_version)
        print ('\tCompressed:\t', info.compress_size, 'bytes')
        print ('\tUncompressed:\t', info.file_size, 'bytes')
        print ()

def extract_zipped_content(archive_name, filename):
    """ to document """
    zf = zipfile.ZipFile(archive_name)
    return zf.read(filename)

def parse_customProp_file(xmlContent):
    """ to document """
    tree = ETfromstring(xmlContent)
    childLst = tree.getchildren()
    var = {}
    for child in childLst:
        var.update({child.get('name') : (child.getchildren())[0].text})
    # navigation sample in tree
    """
    child = childLst[1]
    print( child.get('name') )
    print( (child.getchildren())[0].text)
    """
    return var
       
if __name__ == '__main__':
    #print_content_file_info('REVD-2019.docx')
    print ( 'exctact in rawD variable')
    rawD = extract_zipped_content('REVD-2019.docx','docProps/custom.xml')
    print (
        parse_customProp_file(rawD)
        )
