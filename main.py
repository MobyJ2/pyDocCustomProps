
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
    childLst = list(tree)
    var = {}
    for child in childLst:
        var.update({child.get('name') : list(child)[0].text})
    # navigation sample in tree
    """
    child = childLst[1]
    print( child.get('name') )
    print( (child.getchildren())[0].text)
    """
    return var

# Try to re-construct the customProps.XML file
def data2XMLelementTree( headersDatas):
    ###'<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>'
    root = Element("Properties")
    tree = ET(root)
    for p in headersDatas:
        elt = SubElement(root, p)
        val = headersDatas[f"{p}"]
        elt.text = val
        #print( f"p= {p} and val={val}")
    return tree

def flushXmlFile(tree):
    fichier_xml = open("test.xml","w")
    tree.write(fichier_xml, encoding='unicode', xml_declaration='UTF-8')
    fichier_xml.close()
    return



if __name__ == '__main__':
    #print_content_file_info('REVD-2019.docx')
    print ( 'exctact in rawD variable')
    rawD = extract_zipped_content('REVD-2019.docx','docProps/custom.xml')
    myProps = parse_customProp_file(rawD)
    var = data2XMLelementTree(myProps)
    #print ( myProps)
    flushXmlFile( var )
    #print (var)
    tree = ETfromstring(rawD) #for test on tree.

    Hdr_VTypes = '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}'
    Hdr_Custom = '{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}'
    for elt in tree.iter(f"{Hdr_Custom}property"):
	#print( Element.attrib Element.tag Element.text)
        print(
            #f"elt.tag :{elt.tag}\n"
            f"   ..name : {elt.attrib['name']}\n"
            f"   ..val : {elt[0].text}\n"
            #f"   .text : {elt.text}\n"
	    #f"   .att : {elt.attrib}\n"
            )
        break

