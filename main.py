
#from zipfile import ZipFile as Zf
import datetime
import zipfile
import re
from xml.etree.ElementTree import ElementTree as ET, Element, SubElement, dump
from xml.etree.ElementTree import fromstring as ETfromstring, tostring as ETtostring


xmlns = {'Cust':'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
         'Vt'  :'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'}
    
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

def construct_ET_from_string(xmlFlux):
    RootElt = ETfromstring(xmlFlux)
    tree = ET(RootElt)
    return tree

def test_CustomProps( RootElt ):
    ###'<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>'
    for elt in RootElt.findall("Cust:property[@name]", xmlns):
            print(         f"   prop : {elt.attrib['name']}")
            for chld in elt.findall("Vt:lpwstr", xmlns):
                    print( f"   ---->:{chld.text}")
            print("")
            break #break after the first iteration
    return

def flushXmlFile(tree):
    fichier_xml = open("test.xml","w")
    tree.write(fichier_xml, encoding='unicode', xml_declaration='UTF-8')
    fichier_xml.close()
    return



if __name__ == '__main__':
    #print_content_file_info('REVD-2019.docx')
    print ( 'exctact in rawD variable')
    rawD = extract_zipped_content('REVD-2019.docx','docProps/custom.xml')
    tree = construct_ET_from_string(rawD)
    test_CustomProps(tree)
    flushXmlFile( tree )


    print( tree.find("Cust:property", xmlns ).attrib['name'] )
    print( tree.find("Cust:property", xmlns ).find("Vt:lpwstr", xmlns).text )

    print( tree.find("Cust:property[@name='ref-modele']", xmlns) )
    print( tree.find("Cust:property[@name='ref-modele']", xmlns).find("Vt:lpwstr", xmlns).text )

    propName = 'ref-modele'
    print( tree.find(f"Cust:property[@name='{propName}']", xmlns).find("Vt:lpwstr", xmlns).text )

    
