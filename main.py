
#from zipfile import ZipFile as Zf
import datetime
import zipfile

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
    zf = zipfile.ZipFile(archive_name)
    return zf.read(filename)

def parse_customProp_file(xmlContent):
    pass
    
       
if __name__ == '__main__':
    #print_content_file_info('REVD-2019.docx')
    print (
        extract_zipped_content('REVD-2019.docx','docProps/custom.xml')
        )
