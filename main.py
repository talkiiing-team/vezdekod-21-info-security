import os
import xml.etree.ElementTree as ET
import random
import sys

def word(filename):
    os.system(f'unzip {filename}')

    new_application_name = str(random.randint(0, 100))
    new_creator = str(random.randint(0, 100))
    new_last_edit_by = str(random.randint(0, 100))
    new_create_time = str(random.randint(0, 100))
    new_change_time = str(random.randint(0, 100))

    tree = ET.parse('docProps/app.xml')
    root = tree.getroot()

    #Application Name
    app_application_offset = 5
    element = root[app_application_offset]
    application_old = element.text
    element.text = new_application_name
    tree.write('docProps/app.xml')
    print(f'changed Application Name from {application_old} to {element.text}')

    tree = ET.parse('docProps/core.xml')
    root = tree.getroot()

    #Creator
    core_creator_offset = 2
    element = root[core_creator_offset]
    creator_old = element.text
    element.text = new_creator
    print(f'changed Creator from {creator_old} to {element.text}')

    #Last Edited by
    core_last_edit_by = 5
    element = root[core_last_edit_by]
    last_edit_by_old = element.text
    element.text = new_last_edit_by
    tree.write('docProps/core.xml')
    print(f'changed Last Edited by from {last_edit_by_old} to {element.text}')

    #Create time
    core_last_edit_by = 7
    element = root[core_last_edit_by]
    create_time_old = element.text
    element.text = new_create_time
    tree.write('docProps/core.xml')
    print(f'changed Last Edited by from {create_time_old} to {element.text}')

    #Edition time
    core_last_edit_by = 8
    element = root[core_last_edit_by]
    change_time_old = element.text
    element.text = new_change_time
    tree.write('docProps/core.xml')
    print(f'changed Last Edited by from {change_time_old} to {element.text}')

    os.system('zip edited.docx \[Content_Types\].xml word/ _rels/ docProps/ -r')
    os.system('rm -rf _rels/ docProps/ word/ \[Content_Types\].xml')

def pdf(filename):
    os.system(f'cp {filename} edited{filename[-4:]}')

    os.system(f'exiftool -Keywords=aboba edited.pdf')
    os.system(f'exiftool -Notes=aboba edited.pdf')
    os.system(f'exiftool -ReleaseDate=aboba edited.pdf')
    os.system(f'exiftool -ReleaseTime=00:00:00 edited.pdf')
    os.system(f'exiftool -DateTime=aboba edited.pdf')
    os.system(f'exiftool -Author=aboba edited.pdf')
    os.system(f'exiftool -CreatorTool=aboba edited.pdf')
    os.system(f'exiftool -OwnerName=aboba edited.pdf')
    os.system(f'exiftool -ModificationDate=aboba edited.pdf')
    os.system(f'exiftool -Creator=aboba edited.pdf')
    os.system(f'exiftool -Title=aboba edited.pdf')
    os.system(f'exiftool -Producer=aboba edited.pdf')

if __name__ == '__main__':
    filename = sys.argv[1]

    if filename[-4:] == 'docx':
        word(filename)
    if filename[-4:] == '.pdf':
        pdf(filename)
