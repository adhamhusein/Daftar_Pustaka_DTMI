from xml.etree import ElementTree as ET
from docx import Document
import itertools
import threading
import time
import sys

#%%
print('- This Code Written by: Adh(17)')
print('- Redistribution and use in source and binary forms, with modification, are permitted.')
print('- If you find any bug please report to didikankakseto1@gmail.com\n\n')
print('Preparing..')

done = False
#here is the animation
def animate():
    for c in itertools.cycle(['|', '/', '-', '\\']):
        if done:
            break
        sys.stdout.write('\rloading ' + c)
        sys.stdout.flush()
        time.sleep(0.1)
    sys.stdout.write('')

t = threading.Thread(target=animate)
t.start()


time.sleep(5)
done = True

#%%
print('\n')
fname = str(input('Type file name:   '))

doc = Document()
daftar_pustaka = list()
try:
    dom = ET.parse(fname)
except:
    print('\n')
    print('== You do not input correct file name ==')
    print('== Please restart this program ==')
    sys.exit()

record = dom.findall('records/record')
for x in record:
    authors = x.find('contributors/authors')
    if authors is None: continue
    i = 0
    nama_format = None
    for y in authors:
        i+= 1
        nama_default = y.text
        kata = nama_default.split(', ')
        if len(kata) == 1:
            nama = kata[0] + ','
            nama_lgkp = nama
            if nama_format == None:
                nama_format = nama_lgkp
            else:
                if i == len(authors):
                    nama_format = nama_format + ' dan ' + nama_lgkp
                else:
                    nama_format = nama_format + ' ' + nama_lgkp
            continue
        a = kata[1].split()
        nama = kata[0]+', '
        nama_blkg = ''
        for b in a:
            c = b.split('.')
            try:
                c.remove('')
            except:
                pass
            j = 0
            for d in c:
                j+=1
                if nama_blkg == '' and len(a) == 1:
                    nama_blkg = str(d[0]) + '.,'
                elif nama_blkg == '':
                    nama_blkg = str(d[0]) + '. '
                else:
                    nama_blkg = nama_blkg + str(d[0]) + '. '
            
        if nama_blkg[len(nama_blkg)-1] != ',':
            string_list = list(nama_blkg)
            string_list[len(string_list)-1] = ','
            nama_blkg = ''.join(string_list)
                
        nama_lgkp = nama + nama_blkg
        if nama_format == None:
            nama_format = nama_lgkp
        else:
            if i == len(authors):
                nama_format = nama_format + ' dan ' + nama_lgkp
            else:
                nama_format = nama_format + ' ' + nama_lgkp
#%% 
# tahun OK

               
    year_xml = x.find('dates/year')
    try:
        year = year_xml.text
    except:
        year = 'Tanpa Tahun'
        
    try:
        nama_format = nama_format + ' ' + year + ', '
    except:
        nama_format = 'Tanpa Nama,' + ' ' + year + ', '
#%%
# Judul Jurnal OK


    judul_xml = x.find('titles/title')
    try:
        judul = judul_xml.text
    except:
        judul = 'Tanpa Judul'
    nama_format = nama_format + judul + ', '
#%%
# nama jurnal OK
    
    nama_jurnal_xml = x.find('periodical/full-title')
    try:
        nama_jurnal = nama_jurnal_xml.text
    except:
        nama_jurnal = 'Tanpa Nama Jurnal'   
        
    nama_format = nama_format + nama_jurnal + ', '
#%%
# jilid
       
    jilid_jurnal_xml = x.find('volume')
    try:
        jilid_jurnal = jilid_jurnal_xml.text
        nama_format = nama_format + 'Vol.' + jilid_jurnal + ', '
    except:
        nama_format = nama_format
#%%
# nama penerbit

    nama_penerbit_jurnal_xml = x.find('publisher')
    try:
        nama_penerbit_jurnal = nama_penerbit_jurnal_xml.text
        nama_penerbit_jurnal.replace('"','')
        nama_penerbit_jurnal.replace("'","")
        nama_format = nama_format + nama_penerbit_jurnal + ', '
    except:
        nama_format = nama_format
#%%
# Tempat terbit

    tempat_jurnal_xml = x.find('pub-location')
    try:
        tempat_jurnal = tempat_jurnal_xml.text
        nama_format = nama_format + tempat_jurnal + ', '
    except:
        nama_format = nama_format
#%% 


    nama_format = nama_format[:len(nama_format)-2]
    nama_format = nama_format + '.'
    daftar_pustaka.append(nama_format)
#%%

daftar_pustaka.sort()
numb = 0
for daftar in daftar_pustaka:
    doc.add_paragraph(daftar)
    numb+= 1
    print('Added Paragraph ' + str(numb) + '/' + str(len(daftar_pustaka)))


doc.save('Dafpus.docx')

print('\n\n== Daftar Pustaka Created ==')