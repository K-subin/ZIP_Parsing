import struct
import zlib
import os

# little endian, 10진수로 변환
def little2(hex): 
    return struct.unpack('<H', hex)[0] # 2byte

zipfilename = 'hidden excel.zip'
f = open(zipfilename, 'rb+')

# local file signature offset 찾기
LF_sig = b'\x50\x4B\x03\x04'
CF_sig = b'\x50\x4B\x01\x02'
LF_sig_offset = []
offset = 0

while True:
    f.seek(offset)
    fr =  f.read(4)

    if fr == LF_sig:
        LF_sig_offset.append(offset)
    elif fr == CF_sig:
        CF_sig_start_offset = offset
        break
    offset += 1

# 파일 이름과 data offset 찾기
name, data_offset = [], []
hidden_name, hidden_dataoffset = [], []
sheet_cnt, slide_cnt = 0, 0

for i in range(len(LF_sig_offset)):
    # time
    time_offset = LF_sig_offset[i] + 10
    f.seek(time_offset)
    time_hex = f.read(2)
    
    # date
    date_offset = LF_sig_offset[i] + 12
    f.seek(date_offset)
    date_hex = f.read(2)
    
    # Name Length
    nameLen_offset = LF_sig_offset[i] + 26
    f.seek(nameLen_offset)
    nameLen_hex = f.read(2)
    nameLen = little2(nameLen_hex)

    # Extra Length
    extraLen_offset = nameLen_offset + 2
    f.seek(extraLen_offset)
    extraLen_hex = f.read(2)
    extraLen = little2(extraLen_hex)

    # Name
    name_offset = extraLen_offset + 2
    f.seek(name_offset)
    name_hex = f.read(nameLen)
    name_hex = name_hex.decode()
    name.append(name_hex)
    
    # data
    dataOffset = name_offset + nameLen + extraLen
    data_offset.append(dataOffset)    
    
    if time_hex != b'\x00\x00' or date_hex != b'\x21\x00':
        if name_hex[-1] != '/':
            hidden_name.append(name_hex)
            
    if 'xl/worksheets/sheet' in name_hex:
        sheet_cnt += 1 
    if 'ppt/slides/slide' in name_hex:
        slide_cnt += 1 


# data 길이 구하기
dataLen = []
for i in range(len(data_offset)):
    if i == len(data_offset)-1:
        dataLen.append(CF_sig_start_offset - data_offset[i])
    else:
        dataLen.append(LF_sig_offset[i+1] - data_offset[i])

print('File Name :', name)
print('hidden file using modification time :', hidden_name)

def xmldata(dataoffset, datalen):
    f.seek(dataoffset)
    data = f.read(datalen)
    data = zlib.decompress(data, -zlib.MAX_WBITS)
    return data

for i in range(len(name)):
    # external hidden file
    if name[i][-4:] == '.xml':
        xml_data = xmldata(data_offset[i], dataLen[i])
        if xml_data[:55] != b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>':
            print('external hidden file :', name[i])

    # internal sheet hidden 
    if name[i] == 'xl/workbook.xml':
        wb_data = xmldata(data_offset[i], dataLen[i])
        wb_data = wb_data.decode()
        ws = wb_data.find('<sheets>')
        we = wb_data.find('</sheets>', ws)
        sheets = wb_data[ws+8:we]
        seen_sheet_cnt = sheets.count('<sheet')

        if seen_sheet_cnt != sheet_cnt:
            print('internal sheet hidden', sheet_cnt - seen_sheet_cnt, '개')

    # internal slide hidden 
    if name[i] == 'ppt/presentation.xml':
        pt_data = xmldata(data_offset[i], dataLen[i])
        pt_data = pt_data.decode()
        ps = pt_data.find('<p:sldIdLst>')
        pe = pt_data.find('</p:sldIdLst>', ps)
        slides = pt_data[ps+12:pe]
        seen_slide_cnt = slides.count('<p:sldId')
        if seen_slide_cnt != slide_cnt:
            print('internal slide hidden', slide_cnt - seen_slide_cnt, '개')
