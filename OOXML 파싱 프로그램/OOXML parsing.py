import struct
import zlib
import os

# little endian, 10진수로 변환
def little2(hex): 
    return struct.unpack('<H', hex)[0] # 2byte

zipfilename = 'ooxml/ppt1.zip'
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
name = []
data_offset = []
ooxml = 'No'

for i in range(len(LF_sig_offset)):
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
    name.append(name_hex.decode())

    # data
    dataOffset = name_offset + nameLen + extraLen
    data_offset.append(dataOffset)    
    
    if name_hex.decode() == '[Content_Types].xml':
        ooxml = 'Yes'
    
# data 길이 구하기
dataLen = []
for i in range(len(data_offset)):
    if i == len(data_offset)-1:
        dataLen.append(CF_sig_start_offset - data_offset[i])
    else:
        dataLen.append(LF_sig_offset[i+1] - data_offset[i])

print('File Name :', name)
print('\nData Offset :', data_offset)
print('\nData len :', dataLen)

print('\nOOXML인가? :', ooxml)

# app.xml, core.xml, image,png, slide.xml, sharedStrings.xml, document.xml의
# data offset, data len 구하기
media_name, media_data_offset, media_dataLen = [], [], []
slide_data_offset, slide_dataLen = [], []
for i in range(len(name)):
    if '/media/image' in name[i]:
        media_name.append(name[i])
        media_data_offset.append(data_offset[i])
        media_dataLen.append(dataLen[i])
    elif name[i] == 'docProps/app.xml':
        app_data_offset = data_offset[i]
        app_dataLen = dataLen[i]
    elif name[i] == 'docProps/core.xml':
        core_data_offset = data_offset[i]
        core_dataLen = dataLen[i]
    elif 'ppt/slides/slide' in name[i]:  
        slide_data_offset.append(data_offset[i])
        slide_dataLen.append(dataLen[i])
    elif name[i] == 'xl/sharedStrings.xml':
        sheet_data_offset = data_offset[i]
        sheet_dataLen = dataLen[i]
    elif name[i] == 'word/document.xml':
        pages_data_offset = data_offset[i]
        pages_dataLen = dataLen[i]

# media data
for i in range(len(media_name)):
    f.seek(media_data_offset[i])
    media_data = f.read(media_dataLen[i])
    name = media_name[i].split('/')[-1]
    zipname = zipfilename.split('/')[-1][:-4]
    os.makedirs('ooxml/media/'+zipname, exist_ok=True)
    fw = open('ooxml/media/'+zipname+'/'+name, 'wb+')
    fw.write(media_data)
    fw.close()

def xml(data, words):
    word = words.split()
    if len(word) == 1:
        word_s = word_e = word[0]
    else:
        word_s, word_e = words, word[0]
    
    word_result = []
    while(1):
        word_start = data.find("<"+word_s+">")
        word_end = data.find("</"+word_e+">", word_start)
       
        if word_start == -1 or word_end == -1:
            if len(word_result) == 0:
                word_result.append('')
            break
        else:
            wordsLen = len(words) + 2
            word = data[word_start+wordsLen:word_end]
            data = data[word_end+11:]
            if word == '': continue
            else:
                word_result.append(word)
    
    if len(word_result) == 1:
        return word_result[0]
    else:
        return word_result

def xmldata(dataoffset, datalen):
    f.seek(dataoffset)
    data = f.read(datalen)
    data = zlib.decompress(data, -zlib.MAX_WBITS)
    data = data.decode()
    return data

# app.xml data
app_data = xmldata(app_data_offset, app_dataLen)

ooxml_type = xml(app_data, "Application")
print('\nOOXML Type :', ooxml_type)
print('Version :', xml(app_data, "AppVersion"))

if ooxml_type == 'Microsoft Office Word':
    print('Pages 개수 :', xml(app_data, "Pages"))
elif ooxml_type == 'Microsoft Office PowerPoint':
    print('Slides 개수 :', xml(app_data, "Slides"))
    print('숨겨진 Slides 개수 :', xml(app_data, "HiddenSlides"))
elif ooxml_type == 'Microsoft Excel':
    TitlesOfParts = xml(app_data, "TitlesOfParts")
    TitlesOfParts_num = int(TitlesOfParts.split()[1][6:-1])
    print('Sheets 개수 :', TitlesOfParts_num)
    print('Sheets 타이틀 :', xml(TitlesOfParts, "vt:lpstr"))


# core.xml data
core_data = xmldata(core_data_offset, core_dataLen)

print('\n생성한 사람 :', xml(core_data, "dc:creator"))
print('수정한 사람 :', xml(core_data, "cp:lastModifiedBy"))
print('생성한 날짜 :', xml(core_data, 'dcterms:created xsi:type="dcterms:W3CDTF"'))
print('수정한 날짜 :', xml(core_data, 'dcterms:modified xsi:type="dcterms:W3CDTF"'))

# ppt slide figure, text data
if ooxml_type == 'Microsoft Office PowerPoint':
    for i in range(len(slide_data_offset)):
        slide_data = xmldata(slide_data_offset[i], slide_dataLen[i])
        
        # figure 구하기
        nvSpPr = xml(slide_data, "p:nvSpPr")
        fig, fig_dic = [], {}
        for j in range(len(nvSpPr)):
            cNvPr_e = nvSpPr[j].find(">")
            cNvPr = nvSpPr[j][0:cNvPr_e+1]
            fn_s = cNvPr.find("name=")
            figs = cNvPr[fn_s+6:-2]
            fig = figs.split()
            fig_name = ' '.join(fig[0:-1])
            if fig_name not in fig_dic:
                fig_dic[fig_name] = 1
            else:
                fig_dic[fig_name] += 1
        print('\nSlide'+str(i+1)+' Figure :', fig_dic)

        # text 구하기
        txBody = xml(slide_data, 'p:txBody')
        text = []
        for k in range(len(txBody)):
            at = xml(txBody[k], 'a:t')
            if isinstance(at, list) == True:
                at = ''.join(at)
            if at != '':
                text.append(at)
        print('Slide'+str(i+1)+' Text :', text)

# word pages text
elif ooxml_type == 'Microsoft Office Word':
    pages_data = xmldata(pages_data_offset, pages_dataLen)
    
    p_text = []
    while(1):
        p_s = pages_data.find('<w:p w14:paraId=')
        p_e = pages_data.find('>', p_s)
        if p_s == -1 or p_e == -1:
            break
        wp = xml(pages_data, pages_data[p_s+1:p_e])
        wt = xml(wp, 'w:t')
        wts = xml(wp, 'w:t xml:space="preserve"')
        p_t = ''.join(wt) + ''.join(wts)
        if p_t != '': p_text.append(p_t)
        pages_data = pages_data[p_e:]

    print('\nPages text :', p_text)

# excel sheet text
elif ooxml_type == 'Microsoft Excel':
    sheet_data = xmldata(sheet_data_offset, sheet_dataLen)
    
    print('\nSheet text :', xml(sheet_data, 't'))
