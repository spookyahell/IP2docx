import zipfile
import sys
import urllib.request
import json

langs = {'de': 
		['IP Adresse',
			'Hostname',
			'Land',
			'Stadt',
			'Region',
			'Land',
			'GEO-Daten',
			'Internet Anbieter (ISP)',
			'Dokument wurde erstellt']}

if len(sys.argv) == 2:
	lang = sys.argv[1]
else:
	lang = 'de'
	
if lang not in langs.keys():
	print('This language is not avaialable (yet).')
	sys.exit(1)


with urllib.request.urlopen('http://ipinfo.io/json') as response:
	jsontext = response.read().decode('utf-8')

rj = json.loads(jsontext)

#~ exit()

ip = rj.get('ip')

hostname = rj.get('hostname')

city = rj.get('city')

region = rj.get('region')

country = rj.get('country')

loc = rj.get('loc')

org = rj.get('org')


#~ print(ip)
#~ exit()

file = open('document.xml','w')

file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
file.write('<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">\n')
file.write('    <w:body>\n')
file.write('        <w:p w:rsidR="00372BC2" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
file.write('            <w:pPr>\n')
file.write('                <w:pStyle w:val="Titel"/>\n')
file.write('                <w:jc w:val="center"/>\n')
file.write('            </w:pPr>\n')
file.write('            <w:r>\n')
file.write('                <w:t>IP INFO</w:t>\n')
file.write('            </w:r>\n')
file.write('        </w:p>\n')

file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
file.write('            <w:pPr>\n')
file.write('                <w:pStyle w:val="berschrift1"/>\n')
file.write('            </w:pPr>\n')
file.write('            <w:r>\n')
file.write('                <w:t>IP-Adresse:</w:t>\n')
file.write('            </w:r>\n')
file.write('        </w:p>\n')
file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
file.write('            <w:r w:rsidRPr="001A1564">\n')
file.write(f'                <w:t>{ip}</w:t>\n')
file.write('            </w:r>\n')
file.write('        </w:p>\n')
if hostname != None:
	file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
	file.write('            <w:pPr>\n')
	file.write('                <w:pStyle w:val="berschrift1"/>\n')
	file.write('                <w:rPr>\n')
	file.write('                    <w:lang w:val="en-US"/>\n')
	file.write('                </w:rPr>\n')
	file.write('            </w:pPr>\n')
	file.write('            <w:r>\n')
	file.write('                <w:rPr>\n')
	file.write('                    <w:lang w:val="en-US"/>\n')
	file.write('                </w:rPr>\n')
	file.write('                <w:t>Hostname:</w:t>\n')
	file.write('            </w:r>\n')
	file.write('        </w:p>\n')
	file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
	file.write('            <w:pPr>\n')
	file.write('                <w:rPr>\n')
	file.write('                    <w:lang w:val="en-US"/>\n')
	file.write('                </w:rPr>\n')
	file.write('            </w:pPr>\n')
	file.write('            <w:proofErr w:type="gramStart"/>\n')
	file.write('            <w:r w:rsidRPr="001A1564">\n')
	file.write('                <w:rPr>\n')
	file.write('                    <w:lang w:val="en-US"/>\n')
	file.write('                </w:rPr>\n')
	file.write(f'                <w:t>{hostname}</w:t>\n')
	file.write('            </w:r>\n')
	file.write('            <w:proofErr w:type="gramEnd"/>\n')
	file.write('        </w:p>\n')

if city != '':
	file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
	file.write('            <w:pPr>\n')
	file.write('                <w:pStyle w:val="berschrift1"/>\n')
	file.write('                <w:rPr>\n')
	file.write('                    <w:lang w:val="en-US"/>\n')
	file.write('                </w:rPr>\n')
	file.write('            </w:pPr>\n')
	file.write('            <w:proofErr w:type="spellStart"/>\n')
	file.write('            <w:r>\n')
	file.write('                <w:rPr>\n')
	file.write('                    <w:lang w:val="en-US"/>\n')
	file.write('                </w:rPr>\n')
	file.write('                <w:t>Stadt</w:t>\n')
	file.write('            </w:r>\n')
	file.write('            <w:proofErr w:type="spellEnd"/>\n')
	file.write('            <w:r>\n')
	file.write('                <w:rPr>\n')
	file.write('                    <w:lang w:val="en-US"/>\n')
	file.write('                </w:rPr>\n')
	file.write('                <w:t>:</w:t>\n')
	file.write('            </w:r>\n')
	file.write('        </w:p>\n')
	file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
	file.write('            <w:pPr>\n')
	file.write('                <w:rPr>\n')
	file.write('                    <w:lang w:val="en-US"/>\n')
	file.write('                </w:rPr>\n')
	file.write('            </w:pPr>\n')
	file.write('            <w:r>\n')
	file.write('                <w:rPr>\n')
	file.write('                    <w:lang w:val="en-US"/>\n')
	file.write('                </w:rPr>\n')
	file.write(f'                <w:t>{city}</w:t>\n')
	file.write('            </w:r>\n')
	file.write('        </w:p>\n')
	
file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
file.write('            <w:pPr>\n')
file.write('                <w:pStyle w:val="berschrift1"/>\n')
file.write('                <w:rPr>\n')
file.write('                    <w:lang w:val="en-US"/>\n')
file.write('                </w:rPr>\n')
file.write('            </w:pPr>\n')
file.write('            <w:r>\n')
file.write('                <w:rPr>\n')
file.write('                    <w:lang w:val="en-US"/>\n')
file.write('                </w:rPr>\n')
file.write('                <w:t>Land:</w:t>\n')
file.write('            </w:r>\n')
file.write('        </w:p>\n')
file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
file.write('            <w:pPr>\n')
file.write('                <w:rPr>\n')
file.write('                    <w:lang w:val="en-US"/>\n')
file.write('                </w:rPr>\n')
file.write('            </w:pPr>\n')
file.write('            <w:r>\n')
file.write('                <w:rPr>\n')
file.write('                    <w:lang w:val="en-US"/>\n')
file.write('                </w:rPr>\n')
#~ COUNTRY!
file.write(f'                <w:t>{country}</w:t>\n')
file.write('            </w:r>\n')
file.write('        </w:p>\n')
file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
file.write('            <w:pPr>\n')
file.write('                <w:pStyle w:val="berschrift1"/>\n')
file.write('                <w:rPr>\n')
file.write('                    <w:lang w:val="en-US"/>\n')
file.write('                </w:rPr>\n')
file.write('            </w:pPr>\n')
file.write('            <w:r>\n')
file.write('                <w:rPr>\n')
file.write('                    <w:lang w:val="en-US"/>\n')
file.write('                </w:rPr>\n')
file.write('                <w:t>GEO-Daten:</w:t>\n')
file.write('            </w:r>\n')
file.write('        </w:p>\n')

lspl = loc.split(',')
file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
file.write('            <w:pPr>\n')
file.write('                <w:rPr>\n')
file.write('                    <w:lang w:val="en-US"/>\n')
file.write('                </w:rPr>\n')
file.write('            </w:pPr>\n')
file.write('            <w:r w:rsidRPr="001A1564">\n')
file.write('                <w:rPr>\n')
file.write('                    <w:lang w:val="en-US"/>\n')
file.write('                </w:rPr>\n')
#~ GEODATA lat!
file.write(f'                <w:t>{lspl[0]}</w:t>\n')
file.write('            </w:r>\n')
file.write('            <w:proofErr w:type="gramStart"/>\n')
file.write('            <w:r w:rsidRPr="001A1564">\n')
file.write('                <w:rPr>\n')
file.write('                    <w:lang w:val="en-US"/>\n')
file.write('                </w:rPr>\n')
#~ GEODATA long!
file.write(f'                <w:t>, {lspl[1]}</w:t>\n')
file.write('            </w:r>\n')
file.write('            <w:proofErr w:type="gramEnd"/>\n')
file.write('        </w:p>\n')
file.write('        <w:p w:rsidR="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
file.write('            <w:pPr>\n')
file.write('                <w:pStyle w:val="berschrift1"/>\n')
file.write('            </w:pPr>\n')
file.write('            <w:r>\n')
file.write('                <w:t>Internet Anbieter:</w:t>\n')
file.write('            </w:r>\n')
file.write('        </w:p>\n')
file.write('        <w:p w:rsidR="001A1564" w:rsidRPr="001A1564" w:rsidRDefault="001A1564" w:rsidP="001A1564">\n')
file.write('            <w:r w:rsidRPr="001A1564">\n')
#~ PROVIDER!
file.write(f'                <w:t>{org}</w:t>\n')
file.write('            </w:r>\n')
file.write('            <w:bookmarkStart w:id="0" w:name="_GoBack"/>\n')
file.write('            <w:bookmarkEnd w:id="0"/>\n')
file.write('        </w:p>\n')
file.write('        <w:sectPr w:rsidR="001A1564" w:rsidRPr="001A1564">\n')
file.write('            <w:pgSz w:w="11906" w:h="16838"/>\n')
file.write('            <w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>\n')
file.write('            <w:cols w:space="708"/>\n')
file.write('            <w:docGrid w:linePitch="360"/>\n')
file.write('        </w:sectPr>\n')
file.write('    </w:body>\n')
file.write('</w:document>')

file.close()

zf = zipfile.ZipFile('IPdoc.docx','w')


writefs = ['[Content_Types].xml',
		r'_rels\.rels', 
		r'docProps\app.xml',
		r'docProps\core.xml',
		r'word\_rels\document.xml.rels',
		r'word\theme\theme1.xml',
		r'word\fontTable.xml',
		r'word\settings.xml',
		r'word\styles.xml',
		r'word\stylesWithEffects.xml',
		r'word\webSettings.xml']

for item in writefs:
	#~ print(f'Writing {item!r}') 
	zf.write(item, item)
	
zf.write('document.xml',r'word\document.xml')

zf.close()

print(langs[lang][8])