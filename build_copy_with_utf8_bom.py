import codecs
import sys
fname_from = sys.argv[1]
fname_to = sys.argv[2]
template = None
with open(fname_from, "r",encoding='utf-8') as outfile: template = outfile.read()
with open(fname_to, "wb") as outfile:
    outfile.write(codecs.BOM_UTF8)
with open(fname_to, "a",encoding='utf-8') as outfile: outfile.write(template)
