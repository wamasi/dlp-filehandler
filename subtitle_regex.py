"""
Changing subtitle sytle
"""
import sys
import fileinput
import time
import regex

subtitle = sys.argv[1]
subfont = sys.argv[2]
sf = subfont + ","

# setting font
for font in fileinput.input(subtitle, inplace=True, encoding="utf-8"):
    font = regex.sub(r"(?:^Style: .*?,)\K().*?,", sf, font.strip())
    print(font.strip())
time.sleep(1)

# setting text outline to 1
for outline in fileinput.input(subtitle, inplace=True, encoding="utf-8"):
    outline = regex.sub(r"(?:^Style: .*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,)\K(.*?,.*?,)", "1,1,", outline.strip())
    print(outline.strip())
time.sleep(1)

# setting text shadow to 0
for shadow in fileinput.input(subtitle, inplace=True, encoding="utf-8"):
    # setting text shadow to 0
    shadow = regex.sub(r"(?:^Style: .*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,)\K().*?,", "0,", shadow.strip())
    print(shadow.strip())
time.sleep(1)
