"""
Changing subtitle sytle
"""
import sys
import fileinput
import regex

subtitle = sys.argv[1]
subfont = sys.argv[2]
sf = subfont + ","

# ScriptInfo section, font, outline(note:if borderstlye=1 then this ignored when seeing the subtitle), shadow(note:similar as outline)
replacements = [(r"(?:^Original Script:)\K(.*)", ""),
                (r"(?:^Original Translation:)\K(.*)", ""),
                (r"(?:^Original Editing:)\K(.*)", ""),
                (r"(?:^Original Timing:)\K(.*)", ""),
                (r"(?:^Script Updated By:)\K(.*)", ""),
                (r"(?:^Update Details:)\K(.*)", ""),
                (r"(?:^!|^;)\K(.*)", ""),
                (r"(?:^Style: .*?,)\K().*?,", sf),
                (r"(?:^Style: .*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,)\K(.*?,)", "1,"),
                (r"(?:^Style: .*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,.*?,)\K().*?,", "1,")]
for font in fileinput.input(subtitle, inplace=True, encoding="utf-8"):
    for pat, repl in replacements:
        font = regex.sub(pat, repl, font.strip())
    print(font.strip())
