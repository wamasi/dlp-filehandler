
# $path = 'C:\video.wmv'
# $shell = New-Object -ComObject Shell.Application
# $folder = Split-Path $path
# $file = Split-Path $path -Leaf
#$shellfolder = $shell.Namespace($folder)
#$shellfile = $shellfolder.ParseName($file)
0..287 | ForEach-Object { '{0} = {1}' -f $_, $shellfolder.GetDetailsOf($null, $_) }

## =SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(REPLACE(REPLACE(F3, 1, IFERROR(FIND("//", F3)+1, 0), TEXT(,))&"/", FIND("/", REPLACE(F3, 1, IFERROR(FIND("//", F3)+1, 0), TEXT(,))&"/"), LEN(F3), TEXT(,)), CHAR(46), REPT(CHAR(32), LEN(F3))), LEN(F3)*2)), CHAR(32), CHAR(46))
