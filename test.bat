
for /f "delims=" %%a in ('wmic OS Get localdatetime ^| find "."') do set dt=%%a
set year=%dt:~0,4%
set monthFirstFile=%dt:~4,2%
set monthSecondFile=%dt:~4,2%
set day=%dt:~6,2%
if %monthFirstFile%==01 (
	set monthFirstFile=janvier 
	set monthSecondFile=fevrier
)
if %monthFirstFile%==02 (
	set monthFirstFile=fevrier
	set monthSecondFile=mars
)
if %monthFirstFile%==03 (
	set monthFirstFile=mars
	set monthSecondFile=avril
)
if %monthFirstFile%==04 (
	set monthFirstFile=avril
	set monthSecondFile=mai
)
if %monthFirstFile%==05 (
	set monthFirstFile=mai
	set monthSecondFile=juin
)
if %monthFirstFile%==06 (
	set monthFirstFile=juin
	set monthSecondFile=juilllet
)
if %monthFirstFile%==07 (
	set monthFirstFile=juilllet
	set monthSecondFile=aout
)
if %monthFirstFile%==08 (
	set monthFirstFile=aout
	set monthSecondFile=septembre
)
if %monthFirstFile%==09 (
	set monthFirstFile=septembre
	set monthSecondFile=octobre
)
if %monthFirstFile%==10 (
	set monthFirstFile=octobre
	set monthSecondFile=novembre
)
if %monthFirstFile%==11 (
	set monthFirstFile=novembre
	set monthSecondFile=decembre
)
if %monthFirstFile%==12 (
	set monthFirstFile=decembre
	set monthSecondFile=janvier
)

set FileCut=C:\Users\EXTSLIT\Documents\theme%monthSecondFile%.css
copy /y C:\Users\EXTSLIT\Documents\theme%monthFirstFile%.css %FileCut%


c:\windows\system32\cscript.exe C:\Users\EXTSLIT\Documents\test2.vbs