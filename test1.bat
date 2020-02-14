
@echo off

for /f "delims=" %%a in ('wmic OS Get localdatetime ^| find "."') do set dt=%%a
set year=%dt:~0,4%
set monthFirstFile=%dt:~4,2%
set monthSecondFile=%dt:~4,2%
set day=%dt:~6,2%
if %monthFirstFile%==01 (
	set monthFirstFile=Decembre 
	set monthSecondFile=Janvier
)
if %monthFirstFile%==02 (
	set monthFirstFile=Janvier
	set monthSecondFile=Fevrier
)
if %monthFirstFile%==03 (
	set monthFirstFile=Fevrier
	set monthSecondFile=Mars
)
if %monthFirstFile%==04 (
	set monthFirstFile=Mars
	set monthSecondFile=Avril
)
if %monthFirstFile%==05 (
	set monthFirstFile=Avril
	set monthSecondFile=Mai
)
if %monthFirstFile%==06 (
	set monthFirstFile=Mai
	set monthSecondFile=Juin
)
if %monthFirstFile%==07 (
	set monthFirstFile=Juin
	set monthSecondFile=Juillet
)
if %monthFirstFile%==08 (
	set monthFirstFile=Juillet
	set monthSecondFile=Aout
)
if %monthFirstFile%==09 (
	set monthFirstFile=Aout
	set monthSecondFile=Septembre
)
if %monthFirstFile%==10 (
	set monthFirstFile=Septembre
	set monthSecondFile=Octobre
)
if %monthFirstFile%==11 (
	set monthFirstFile=Octobre
	set monthSecondFile=Novembre
)
if %monthFirstFile%==12 (
	set monthFirstFile=Novembre
	set monthSecondFile=Decembre
)

set FileCut="L:\7-DINSI\Services en production\Relations B‚n‚ficiaires\_Centre de service\Pilotage SVP\suivi SLA DCS\stat grandlyon\2020\31 Stats %monthSecondFile%2020test.xlsm"

copy /y "L:\7-DINSI\Services en production\Relations B‚n‚ficiaires\_Centre de service\Pilotage SVP\suivi SLA DCS\stat grandlyon\2020\31 Stats %monthFirstFile%2020 - Copie.xlsm" %FileCut%

echo "Choisir le fichier de statistique qui vient d'être g,n,r,"
pause 


c:\windows\system32\cscript.exe "P:\automation process\test2.vbs"

cd "C:\Program Files (x86)\Alcatel\A4400 Call Center Supervisor\"
start ccs.exe

pause
echo "etape 2 test"