@ECHO OFF
TASKKILL /IM WeekNumberLite2+.exe /F >NUL 2>&1 
"C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x86\rc.exe" res.rc
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe" /nologo /target:winexe /platform:x86 /optimize+ /debug- /filealign:512 /win32res:res.res /main:P /out:WeekNumberLite2+.exe p.cs
START "" WeekNumberLite2+.exe
