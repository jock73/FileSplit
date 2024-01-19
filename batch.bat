@echo off
c:\windows\syswow64\wscript.exe  Splitfile.vbs
move "BrandA\LRJAGUS1\BrandA\*.csv" ..\2007642\files
move "Brandb\J2107\BrandB\*.csv" ..\2011860\files
move "Brandb\J2109\BrandB\*.csv" ..\2011861\files
move "Brandb\J3204\BrandB\*.csv" ..\2011862\files
move "Brandb\J3305\BrandB\*.csv" ..\2011863\files
move "Brandb\J8844\BrandB\*.csv" ..\2011864\files
move "Brandb\J4389\BrandB\*.csv" ..\2011865\files
move "Brandb\J4512\BrandB\*.csv" ..\2011866\files
move "Brandb\J4486\BrandB\*.csv" ..\2011867\files
move "Brandb\J4510\BrandB\*.csv" ..\2011868\files
move "Brandb\J4495\BrandB\*.csv" ..\2011869\files
move "Brandb\J4497\BrandB\*.csv" ..\2011870\files
move "Brandb\J4498\BrandB\*.csv" ..\2011871\files
move "Brandb\J8843\BrandB\*.csv" ..\2011872\files
move "Brandb\J4502\BrandB\*.csv" ..\2011873\files
move "Brandb\J4503\BrandB\*.csv" ..\2011874\files
move "Brandb\J4505\BrandB\*.csv" ..\2011875\files

RMDIR /S/Q BrandA\
RMDIR /S/Q Brandb\
exit
