rem example de fichier Make pour Gate, qui utilise une feuille Excel pour les données et une feuille Excel séparée pour les styles.

cd r:\01-Commun\bin
r:

set ROOTDIR=R:\EXAMPLEDIR
set WORKDIR=%ROOTDIR%\doc
set STYLEDIR=%ROOTDIR%\styles

perl gate.pl "%WORKDIR%\Data.xlsx" --style "%STYLEDIR%\StyleSheet.xlsx;MyStyles" > "%WORKDIR%\out.html"

rem Utilise les données de R:\EXAMPLEDIR\doc\Data.xlsx, le style de l'onglet "MyStyles" de la R:\EXAMPLEDIR\styles\StyleSheet.xlsx et met le résultat dans R:\EXAMPLEDIR\doc\out.html


pause
