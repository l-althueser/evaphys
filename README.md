# Split and convert HISLSF XML export for EvaSys

This repository contains several python3 scripts to convert or split XML files exported from the WWU Münster HISLSF system.

## Get newest XML data
```
python EvaSysXML.py --download-semester 20181
```
The `-i *.xml` parameter can be replaced by `--download-semester 20181` in the following!

## Convert XML to html, csv or excel
```
python EvaSysXML.py --convert-to html -i HISLSF-Export.xml
python EvaSysXML.py --convert-to csv -i HISLSF-Export.xml
python EvaSysXML.py --convert-to excel -i HISLSF-Export.xml
```

## Filter by event type
```
python EvaSysXML.py --include-type Vorlesung,V/Ü,"Grundkurs Vorlesung",Standardvorlesung,Vorlesung/Praktikum,Ringvorlesung,Vorlesung/Seminar -i HISLSF-Export.xml
```

## Split and convert XML by event number or organization
```
python EvaSysXML.py --convert-to html --split-by ID  -i HISLSF-Export.xml
python EvaSysXML.py --convert-to html --split-by ORG -i HISLSF-Export.xml
```
For a single department or organization:
```
python EvaSysXML.py --convert-to html --ID 11 -i HISLSF-Export.xml
python EvaSysXML.py --convert-to html --ORG "Institut für Musikpädagogik" -i HISLSF-Export.xml
```

## For FB11
```
python EvaSysXML.py --download-semester 20181 --convert-to excel --remove-duplicates --ID 11 --exclude-type Kolloquium,Forschungsseminar,"Anleitung zum wissenschaftlichen Arbeiten",Exkursion,Vorkurs,Tutorium,Übung,Workshop,Alle,Einführung,Reservierung
```
