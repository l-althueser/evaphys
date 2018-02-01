# Split and convert HISLSF XML export for EvaSys

This repository contains several python3 scripts to convert or split XML files exported from the WWU Münster HISLSF system. 

## Convert XML to HTML list
```
python EvaSysXML.py --convert-to-html -i HISLSF-Export.xml
```

## Split and convert XML by event number
```
python EvaSysXML.py --convert-to-html --split-by-ID -i HISLSF-Export.xml
```
For a single department:
```
python EvaSysXML.py --convert-to-html --split-by-ID -k 11 -i HISLSF-Export.xml
```

## Split and convert XML by organization
```
python EvaSysXML.py --convert-to-html --split-by-ORG -i HISLSF-Export.xml
```
For a single department:
```
python EvaSysXML.py --convert-to-html --split-by-ORG -k "Institut für Musikpädagogik" -i HISLSF-Export.xml
```