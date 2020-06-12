# excel2json

This simple tool converts XLS/XLSX files into regular JSON structure 
for simple way of parsing and so on.

Result file will be saved into the same path as the source and with the same filename with suffix '.json'
## build

`./gradlew shadowJar`

## run

```
java -jar xl2json.jar --filename|-f <filename.xlsx>
	Excel workbook xls/xlsx will be converted into <filename.xlsx.json>
options:
	--parse-hidden-cols|-c toggle parsing hidden columns
	--parse-hidden-rows|-r toggle parsing hidden rows
	--parse-empty-rows|-e toggle parsing empty rows (all the row's cells are empty)
	--help|-h prints this message
```

## requirements

*JDK 8*