# excel2json

This simple tool converts XLS/XLSX files into regular JSON structure 
for simple way of parsing and so on.

Result file will be saved into the same path as the source and with the same filename with suffix '.json'
## build

`./gradlew shadowJar`

## run

`java -jar build/libs/xl2json.jar source_file.xlsx`