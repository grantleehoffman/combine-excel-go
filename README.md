# combine-excel-go

# Build
Install Deps
`go mod tidy`

Windows 64
`GOOS=windows GOARCH=amd64 go build -o combine-excel.exe`

Windows 32
`GOOS=windows GOARCH=386 go build -o combine-excel.exe`

# Windows Run
`combine-excel.exe -i "C:\path\to\inputDir" -o "C:\path\to\outputFile.xlsx" -k "Value1,Value2,Value3" -r 5`

## Options
- i: The input file directory where xlsx files are located
- o: The output file path to write to
- k: Comma delimited list of keywords columns to copy over
- r: The row where the keywords are searched for