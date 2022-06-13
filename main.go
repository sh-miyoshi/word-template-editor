package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"
	"regexp"

	"github.com/nguyenthenguyen/docx"
)

func main() {
	var inputFile, valueFile, outputFile string

	flag.StringVar(&inputFile, "input", "input.docx", "input docx file name")
	flag.StringVar(&valueFile, "value", "value.csv", "template values csv file name")
	flag.StringVar(&outputFile, "output", "output.docx", "output docx file name")
	flag.Parse()

	// Read value file
	fp, err := os.Open(valueFile)
	if err != nil {
		fmt.Printf("Failed to read value file: %+v\n", err)
		return
	}
	values := make(map[string]string)
	csvData := csv.NewReader(fp)
	for {
		record, err := csvData.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			fmt.Printf("Failed to read csv file data: %+v\n", err)
			return
		}

		if len(record) < 2 {
			fmt.Printf("csv data requires key and value, but got %s\n", record)
			return
		}

		values[record[0]] = record[1]
	}

	// Read docx file
	r, err := docx.ReadDocxFile(inputFile)
	if err != nil {
		fmt.Printf("Failed to read input file: %+v\n", err)
		return
	}

	docx := r.Editable()

	// Replace template values
	content := docx.GetContent()
	for key, value := range values {
		str := fmt.Sprintf(`{{\s*%s\s*}}`, key)
		re := regexp.MustCompile(str)
		content = re.ReplaceAllString(content, value)
	}
	docx.SetContent(content)

	docx.WriteToFile(outputFile)
}
