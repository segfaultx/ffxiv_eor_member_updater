package main

import (
	"fmt"
	"github.com/xivapi/godestone/v2"
	"github.com/xuri/excelize/v2"
	"log"
)

var (
	ClassColumnIndices = make(map[string]int)
)

func ProcessExcel(filename string) {
	f, err := excelize.OpenFile(filename)

	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	activeSheetName := f.GetSheetName(f.GetActiveSheetIndex())
	rows, err := f.GetRows(activeSheetName)

	if err != nil {
		log.Fatalln(err)
	}

	if len(rows) == 0 {
		log.Fatalln("Empty spreadsheet found")
	}
	generateClassRangeAndGetLastIndex(rows[0])

	for index, row := range rows[1:] {
		characterId, err := GetCharacterId(fmt.Sprintf("%s %s", row[0], row[1]))

		if err != nil {
			log.Fatalln(err)
		}

		characterInfo := GetCharacterInfo(characterId)
		updateCharacterData(f, characterInfo, index+2, activeSheetName) // add 2 to start at row 2, skipping the header row
	}

	err = f.Save()

	if err != nil {
		log.Fatalln(err)
	}
}

func updateCharacterData(sheet *excelize.File, characterData *godestone.Character, row int, sheetName string) {
	log.Printf("Updating Character Data for: %s\n", characterData.Name)

	for _, job := range characterData.ClassJobs {
		var mappedJob string
		if value, ok := Sub30MappingMap[job.UnlockedState.Name]; ok {
			mappedJob = value
		} else {
			mappedJob = job.UnlockedState.Name
		}

		columnId, err := excelize.ColumnNumberToName(ClassColumnIndices[mappedJob])

		if err != nil {
			log.Fatalln(err)
		}
		column := fmt.Sprintf("%s%d", columnId, row)

		if DEBUG {
			fmt.Printf("Job: %s, Level: %d, Column: %s\n", mappedJob, job.Level, column)
		}
		err = sheet.SetCellValue(sheetName, column, job.Level)

		if err != nil {
			log.Fatalln(err)
		}
	}
}

func generateClassRangeAndGetLastIndex(data []string) {
	startSet := false

	for index, value := range data {
		if _, ok := GermanToEnglishClassMap[value]; ok && !startSet {
			startSet = true
		}
		val, ok := GermanToEnglishClassMap[value]

		if !ok && startSet {
			log.Fatalf("Unknown class: %s, valid classes: %s\n", value, GermanToEnglishClassMap)
		}
		if ok && startSet {
			ClassColumnIndices[val] = index + 1
		}
	}

	if DEBUG {
		log.Printf("Class Columns: %s\n", ClassColumnIndices)
	}
}
