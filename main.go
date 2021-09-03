package main

import (
	"log"
	"os"
	"strconv"

	ical "github.com/arran4/golang-ical"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var (
	rootCmd = &cobra.Command{
		Use:   "icstoexcel",
		Short: "iCal to Excel converter",
		Run:   run,
	}
	input  string
	output string
)

func main() {
	rootCmd.PersistentFlags().StringVar(&input, "input", "", "Define the input file (*.ics)")
	rootCmd.PersistentFlags().StringVar(&output, "output", "", "Define the output file (*.xlsx)")

	if err := rootCmd.Execute(); err != nil {
		log.Fatalln(err.Error())
	}
}

func propval(event *ical.VEvent, prop ical.ComponentProperty) string {
	property := event.GetProperty(prop)
	if property == nil {
		return ""
	}
	return property.Value
}

func run(cmd *cobra.Command, args []string) {
	icalFile, err := os.Open(input)
	if err != nil {
		log.Fatalln(err.Error())
	}
	defer icalFile.Close()

	cal, err := ical.ParseCalendar(icalFile)
	if err != nil {
		log.Fatalln(err.Error())
	}

	ws := excelize.NewFile()
	sheet := "Sheet1"

	_ = ws.SetCellStr(sheet, "A1", "Summary")
	_ = ws.SetCellStr(sheet, "B1", "From")
	_ = ws.SetCellStr(sheet, "C1", "To")
	_ = ws.SetCellStr(sheet, "D1", "Location")

	for i, event := range cal.Events() {

		row := i + 2
		_ = ws.SetCellStr(sheet, "A"+strconv.Itoa(row), propval(event, ical.ComponentPropertySummary))
		_ = ws.SetCellStr(sheet, "B"+strconv.Itoa(row), propval(event, ical.ComponentPropertyDtStart))
		_ = ws.SetCellStr(sheet, "C"+strconv.Itoa(row), propval(event, ical.ComponentPropertyDtEnd))
		_ = ws.SetCellStr(sheet, "D"+strconv.Itoa(row), propval(event, ical.ComponentPropertyLocation))
	}

	if err := ws.SaveAs(output); err != nil {
		log.Fatalln(err)
	}
}
