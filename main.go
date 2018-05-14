package main

import (
	"fmt"
	"path/filepath"
	"strings"
	"io/ioutil"
	"os"
	"bytes"
	"baliance.com/gooxml/document"
	"log"
)

func main()  {
	inputFilePath := filepath.Join("./test_input")
	outputFilePath := filepath.Join("./test_output")

	err := filepath.Walk(inputFilePath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if info.IsDir() {
			return nil
		}

		f, err := os.Open(path)
		if err != nil {
			return err
		}
		defer f.Close()

		doc, err := document.Read(f, info.Size())
		if err != nil {
			return err
		}

		var buf bytes.Buffer

		// All Paragraphs
		for _, par := range doc.Paragraphs() {

			for _, run := range par.Runs() {
				txt := run.Text()

				if run.Properties().Color().X().ValAttr.String() != "" {
					txt = fmt.Sprintf("[%s](%s)", txt, run.Properties().Color().X().ValAttr.String())
				}

				_, err = buf.WriteString(txt)
				if err != nil {
					return err
				}
			}
			_, err = buf.WriteString("\n")
			if err != nil {
				return err
			}
		}

		// Tables
		for _, tbl := range doc.Tables() {
			for _, row := range tbl.Rows() {
				for _, cell := range row.Cells() {
					for _, par := range cell.Paragraphs() {
						for _, run := range par.Runs() {
							for _, v := range run.DrawingAnchored() {
								fmt.Printf("HAS DR%v \n", v)
							}
							_, err = buf.WriteString(run.Text())
							if err != nil {
								return err
							}
						}
						_, err = buf.WriteString("\n")
						if err != nil {
							return err
						}
					}
				}
			}
		}

		// Tags
		for _, tag := range doc.StructuredDocumentTags() {
			for _, par := range tag.Paragraphs() {
				for _, run := range par.Runs() {

					_, err = buf.WriteString(run.Text())
					if err != nil {
						return err
					}
				}
				_, err = buf.WriteString("\n")
				if err != nil {
					return err
				}
			}
		}

		outputFileName := filepath.Join(outputFilePath, strings.TrimSuffix(info.Name(), ".docx")) + ".txt"
		fmt.Printf(outputFilePath)
		err = ioutil.WriteFile(outputFileName, buf.Bytes(), 0655)
		if err != nil {
			return err
		}

		return nil
	})

	if err != nil {
		log.Fatal(err)
	}
}