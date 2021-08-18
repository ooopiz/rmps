package main

import (
	"bufio"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

func getTxtContent(txt string) []string {
	file, err := os.Open(txt)

	if err != nil {
		log.Fatalf("failed opening file: %s", err)
	}

	scanner := bufio.NewScanner(file)
	scanner.Split(bufio.ScanLines)
	var txtlines []string

	for scanner.Scan() {
		txtlines = append(txtlines, scanner.Text())
	}

	file.Close()

	return txtlines
	// for _, eachline := range txtlines {
	// fmt.Println(eachline)
	// }
}

func main() {
	if len(os.Args) != 2 {
		fmt.Println("rmps filename.xlsx")
		os.Exit(0)
	}

	xlsx := os.Args[1]
	pwdfile := "password.txt"

	if _, err := os.Stat(xlsx); os.IsNotExist(err) {
		fmt.Println(xlsx + " does not exist")
		os.Exit(0)
	}
	if _, err := os.Stat(pwdfile); os.IsNotExist(err) {
		fmt.Println(pwdfile + " does not exist")
		os.Exit(0)
	}

	txtlines := getTxtContent(pwdfile)
	for _, eachline := range txtlines {
		if eachline != "" {
			f, err := excelize.OpenFile(xlsx, excelize.Options{Password: eachline})
			if err != nil {
				fmt.Println(err)
				continue
			}
			if err := f.Save(); err != nil {
				fmt.Println(err)
			}
			fmt.Println("success")
			break
		}
	}

}
