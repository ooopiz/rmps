package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	if len(os.Args) != 3 {
		fmt.Println("arg1: excelfile, arg2: pwdfile")
		os.Exit(0)
	}

	xlsx := os.Args[1]
	pwdfile := os.Args[2]

	if _, err := os.Stat(xlsx); os.IsNotExist(err) {
		fmt.Println(xlsx + " does not exist")
		os.Exit(0)
	}
	if _, err := os.Stat(pwdfile); os.IsNotExist(err) {
		fmt.Println(pwdfile + " does not exist")
		os.Exit(0)
	}

	fileIO, err := os.OpenFile(pwdfile, os.O_RDWR, 0600)
	if err != nil {
		panic(err)
	}
	defer fileIO.Close()
	rawBytes, err := ioutil.ReadAll(fileIO)
	if err != nil {
		panic(err)
	}
	lines := strings.Split(string(rawBytes), "\n")
	pwd := lines[0]

	f, err := excelize.OpenFile(xlsx, excelize.Options{Password: pwd})
	if err != nil {
		fmt.Println(err)
		return
	}
	if err := f.Save(); err != nil {
		fmt.Println(err)
	}
}
