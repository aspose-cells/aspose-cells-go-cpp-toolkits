package io

import "os"

func WriteFile(data []byte, path string) error {
	file, errCreate := os.Create(path)
	if errCreate != nil {
		panic(errCreate)
	}
	defer file.Close()

	_, err := file.Write(data)
	if err != nil {
		panic(err)
	}
	return err
}
