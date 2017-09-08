package docxfun

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"html"
	"io/ioutil"
	"os"
	"regexp"

	"github.com/clbanning/mxj"
)

//Docx zip struct
type Docx struct {
	zipReader    *zip.ReadCloser
	files        []*zip.File
	filesContent map[string][]byte
}

//OpenDocx open and load all files content
func OpenDocx(path string) (*Docx, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}

	wordDoc := Docx{
		zipReader:    reader,
		files:        reader.File,
		filesContent: map[string][]byte{},
	}

	for _, f := range wordDoc.files {
		contents, _ := wordDoc.retrieveFileContents(f.Name)
		wordDoc.filesContent[f.Name] = contents
	}

	return &wordDoc, nil
}

//Close is close reader
func (d *Docx) Close() error {
	return d.zipReader.Close()
}

//Read all files contents
func (d *Docx) retrieveFileContents(filename string) ([]byte, error) {
	var file *zip.File
	for _, f := range d.files {
		if f.Name == filename {
			file = f
		}
	}

	if file == nil {
		return []byte{}, errors.New(filename + " file not found")
	}

	reader, err := file.Open()
	if err != nil {
		return []byte{}, err
	}

	return ioutil.ReadAll(reader)
}

//Save files to new docx file
func (d *Docx) Save(fileName string) error {
	// Create a buffer to write our archive to.
	buf := new(bytes.Buffer)

	// Create a new zip archive.
	w := zip.NewWriter(buf)

	for fName, content := range d.filesContent {
		f, err := w.Create(fName)
		if err != nil {
			return err
		}
		_, err = f.Write([]byte(content))
		if err != nil {
			return err
		}
	}
	err := w.Close()
	if err != nil {
		return err
	}

	//write to file
	zipfile, err := os.Create(fileName)
	zipfile.Write(buf.Bytes())
	zipfile.Close()
	return nil
}

//Do document replacement, default word/document.xml
func (d *Docx) DocumentReplace(fileName string, replaceMap map[string]string) error {
	if fileName == "" {
		fileName = "word/document.xml"
	}
	document := d.filesContent[fileName]
	//replace <w:t *>$content </w:t>
	//regex: (<w:t.*>)(.*)(</w:t>)
	for k, v := range replaceMap {
		newK := html.EscapeString(k)
		match := fmt.Sprintf(`(<w:t.*>)%s(</w:t>)`, newK)
		repString := []byte(fmt.Sprintf(`$1 %s $2`, v))
		r := regexp.MustCompile(match)
		docStr := r.ReplaceAll(document, repString)
		document = docStr
	}
	d.filesContent["word/document.xml"] = document
	return nil
}

//ListWording list out all text in tag <w:t> </w:t>
func (d *Docx) ListWording() (result []string, err error) {
	xmlData := d.filesContent["word/document.xml"]
	m, err := mxj.NewMapXml(xmlData)
	if err != nil {
		return nil, err
	}

	wList, err := m.ValuesForKey("t")
	if err != nil {
		return nil, err
	}
	for _, item := range wList {
		switch v := item.(type) {
		case string:
			result = append(result, v)
		default:
			fmt.Printf("no string type: %T", v)
		}
	}
	return result, err
}
