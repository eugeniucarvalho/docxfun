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
	zipFileReader *zip.ReadCloser
	zipReader     *zip.Reader
	Files         []*zip.File
	FilesContent  map[string][]byte
	WordsList     []*Words
}

type Words struct {
	Pid       string
	RawString string
	Content   []string
}

//OpenUploaded
func OpenDocxByte(buff []byte) (*Docx, error) {

	reader, err := zip.NewReader(bytes.NewReader(buff), int64(len(buff)))
	if err != nil {
		return nil, err
	}

	wordDoc := Docx{
		zipReader:    reader,
		Files:        reader.File,
		FilesContent: map[string][]byte{},
	}

	for _, f := range wordDoc.Files {
		contents, _ := wordDoc.retrieveFileContents(f.Name)
		wordDoc.FilesContent[f.Name] = contents
	}

	return &wordDoc, nil
}

//OpenDocx open and load all files content
func OpenDocx(path string) (*Docx, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}

	wordDoc := Docx{
		zipFileReader: reader,
		Files:         reader.File,
		FilesContent:  map[string][]byte{},
	}

	for _, f := range wordDoc.Files {
		contents, _ := wordDoc.retrieveFileContents(f.Name)
		wordDoc.FilesContent[f.Name] = contents
	}

	return &wordDoc, nil
}

//Close is close reader
func (d *Docx) Close() error {
	return d.zipFileReader.Close()
}

//Read all files contents
func (d *Docx) retrieveFileContents(filename string) ([]byte, error) {
	var file *zip.File
	for _, f := range d.Files {
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

	for fName, content := range d.FilesContent {
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
	document := d.FilesContent[fileName]
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
	d.FilesContent["word/document.xml"] = document
	return nil
}

//ListWording list out all text in tag <w:t> </w:t>
func (d *Docx) ListWording() (result []string, err error) {
	xmlData := d.FilesContent["word/document.xml"]
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
			continue
			// return nil, errors.New(fmt.Sprintf("Non string type found %T, %v", v, v))
		}
	}
	return result, err
}

func (d *Docx) GetWording() (err error) {
	xmlData := string(d.FilesContent["word/document.xml"])
	listP(xmlData, d)
	return nil
}

func getT(item []string, d *Docx) {
	data := item[1]
	pId := item[0]
	re := regexp.MustCompile(`(?U)(<w:t>|<w:t .*>)(.*)(</w:t>)`)
	w := new(Words)
	w.RawString = data
	w.Pid = pId
	content := []string{}
	for _, match := range re.FindAllStringSubmatch(string(data), -1) {
		content = append(content, match[2])
	}
	w.Content = content
	d.WordsList = append(d.WordsList, w)
}
func hasP(data string) bool {
	re := regexp.MustCompile(`(?U)<w:p w:rsidR="(\w*)"[^>]*>(.*)</w:p>`)
	result := re.MatchString(data)
	fmt.Println("yes?", result)
	return result
}

func listP(data string, d *Docx) {
	result := [][]string{}
	fmt.Println("in Listp")
	re := regexp.MustCompile(`(?U)<w:p w:rsidR="(\w*)"[^>]*>(.*)</w:p>`)
	for _, match := range re.FindAllStringSubmatch(string(data), -1) {
		result = append(result, []string{match[1], match[2]})
	}
	for _, item := range result {
		if hasP(item[1]) {
			listP(item[1], d)
			continue
		}
		getT(item, d)
	}
	return
}
