package xlsxscan

import (
	"archive/zip"
	"encoding/xml"
	"path/filepath"
	"strings"

	"github.com/mazzegi/xlsxscan/xlsx"
	"github.com/pkg/errors"
)

func OpenFile(fileName string) (*Reader, error) {
	zr, err := zip.OpenReader(fileName)
	if err != nil {
		return nil, errors.Wrapf(err, "zip: open-reader from %q", fileName)
	}
	r := &Reader{
		zipReader: zr,
	}
	err = r.init()
	if err != nil {
		return nil, errors.Wrap(err, "init")
	}
	return r, nil
}

type Reader struct {
	zipReader    *zip.ReadCloser
	content      xlsx.Content
	workbookDir  string
	workbook     xlsx.Workbook
	workbookRels xlsx.WorkbookRels
}

func (r *Reader) Close() {
	r.zipReader.Close()
}

func (r *Reader) OpenSheetByName(name string) (*SheetReader, error) {
	sheet, ok := r.workbook.FindSheet(name)
	if !ok {
		return nil, errors.Errorf("no such sheet %q", name)
	}
	sheetRel, ok := r.workbookRels.FindRel(sheet.RID)
	if !ok {
		return nil, errors.Errorf("no rel for sheet sheet %q", name)
	}
	//resolve path relative to workbook dir
	sheetPath := filepath.Join(r.workbookDir, sheetRel.Target)
	f, err := r.findZipFile(sheetPath)
	if err != nil {
		return nil, errors.Errorf("found no %q file", sheetPath)
	}
	return NewSheetReader(f), nil
	// fz, err := f.Open()
	// if err != nil {
	// 	return nil, errors.Wrapf(err, "open zip-file %q", sheetPath)
	// }
	// return NewSheetReader(fz), nil

}

func (r *Reader) findZipFile(path string) (*zip.File, error) {
	// strip leading "/"" to match zip-file naming
	path = strings.TrimPrefix(path, "/")

	for _, f := range r.zipReader.File {
		if f.Name == path {
			return f, nil
		}
	}
	return nil, errors.Errorf("file not found in zip %q", path)
}

func (r *Reader) init() error {
	err := r.initContent()
	if err != nil {
		return errors.Wrap(err, "init-content")
	}
	err = r.initWorkbook()
	if err != nil {
		return errors.Wrap(err, "init-workbook")
	}
	err = r.initWorkbookRels()
	if err != nil {
		return errors.Wrap(err, "init-workbook-rels")
	}
	return nil
}

func (r *Reader) initContent() error {
	// read content
	f, err := r.findZipFile(xlsx.ContentTypesFileName)
	if err != nil {
		return errors.Errorf("found no %q file", xlsx.ContentTypesFileName)
	}
	fz, err := f.Open()
	if err != nil {
		return errors.Wrapf(err, "open zip-file %q", xlsx.ContentTypesFileName)
	}
	defer fz.Close()

	var cts xlsx.Content
	err = xml.NewDecoder(fz).Decode(&cts)
	if err != nil {
		return errors.Wrapf(err, "decode %q", xlsx.ContentTypesFileName)
	}
	r.content = cts
	return nil
}

func (r *Reader) initWorkbook() error {
	// read workbook
	or, err := r.content.FindFirstOverrideByContentType(xlsx.PrefixXL, xlsx.ContentTypeWorkbook)
	if err != nil {
		return errors.Errorf("no workbook found")
	}
	f, err := r.findZipFile(or.PartName)
	if err != nil {
		return errors.Wrapf(err, "find workbook zip-file %q", or.PartName)
	}
	fz, err := f.Open()
	if err != nil {
		return errors.Wrapf(err, "open zip-file %q", or.PartName)
	}
	defer fz.Close()
	var wb xlsx.Workbook
	err = xml.NewDecoder(fz).Decode(&wb)
	if err != nil {
		return errors.Wrapf(err, "decode %q", or.PartName)
	}
	r.workbook = wb
	r.workbookDir = filepath.Dir(or.PartName)
	return nil
}

func (r *Reader) initWorkbookRels() error {
	// read workbook rels
	or, err := r.content.FindFirstOverrideByContentType(xlsx.PrefixXL, xlsx.ContentTypeRels)
	if err != nil {
		return errors.Errorf("no workbook found")
	}
	f, err := r.findZipFile(or.PartName)
	if err != nil {
		return errors.Wrapf(err, "find workbook zip-file %q", or.PartName)
	}
	fz, err := f.Open()
	if err != nil {
		return errors.Wrapf(err, "open zip-file %q", or.PartName)
	}
	defer fz.Close()
	var wbrs xlsx.WorkbookRels
	err = xml.NewDecoder(fz).Decode(&wbrs)
	if err != nil {
		return errors.Wrapf(err, "decode %q", or.PartName)
	}
	r.workbookRels = wbrs
	return nil
}
