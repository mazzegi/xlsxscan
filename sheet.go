package xlsxscan

import (
	"archive/zip"

	"github.com/pkg/errors"
)

const (
	tagSheetData = "sheetData"
	tagRow       = "row"
	tagCol       = "c"
	tagValue     = "v"
)

const (
	dataTypeBool           = "b"
	dataTypeInlineString   = "inlineStr"
	dataTypeNumber         = "n"
	dataTypeSharedString   = "s"
	dataTypeFormularString = "str"
)

func NewSheetReader(file *zip.File) *SheetReader {
	return &SheetReader{
		file: file,
	}
}

type SheetReader struct {
	file *zip.File
}

func (sr SheetReader) OpenRowScanner() (*RowScanner, error) {
	fz, err := sr.file.Open()
	if err != nil {
		return nil, errors.Wrapf(err, "open zip-file %q", sr.file.Name)
	}
	rs, err := NewRowScanner(fz)
	if err != nil {
		return nil, errors.Wrapf(err, "new-row-scanner")
	}
	return rs, nil
}
