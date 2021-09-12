package xlsxscan

import (
	"fmt"
	"io"
	"strconv"
	"strings"

	"github.com/mazzegi/xlsxscan/xlsx"
	"github.com/pkg/errors"
)

func makeValue(s string, dt string) interface{} {
	switch dt {
	case dataTypeInlineString:
		return s
	case dataTypeNumber:
		n, err := strconv.ParseInt(s, 10, 64)
		if err != nil {
			return nil
		}
		return n
	default:
		return s
	}
}

type Cell struct {
	Ref   string
	Value interface{}
}

func (c Cell) String() string {
	return fmt.Sprintf("%s:%v", c.Ref, c.Value)
}

type Row struct {
	Number int
	Cells  []Cell
}

func (r Row) String() string {
	var sl []string
	for _, c := range r.Cells {
		sl = append(sl, c.String())
	}
	return fmt.Sprintf("%d: [%s]", r.Number, strings.Join(sl, ", "))
}

func NewRowScanner(rc io.ReadCloser) (*RowScanner, error) {
	xs := NewXMLStreamer(rc)
	_, err := xs.advanceToNext(tagSheetData)
	if err != nil {
		return nil, errors.Errorf("no sheet-data found")
	}

	return &RowScanner{
		rc: rc,
		xs: xs,
	}, nil
}

type RowScanner struct {
	rc io.ReadCloser
	xs *XMLStreamer
}

func (rs *RowScanner) Close() {
	rs.rc.Close()
}

func (rs *RowScanner) Scan() (Row, bool) {
	se, err := rs.xs.advanceToNext(tagRow)
	if err != nil {
		return Row{}, false
	}
	var xrow xlsx.Row
	err = rs.xs.dec.DecodeElement(&xrow, se)
	if err != nil {
		return Row{}, false
	}

	row := Row{
		Number: xrow.Number,
	}
	for _, xc := range xrow.Cells {
		row.Cells = append(row.Cells, Cell{
			Ref:   xc.Reference,
			Value: makeValue(xc.Value, xc.Type),
		})
	}

	return row, true
}
