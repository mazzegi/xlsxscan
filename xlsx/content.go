package xlsx

import (
	"encoding/xml"
	"strings"

	"github.com/pkg/errors"
)

// <?xml version="1.0" encoding="UTF-8"?>
// <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
// 	<Default Extension="xml" ContentType="application/xml"/>
// 	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
// 	<Default Extension="png" ContentType="image/png"/>
// 	<Default Extension="jpeg" ContentType="image/jpeg"/>
// 	<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
// 	<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
// 	<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
// 	<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
// 	<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
// 	<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
// 	<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
// 	<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
// </Types>

const ContentTypesFileName = "[Content_Types].xml"
const PrefixXL = "/xl"

type ContentType string

const (
	ContentTypeXML      ContentType = "application/xml"
	ContentTypeWorkbook ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	ContentTypeSheet    ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	ContentTypeRels     ContentType = "application/vnd.openxmlformats-package.relationships+xml"
)

type DefaultType struct {
	xml.Name    `xml:"Default"`
	Extension   string      `xml:"Extension,attr"`
	ContentType ContentType `xml:"ContentType,attr"`
}

type OverrideType struct {
	xml.Name    `xml:"Override"`
	PartName    string      `xml:"PartName,attr"`
	ContentType ContentType `xml:"ContentType,attr"`
}

type Content struct {
	xml.Name      `xml:"Types"`
	DefaultTypes  []DefaultType  `xml:"Default"`
	OverrideTypes []OverrideType `xml:"Override"`
}

func (cts Content) FindFirstOverrideByContentType(partNamePrefix string, cty ContentType) (OverrideType, error) {
	for _, o := range cts.OverrideTypes {
		if strings.HasPrefix(o.PartName, partNamePrefix) && o.ContentType == cty {
			return o, nil
		}
	}
	return OverrideType{}, errors.Errorf("override with content-type %q not found", cty)
}
