package xlsx

import "encoding/xml"

// <sheet name="Sheet1" sheetId="1" state="visible" r:id="rId2"/>

type Sheet struct {
	xml.Name  `xml:"sheet"`
	SheetName string `xml:"name,attr"`
	ID        string `xml:"sheetId,attr"`
	RID       string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

type Workbook struct {
	xml.Name `xml:"workbook"`
	Sheets   []Sheet `xml:"sheets>sheet"`
}

func (wb Workbook) FindSheet(name string) (Sheet, bool) {
	for _, s := range wb.Sheets {
		if s.SheetName == name {
			return s, true
		}
	}
	return Sheet{}, false
}

//
// <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>

type WorkbookRel struct {
	xml.Name `xml:"Relationship"`
	ID       string `xml:"Id,attr"`
	Type     string `xml:"Type,attr"`
	Target   string `xml:"Target,attr"`
}

type WorkbookRels struct {
	xml.Name `xml:"Relationships"`
	Rels     []WorkbookRel `xml:"Relationship"`
}

func (wrs WorkbookRels) FindRel(id string) (WorkbookRel, bool) {
	for _, r := range wrs.Rels {
		if r.ID == id {
			return r, true
		}
	}
	return WorkbookRel{}, false
}
