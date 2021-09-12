package xlsx

import "encoding/xml"

// <row r="3" customFormat="false" ht="12.8" hidden="false" customHeight="false" outlineLevel="0" collapsed="false">
// 			<c r="B3" s="0" t="s">
// 				<v>0</v>
// 			</c>
// 			<c r="C3" s="0" t="s">
// 				<v>1</v>
// 			</c>
// 			<c r="D3" s="0" t="s">
// 				<v>2</v>
// 			</c>
// 		</row>

type Cell struct {
	xml.Name  `xml:"c"`
	Reference string `xml:"r,attr"`
	Type      string `xml:"t,attr"`
	Value     string `xml:"v"`
}

type Row struct {
	xml.Name `xml:"row"`
	Number   int    `xml:"r,attr"`
	Cells    []Cell `xml:"c"`
}
