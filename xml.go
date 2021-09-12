package xlsxscan

import (
	"encoding/xml"
	"io"
)

func NewXMLStreamer(r io.Reader) *XMLStreamer {
	dec := xml.NewDecoder(r)
	return &XMLStreamer{
		dec: dec,
	}
}

type XMLStreamer struct {
	dec *xml.Decoder
}

func (s *XMLStreamer) advanceToNext(tag string) (*xml.StartElement, error) {
	for {
		token, err := s.dec.Token()
		if err != nil {
			return nil, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name.Local == tag {
				return &token, nil
			}
		default:
			continue
		}
	}
}
