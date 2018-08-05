package excel

import "encoding/json"

// Sheet represents a sheet in a file
type Sheet struct {
	Name string
	File *File
	Rows []*Row
}

// AddRow adds a row to the sheet
func (s *Sheet) AddRow() *Row {
	r := Row{
		ID:    len(s.Rows) + 1,
		Sheet: s,
	}
	s.Rows = append(s.Rows, &r)

	return &r
}

// SetStyle applies a style to the given cell range in this sheet
func (s *Sheet) SetStyle(ulCell, lrCell string, styleID int) *Sheet {
	s.File.SetCellStyle(s.Name, ulCell, lrCell, styleID)
	return s
}

// AddPicture adds a picture to the given cell in the sheet
func (s *Sheet) AddPicture(cell, picture string, format ImageFormat) error {
	bts, err := json.Marshal(format)
	if err != nil {
		return err
	}

	return s.File.AddPicture(s.Name, cell, picture, string(bts))
}
