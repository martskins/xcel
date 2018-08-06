package xcel

// Row represents a row in a sheet
type Row struct {
	ID    int
	Sheet *Sheet
	Cells []*Cell
}

// GetFile gets the file the row belongs to
func (r *Row) GetFile() *File {
	return r.Sheet.File
}

// GetSheet gets the sheet the row belongs to
func (r *Row) GetSheet() *Sheet {
	return r.Sheet
}

// AddCell adds a cell to the given row
func (r *Row) AddCell() *Cell {
	cell := Cell{
		ID:  len(r.Cells),
		Row: r,
	}

	r.Cells = append(r.Cells, &cell)
	return &cell
}

func (r *Row) SetHeight(val float64) {
	r.GetFile().SetRowHeight(r.GetSheet().Name, r.ID, val)
}
