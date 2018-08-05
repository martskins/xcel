package excel

import (
	"encoding/json"
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// Cell represents a cell in a row
type Cell struct {
	ID  int
	Row *Row
}

// GetFile gets the file that the cell belongs to
func (c *Cell) GetFile() *File {
	return c.Row.Sheet.File
}

// GetSheet gets the sheet that the cell belongs to
func (c *Cell) GetSheet() *Sheet {
	return c.Row.Sheet
}

// GetRow gets the row that the cell belongs to
func (c *Cell) GetRow() *Row {
	return c.Row
}

// GetAxis gets the axis (ColRow)
func (c *Cell) GetAxis() string {
	col := excelize.ToAlphaString(c.ID)
	return fmt.Sprintf("%s%d", col, c.Row.ID)
}

// SetString sets a string value in the cell
// returns the modified cell
func (c *Cell) SetString(val string) *Cell {
	c.GetFile().SetCellStr(c.GetSheet().Name, c.GetAxis(), val)
	return c
}

// SetInt sets an int value in the cell
// returns the modified cell
func (c *Cell) SetInt(val int) *Cell {
	c.GetFile().SetCellInt(c.GetSheet().Name, c.GetAxis(), val)
	return c
}

// SetBool sets a bool value in the cell
// returns the modified cell
func (c *Cell) SetBool(val bool) *Cell {
	c.GetFile().SetCellBool(c.GetSheet().Name, c.GetAxis(), val)
	return c
}

// SetFormula sets a formula in the cell
// returns the modified cell
func (c *Cell) SetFormula(val string) *Cell {
	c.GetFile().SetCellFormula(c.GetSheet().Name, c.GetAxis(), val)
	return c
}

// SetFloat sets a float in the cell
// returns the modified cell
func (c *Cell) SetFloat(val float64) *Cell {
	c.GetFile().SetCellValue(c.GetSheet().Name, c.GetAxis(), val)
	return c
}

// SetStyle applies a style to the cell
func (c *Cell) SetStyle(styleID int) *Cell {
	c.GetFile().SetCellStyle(c.GetSheet().Name, c.GetAxis(), c.GetAxis(), styleID)
	return c
}

// AddPicture adds a picture to the cell
func (c *Cell) AddPicture(picture string, format ImageFormat) error {
	bts, err := json.Marshal(format)
	if err != nil {
		return err
	}

	return c.GetFile().AddPicture(c.GetSheet().Name, c.GetAxis(), picture, string(bts))
}
