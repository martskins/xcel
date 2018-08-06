package xcel

// Style struct
type Style struct {
	Font         *Font      `json:"font"`
	Fill         *Fill      `json:"fill"`
	Border       []*Border  `json:"border"`
	Alignment    *Alignment `json:"alignment"`
	NumberFormat string     `json:"custom_number_format,omitempty"`
}

type Alignment struct {
	Horizontal  string `json:"horizontal"`
	Vertical    string `json:"vertical"`
	ShrinkToFit bool   `json:"shrink_to_fit"`
	WrapText    bool   `json:"wrap_text"`
}

// Font struct
type Font struct {
	Bold      bool   `json:"bold"`
	Italic    bool   `json:"italic"`
	Underline string `json:"underline"`
	Family    string `json:"family"`
	Size      int    `json:"size"`
	Color     string `json:"color"`
}

// Fill struct
type Fill struct {
	Type    string   `json:"type"`
	Pattern int      `json:"pattern"`
	Color   []string `json:"color"`
}

// Border struct
type Border struct {
	Type  string `json:"type"`
	Color string `json:"color"`
	Style int    `json:"style"`
}

// ImageFormat struct
type ImageFormat struct {
	XScale          float64 `json:"x_scale,omitempty"`
	YScale          float64 `json:"y_scale,omitempty"`
	XOffset         float64 `json:"x_offset,omitempty"`
	YOffset         float64 `json:"y_offset,omitempty"`
	LockAspectRatio bool    `json:"lock_aspect_ratio,omitempty"`
	Positioning     string  `json:"positioning,omitempty"`
}

const (
	PatternSolid = 1

	HorizontalAlignmentLeft    = "left"
	HorizontalAlignmentRight   = "right"
	HorizontalAlignmentCenter  = "center"
	HorizontalAlignmentFill    = "fill"
	HorizontalAlignmentJustify = "justify"

	VerticalAlignmentTop     = "left"
	VerticalAlignmentCenter  = "center"
	VerticalAlignmentJustify = "justify"

	UnderlineSingle = "single"
	UnderlineDouble = "double"

	BorderNone             = 0
	BorderDash             = 3
	BorderDashMedium       = 8
	BorderDot              = 4
	BorderContinuous       = 1
	BorderContinuousLight  = 7
	BorderContinuousMedium = 2
	BorderContinuousHeavy  = 5

	PositioningDefault  = ""
	PositioningAbsolute = "absolute"
	PositioningOneCell  = "oneCell"
)
