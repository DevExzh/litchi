use crate::xls::{XlsError, XlsResult};

/// Safe, inert OfficeArt primitive supported by the BIFF8 writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsShapeKind {
    Rectangle,
    RoundedRectangle,
    Ellipse,
    Line,
    TextBox,
}

impl XlsShapeKind {
    pub(crate) const fn officeart_type(self) -> u16 {
        match self {
            Self::Rectangle => 1,
            Self::RoundedRectangle => 2,
            Self::Ellipse => 3,
            Self::Line => 20,
            Self::TextBox => 202,
        }
    }

    pub(crate) const fn object_type(self) -> u16 {
        match self {
            Self::Line => 1,
            Self::Rectangle | Self::RoundedRectangle => 2,
            Self::Ellipse => 3,
            Self::TextBox => 6,
        }
    }
}

/// RGB color used by a primitive shape fill or outline.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsShapeColor {
    pub red: u8,
    pub green: u8,
    pub blue: u8,
}

impl XlsShapeColor {
    pub const fn rgb(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }

    pub(crate) const fn officeart_color(self) -> u32 {
        self.red as u32 | ((self.green as u32) << 8) | ((self.blue as u32) << 16)
    }
}

/// Fill style for a writable primitive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsShapeFill {
    None,
    Solid(XlsShapeColor),
}

/// Outline style for a writable primitive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsShapeLine {
    None,
    Solid {
        color: XlsShapeColor,
        /// Line width in English Metric Units.
        width_emu: u32,
    },
}

/// Cell-relative BIFF8 client anchor.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsShapeAnchor {
    pub move_with_cells: bool,
    pub size_with_cells: bool,
    pub first_column: u16,
    pub first_column_offset: u16,
    pub first_row: u32,
    pub first_row_offset: u16,
    pub last_column: u16,
    pub last_column_offset: u16,
    pub last_row: u32,
    pub last_row_offset: u16,
}

impl XlsShapeAnchor {
    pub(crate) fn validate(self) -> XlsResult<()> {
        let horizontal_order = (self.first_column, self.first_column_offset)
            < (self.last_column, self.last_column_offset);
        let vertical_order =
            (self.first_row, self.first_row_offset) < (self.last_row, self.last_row_offset);
        if self.first_column > 255
            || self.last_column > 255
            || self.first_row > 65_535
            || self.last_row > 65_535
            || self.first_column_offset > 1023
            || self.last_column_offset > 1023
            || self.first_row_offset > 255
            || self.last_row_offset > 255
            || !horizontal_order
            || !vertical_order
            || (self.move_with_cells && !self.size_with_cells)
        {
            return Err(XlsError::InvalidData(
                "shape anchor is outside BIFF8 bounds or has invalid movement flags".to_string(),
            ));
        }
        Ok(())
    }
}

/// One rich-text run in a writable shape's TXO text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsShapeTextRun {
    /// UTF-16 code-unit position where this run starts.
    pub character_index: u16,
    pub font_index: u16,
}

/// Optional TXO text attached to a primitive shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsShapeText {
    pub value: String,
    pub runs: Vec<XlsShapeTextRun>,
    pub font_when_empty: u16,
}

impl XlsShapeText {
    pub fn new(value: impl Into<String>) -> Self {
        Self {
            value: value.into(),
            runs: Vec::new(),
            font_when_empty: 0,
        }
    }

    fn validate(&self) -> XlsResult<()> {
        let units = self.value.encode_utf16().collect::<Vec<_>>();
        if units.len() > usize::from(u16::MAX) {
            return Err(XlsError::InvalidData(
                "shape text exceeds 65535 UTF-16 code units".to_string(),
            ));
        }
        if units.is_empty() && !self.runs.is_empty() {
            return Err(XlsError::InvalidData(
                "empty shape text cannot contain formatting runs".to_string(),
            ));
        }
        if !self.runs.is_empty() {
            if self.runs[0].character_index != 0 {
                return Err(XlsError::InvalidData(
                    "the first shape formatting run must start at character zero".to_string(),
                ));
            }
            let mut previous = None;
            for run in &self.runs {
                let index = usize::from(run.character_index);
                if index >= units.len()
                    || previous.is_some_and(|value| value >= run.character_index)
                    || (index > 0 && (0xDC00..=0xDFFF).contains(&units[index]))
                {
                    return Err(XlsError::InvalidData(
                        "shape formatting runs must be ordered UTF-16 character boundaries"
                            .to_string(),
                    ));
                }
                previous = Some(run.character_index);
            }
        }
        let run_count = if units.is_empty() {
            0
        } else {
            self.runs.len().max(1)
        };
        let run_bytes = run_count
            .checked_add(1)
            .and_then(|count| count.checked_mul(8))
            .ok_or_else(|| XlsError::InvalidData("shape formatting run size overflows".into()))?;
        if run_bytes > 65_528 {
            return Err(XlsError::InvalidData(
                "shape formatting runs exceed the BIFF8 cbRuns limit".to_string(),
            ));
        }
        Ok(())
    }
}

/// Writable, macro-inert BIFF8 worksheet primitive.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsShapeWrite {
    pub kind: XlsShapeKind,
    pub anchor: XlsShapeAnchor,
    /// Optional requested OBJ identifier. `None` assigns the first free canonical ID.
    pub object_id: Option<u16>,
    pub text: Option<XlsShapeText>,
    pub fill: XlsShapeFill,
    pub line: XlsShapeLine,
    pub visible: bool,
    pub locked: bool,
}

impl XlsShapeWrite {
    pub fn new(kind: XlsShapeKind, anchor: XlsShapeAnchor) -> Self {
        Self {
            kind,
            anchor,
            object_id: None,
            text: None,
            fill: XlsShapeFill::Solid(XlsShapeColor::rgb(255, 255, 255)),
            line: XlsShapeLine::Solid {
                color: XlsShapeColor::rgb(0, 0, 0),
                width_emu: 12_700,
            },
            visible: true,
            locked: true,
        }
    }

    pub(crate) fn validate(&self) -> XlsResult<()> {
        self.anchor.validate()?;
        if matches!(self.object_id, Some(0 | u16::MAX)) {
            return Err(XlsError::InvalidData(
                "shape object ID 0 and 65535 are reserved".to_string(),
            ));
        }
        if let XlsShapeLine::Solid { width_emu, .. } = self.line
            && !(1..=20_116_800).contains(&width_emu)
        {
            return Err(XlsError::InvalidData(
                "shape line width must be 1..=20116800 EMU".to_string(),
            ));
        }
        if self.kind == XlsShapeKind::Line
            && (self.fill != XlsShapeFill::None || self.text.is_some())
        {
            return Err(XlsError::InvalidData(
                "line primitives do not support fill or text".to_string(),
            ));
        }
        if let Some(text) = &self.text {
            text.validate()?;
        }
        Ok(())
    }
}
