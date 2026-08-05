use crate::{XlsError, XlsResult};

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

/// How a cell-relative anchor responds when its underlying cells change.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum Behavior {
    /// Neither the BIFF `fMove` nor `fSize` flag is set.
    Fixed,
    /// Only the BIFF `fSize` flag is set.
    Size,
    /// Both the BIFF `fMove` and `fSize` flags are set.
    MoveAndSize,
}

impl Behavior {
    const fn bits(self) -> u16 {
        match self {
            Self::Fixed => 0,
            Self::Size => 0b10,
            Self::MoveAndSize => 0b11,
        }
    }
}

/// A checked cell-relative point in a BIFF8 worksheet anchor.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Point {
    row: u16,
    column: u8,
    row_offset: u8,
    column_offset: u16,
}

impl Point {
    /// Create a point at a cell boundary, rejecting locations outside the BIFF8 grid.
    pub fn new(row: u32, column: u16) -> XlsResult<Self> {
        let row = u16::try_from(row)
            .map_err(|_| XlsError::InvalidData("shape anchor row must be <= 65535".to_string()))?;
        let column = u8::try_from(column)
            .map_err(|_| XlsError::InvalidData("shape anchor column must be <= 255".to_string()))?;
        Ok(Self::cell(row, column))
    }

    /// Create a zero-offset point from already narrow BIFF8 coordinates.
    pub const fn cell(row: u16, column: u8) -> Self {
        Self {
            row,
            column,
            row_offset: 0,
            column_offset: 0,
        }
    }

    /// Set the row and column fractions, moving the checked point on success.
    pub fn offset(mut self, row: u16, column: u16) -> XlsResult<Self> {
        let row = u8::try_from(row).map_err(|_| {
            XlsError::InvalidData("shape anchor row offset must be <= 255".to_string())
        })?;
        if column > 1023 {
            return Err(XlsError::InvalidData(
                "shape anchor column offset must be <= 1023".to_string(),
            ));
        }
        self.row_offset = row;
        self.column_offset = column;
        Ok(self)
    }

    /// Return the zero-based row.
    pub const fn row(self) -> u16 {
        self.row
    }

    /// Return the zero-based column.
    pub const fn column(self) -> u8 {
        self.column
    }

    /// Return the offset in 256ths of the row height.
    pub const fn row_offset(self) -> u8 {
        self.row_offset
    }

    /// Return the offset in 1024ths of the column width.
    pub const fn column_offset(self) -> u16 {
        self.column_offset
    }
}

/// A checked cell-relative BIFF8 client anchor.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Anchor {
    behavior: Behavior,
    first: Point,
    last: Point,
}

impl Anchor {
    /// Create an anchor whose horizontal and vertical endpoints are strictly ordered.
    pub fn new(first: Point, last: Point, behavior: Behavior) -> XlsResult<Self> {
        let horizontal_order =
            (first.column, first.column_offset) < (last.column, last.column_offset);
        let vertical_order = (first.row, first.row_offset) < (last.row, last.row_offset);
        if !horizontal_order || !vertical_order {
            return Err(XlsError::InvalidData(
                "shape anchor endpoints must be strictly ordered on both axes".to_string(),
            ));
        }
        Ok(Self {
            behavior,
            first,
            last,
        })
    }

    /// Create a zero-offset anchor directly from wide worksheet coordinates.
    pub fn cells(
        first_row: u32,
        first_column: u16,
        last_row: u32,
        last_column: u16,
        behavior: Behavior,
    ) -> XlsResult<Self> {
        Self::new(
            Point::new(first_row, first_column)?,
            Point::new(last_row, last_column)?,
            behavior,
        )
    }

    /// Return the cell-change behavior.
    pub const fn behavior(self) -> Behavior {
        self.behavior
    }

    /// Return the top-left point.
    pub const fn first(self) -> Point {
        self.first
    }

    /// Return the bottom-right point.
    pub const fn last(self) -> Point {
        self.last
    }

    pub(crate) const fn fields(self) -> [u16; 9] {
        [
            self.behavior.bits(),
            self.first.column as u16,
            self.first.column_offset,
            self.first.row,
            self.first.row_offset as u16,
            self.last.column as u16,
            self.last.column_offset,
            self.last.row,
            self.last.row_offset as u16,
        ]
    }

    pub(crate) fn default_for_cell(row: u16, column: u8) -> Self {
        let first_column = column.saturating_add(1).min(252);
        let first_row = row.min(65_531);
        Self {
            behavior: Behavior::MoveAndSize,
            first: Point::cell(first_row, first_column),
            last: Point::cell(first_row + 4, first_column + 3),
        }
    }
}

/// A checked rectangle in a shape group's child coordinate space.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rect {
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
}

impl Rect {
    pub(crate) const DEFAULT_GROUP: Self = Self {
        left: 0,
        top: 0,
        right: 1023,
        bottom: 255,
    };

    /// Create a rectangle with strictly increasing axes.
    pub fn new(left: i32, top: i32, right: i32, bottom: i32) -> XlsResult<Self> {
        if left >= right || top >= bottom {
            return Err(XlsError::InvalidData(
                "group rectangle must have left < right and top < bottom".to_string(),
            ));
        }
        Ok(Self {
            left,
            top,
            right,
            bottom,
        })
    }

    /// Return the left coordinate.
    pub const fn left(self) -> i32 {
        self.left
    }

    /// Return the top coordinate.
    pub const fn top(self) -> i32 {
        self.top
    }

    /// Return the right coordinate.
    pub const fn right(self) -> i32 {
        self.right
    }

    /// Return the bottom coordinate.
    pub const fn bottom(self) -> i32 {
        self.bottom
    }

    pub(crate) const fn fields(self) -> [i32; 4] {
        [self.left, self.top, self.right, self.bottom]
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
        let mut units = Vec::new();
        units
            .try_reserve_exact(self.value.len())
            .map_err(|_| XlsError::Allocation("reserving shape text validation storage"))?;
        units.extend(self.value.encode_utf16());
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
    pub anchor: Anchor,
    /// Optional requested OBJ identifier. `None` assigns the first free canonical ID.
    pub object_id: Option<u16>,
    pub text: Option<XlsShapeText>,
    pub fill: XlsShapeFill,
    pub line: XlsShapeLine,
    pub visible: bool,
    pub locked: bool,
}

impl XlsShapeWrite {
    pub fn new(kind: XlsShapeKind, anchor: Anchor) -> Self {
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
        validate_shape_style(
            self.kind,
            self.object_id,
            self.fill,
            self.line,
            self.text.as_ref(),
        )
    }
}

/// Validate the primitive style fields shared by top-level and grouped shapes.
pub(crate) fn validate_shape_style(
    kind: XlsShapeKind,
    object_id: Option<u16>,
    fill: XlsShapeFill,
    line: XlsShapeLine,
    text: Option<&XlsShapeText>,
) -> XlsResult<()> {
    if matches!(object_id, Some(0 | u16::MAX)) {
        return Err(XlsError::InvalidData(
            "shape object ID 0 and 65535 are reserved".to_string(),
        ));
    }
    if let XlsShapeLine::Solid { width_emu, .. } = line
        && !(1..=20_116_800).contains(&width_emu)
    {
        return Err(XlsError::InvalidData(
            "shape line width must be 1..=20116800 EMU".to_string(),
        ));
    }
    if kind == XlsShapeKind::Line && (fill != XlsShapeFill::None || text.is_some()) {
        return Err(XlsError::InvalidData(
            "line primitives do not support fill or text".to_string(),
        ));
    }
    if let Some(text) = text {
        text.validate()?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::panic::catch_unwind;

    #[test]
    fn point_boundaries_are_checked_without_unwinding() {
        let outcome = catch_unwind(|| {
            assert!(Point::new(65_535, 255).is_ok());
            assert!(Point::new(65_536, 255).is_err());
            assert!(Point::new(65_535, 256).is_err());
            assert!(Point::cell(u16::MAX, u8::MAX).offset(255, 1023).is_ok());
            assert!(Point::cell(0, 0).offset(256, 1023).is_err());
            assert!(Point::cell(0, 0).offset(255, 1024).is_err());
        });
        assert!(outcome.is_ok());
        assert_eq!(size_of::<Point>(), 6);
    }

    #[test]
    fn anchor_and_group_rect_reject_degenerate_axes() {
        let first = Point::cell(4, 3);
        let right = Point::cell(4, 3).offset(1, 1).unwrap();
        let anchor = Anchor::new(first, right, Behavior::MoveAndSize).unwrap();
        assert_eq!(anchor.first(), first);
        assert_eq!(anchor.last(), right);
        assert_eq!(anchor.behavior(), Behavior::MoveAndSize);

        let outcome = catch_unwind(|| {
            assert!(Anchor::new(first, first, Behavior::Fixed).is_err());
            assert!(Anchor::new(right, first, Behavior::Size).is_err());
            assert!(Anchor::cells(0, 0, 65_536, 255, Behavior::Fixed).is_err());
            assert!(
                Anchor::new(
                    Point::new(65_534, 254).unwrap().offset(255, 1023).unwrap(),
                    Point::new(65_535, 255).unwrap().offset(255, 1023).unwrap(),
                    Behavior::MoveAndSize,
                )
                .is_ok()
            );
        });
        assert!(outcome.is_ok());
        assert!(Rect::new(i32::MIN, i32::MIN, i32::MAX, i32::MAX).is_ok());
        assert!(Rect::new(0, 0, 0, 1).is_err());
        assert!(Rect::new(0, 0, 1, 0).is_err());
    }

    #[test]
    fn text_validation_checks_utf16_boundaries_without_unwinding() {
        let outcome = catch_unwind(|| {
            let oversized = XlsShapeText::new("a".repeat(usize::from(u16::MAX) + 1));
            assert!(oversized.validate().is_err());

            let split_surrogate = XlsShapeText {
                value: "😀".to_string(),
                runs: vec![XlsShapeTextRun {
                    character_index: 1,
                    font_index: 0,
                }],
                font_when_empty: 0,
            };
            assert!(split_surrogate.validate().is_err());
        });
        assert!(outcome.is_ok());
    }
}
