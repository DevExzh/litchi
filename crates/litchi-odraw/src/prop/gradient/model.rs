use crate::prop::{Array, ColorRef};

/// Maximum number of shade stops accepted by the typed decoder or encoder.
///
/// The OfficeArt array count is a `u16`, but the smaller explicit ceiling keeps
/// a hostile property table from turning one optional visual property into an
/// unbounded validation walk.
pub const MAX_STOPS: usize = 4096;

/// A checked unsigned 16.16 gradient position.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Position(i32);

impl Position {
    /// The raw fixed-point value for the beginning of the gradient.
    pub const START: Self = Self(0);
    /// The raw fixed-point value for the end of the gradient.
    pub const END: Self = Self(1 << 16);

    /// Creates a position when it is within the inclusive `[0.0, 1.0]` range
    /// required by `[MS-ODRAW]`.
    pub const fn new(raw: i32) -> Option<Self> {
        if raw >= 0 && raw <= Self::END.0 {
            Some(Self(raw))
        } else {
            None
        }
    }

    /// Returns the exact signed 16.16 wire value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// One lossless `MSOSHADECOLOR` element.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Stop {
    color: ColorRef,
    position: Position,
}

impl Stop {
    /// Creates a shade stop from already checked semantic values.
    pub const fn new(color: ColorRef, position: Position) -> Self {
        Self { color, position }
    }

    /// Returns the exact OfficeArt color reference, including indirect flags.
    pub const fn color(self) -> ColorRef {
        self.color
    }

    /// Returns the checked relative gradient position.
    pub const fn position(self) -> Position {
        self.position
    }
}

/// A borrowed, validated `fillShadeColors_complex` array.
#[derive(Debug, Clone, Copy)]
pub struct Stops<'data> {
    array: Array<'data>,
}

impl<'data> Stops<'data> {
    pub(crate) const fn from_array(array: Array<'data>) -> Self {
        Self { array }
    }

    /// Returns the number of shade stops.
    pub fn len(&self) -> usize {
        self.array.element_count() as usize
    }

    /// Returns whether the array contains no shade stops.
    pub fn is_empty(&self) -> bool {
        self.array.element_count() == 0
    }

    /// Returns one typed shade stop without copying its source array.
    pub fn get(&self, index: usize) -> Option<Stop> {
        let element = self.array.get_element(index)?;
        let color = u32::from_le_bytes(element.get(..4)?.try_into().ok()?);
        let position = i32::from_le_bytes(element.get(4..8)?.try_into().ok()?);
        Some(Stop::new(
            ColorRef::from_raw(color),
            Position::new(position)?,
        ))
    }

    /// Iterates over typed stops in source order without allocating.
    pub fn iter(&self) -> impl Iterator<Item = Stop> + '_ {
        (0..self.len()).filter_map(|index| self.get(index))
    }

    /// Returns the exact source `IMsoArray`, including its header.
    pub fn array(&self) -> Array<'data> {
        self.array
    }

    /// Returns the exact source bytes for lossless replay.
    pub fn payload(&self) -> &'data [u8] {
        self.array.raw_data()
    }
}
