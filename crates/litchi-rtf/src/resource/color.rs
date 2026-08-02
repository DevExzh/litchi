//! Borrowed color-palette views.

use crate::types::{Color as RawColor, ColorRef, ColorTable};
use std::iter::FusedIterator;

/// An RGB color value.
///
/// Components are private so the ordinary facade can evolve independently of
/// the retained RTF representation. Use [`Self::new`] to construct a value and
/// the component accessors to inspect one.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct Color {
    red: u8,
    green: u8,
    blue: u8,
}

impl Color {
    /// Construct an RGB color.
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }

    /// Red component.
    pub const fn red(self) -> u8 {
        self.red
    }

    /// Green component.
    pub const fn green(self) -> u8 {
        self.green
    }

    /// Blue component.
    pub const fn blue(self) -> u8 {
        self.blue
    }

    /// All components in RGB order.
    pub const fn rgb(self) -> [u8; 3] {
        [self.red, self.green, self.blue]
    }

    pub(crate) const fn from_raw(raw: RawColor) -> Self {
        Self::new(raw.red, raw.green, raw.blue)
    }
}

/// One semantic color-table entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Value {
    /// The document consumer chooses the default color.
    Automatic,
    /// An explicit RGB color.
    Rgb(Color),
}

impl Value {
    /// Return the explicit RGB value, or `None` for [`Self::Automatic`].
    pub const fn color(self) -> Option<Color> {
        match self {
            Self::Automatic => None,
            Self::Rgb(color) => Some(color),
        }
    }
}

/// A borrowed document color palette.
///
/// The palette is a small copyable view. Traversal borrows the immutable
/// snapshot and performs no allocation.
#[derive(Clone, Copy)]
pub struct Palette<'a> {
    raw: &'a ColorTable,
}

impl<'a> Palette<'a> {
    pub(crate) const fn new(raw: &'a ColorTable) -> Self {
        Self { raw }
    }

    /// Number of retained palette entries.
    pub fn len(self) -> usize {
        self.raw.colors().len()
    }

    /// Whether the palette has no entries.
    pub fn is_empty(self) -> bool {
        self.raw.colors().is_empty()
    }

    /// Return a checked zero-based palette entry.
    pub fn at(self, position: usize) -> Option<Value> {
        let color = self.raw.colors().get(position).copied()?;
        Some(if self.raw.is_automatic_at(position) {
            Value::Automatic
        } else {
            Value::Rgb(Color::from_raw(color))
        })
    }

    /// Lazily traverse palette entries in source order.
    pub fn iter(self) -> Iter<'a> {
        Iter {
            palette: self,
            front: 0,
            back: self.len(),
        }
    }

    pub(crate) fn resolve(self, reference: ColorRef) -> Option<Value> {
        self.at(usize::from(reference))
    }
}

impl std::fmt::Debug for Palette<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Palette")
            .field("len", &self.len())
            .finish()
    }
}

impl<'a> IntoIterator for Palette<'a> {
    type Item = Value;
    type IntoIter = Iter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

/// Lazy borrowed color traversal.
#[derive(Clone)]
pub struct Iter<'a> {
    palette: Palette<'a>,
    front: usize,
    back: usize,
}

impl Iterator for Iter<'_> {
    type Item = Value;

    fn next(&mut self) -> Option<Self::Item> {
        if self.front >= self.back {
            return None;
        }
        let position = self.front;
        self.front = self.front.saturating_add(1);
        self.palette.at(position)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.back.saturating_sub(self.front);
        (remaining, Some(remaining))
    }
}

impl DoubleEndedIterator for Iter<'_> {
    fn next_back(&mut self) -> Option<Self::Item> {
        if self.front >= self.back {
            return None;
        }
        self.back = self.back.saturating_sub(1);
        self.palette.at(self.back)
    }
}

impl ExactSizeIterator for Iter<'_> {}
impl FusedIterator for Iter<'_> {}
