use super::super::Stream;
use std::ops::{Deref, DerefMut};

/// Mutable slice guard that preserves pristine replay until mutation occurs.
pub struct Edit<'a, T> {
    pub(in crate::chart) values: &'a mut [T],
    pub(in crate::chart) dirty: &'a mut bool,
    pub(in crate::chart) parsed: bool,
}

impl<T> Deref for Edit<'_, T> {
    type Target = [T];

    fn deref(&self) -> &Self::Target {
        self.values
    }
}

impl<T> DerefMut for Edit<'_, T> {
    fn deref_mut(&mut self) -> &mut Self::Target {
        if self.parsed {
            *self.dirty = true;
        }
        self.values
    }
}

/// Legend rectangle and layout properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Legend {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
    pub position: u8,
    pub spacing: u8,
    pub flags: u16,
}

/// Opaque data-label record retained as part of the semantic inventory.
#[derive(Debug, PartialEq, Eq)]
pub struct Label {
    pub kind: litchi_biff::Kind,
    pub data: Vec<u8>,
}

/// Opaque record observed during parsing, in original record order.
#[derive(Debug, PartialEq, Eq)]
pub struct Raw {
    kind: litchi_biff::Kind,
    data: Vec<u8>,
    offset: usize,
}

impl Raw {
    /// BIFF record identifier.
    #[must_use]
    pub const fn kind(&self) -> litchi_biff::Kind {
        self.kind
    }

    /// Exact record payload.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Original byte offset in the chart substream.
    #[must_use]
    pub const fn offset(&self) -> usize {
        self.offset
    }

    pub(in crate::chart) fn parsed(kind: litchi_biff::Kind, data: Vec<u8>, offset: usize) -> Self {
        Self { kind, data, offset }
    }
}

#[derive(Debug)]
pub(in crate::chart) enum Origin {
    Fresh,
    Parsed(Stream),
}
