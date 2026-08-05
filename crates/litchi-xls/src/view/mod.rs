//! Semantic and lossless worksheet-window support for BIFF8.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub(crate) use codec::ViewCollector;
pub(crate) use model::pane_exists;
pub use model::{Pane, PaneType, Range, Selection, View};

pub(crate) const WINDOW2_RECORD_TYPE: u16 = 0x023e;
pub(crate) const SCL_RECORD_TYPE: u16 = 0x00a0;
pub(crate) const PANE_RECORD_TYPE: u16 = 0x0041;
pub(crate) const SELECTION_RECORD_TYPE: u16 = 0x001d;
