//! Compatibility re-exports for the worksheet-view model.
//!
//! The typed values are owned by [`litchi_xlsx`]. This module remains at the
//! historical OOXML path so the host parser, writer, and existing callers can
//! migrate without a public-path break.

pub use litchi_xlsx::views::{
    SheetPane, SheetPanePosition, SheetPaneState, SheetSelection, SheetView, SheetViewType,
};
