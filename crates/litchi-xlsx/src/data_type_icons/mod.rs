//! Source-bound, inert support for the MS-XLSX data-type-icon extensions.
//!
//! The two schema elements are carried by existing SpreadsheetML `ext` owners:
//! `showDataTypeIcons` belongs to a worksheet `sheetView`, while
//! `showDataTypeIconsCustomSheetView` belongs to a Named Sheet View.  This
//! module only parses and edits that metadata.  It never computes or displays
//! an icon and never performs a refresh or other external action.
//!
//! [`crate::Package::edit_data_type_icons`] starts an [`Edit`] of one
//! worksheet's metadata; [`Edit::commit`] returns a [`Commit`] holding the
//! new package, the exact [`Patch`] and the [`Diagnostics`], and
//! [`crate::Package::apply_data_type_icons`] replays a patch. The source
//! package never changes.

pub(crate) mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use model::ShowDataTypeIcons;
#[cfg(test)]
pub(crate) use package::edit;
pub use package::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Snapshot, load};

pub const SHOW_DATA_TYPE_ICONS_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2023/showDataTypeIcons";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Target {
    Worksheet,
    CustomSheetView,
}

impl Target {
    pub(crate) const fn local_name(self) -> &'static [u8] {
        match self {
            Self::Worksheet => b"showDataTypeIcons",
            Self::CustomSheetView => b"showDataTypeIconsCustomSheetView",
        }
    }

    pub(crate) const fn display_name(self) -> &'static str {
        match self {
            Self::Worksheet => "showDataTypeIcons",
            Self::CustomSheetView => "showDataTypeIconsCustomSheetView",
        }
    }
}

pub(crate) const MAX_EXTENSION_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_PART_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_DEPTH: usize = 256;
pub(crate) const MAX_NODES: usize = 100_000;

pub(crate) fn invalid(value: impl Into<String>) -> Error {
    Error::Invalid(value.into())
}
