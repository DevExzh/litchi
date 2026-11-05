//! Source-bound CRUD for the XLSX rich-value `refreshIntervals` extension.
//!
//! [MS-XLSX] defines `refreshIntervals` as a global element in the
//! `richvaluerefresh` namespace, but its `CT_RichValueType` reference and the
//! `rvTypesInfo` part schema place the owning metadata under a rich-value
//! type's existing SpreadsheetML `extLst`.  This module edits only that
//! `rdRichValuetypes` part.  It never scans worksheets, follows relationships,
//! contacts a service, or performs a refresh.
//!
//! A changed collection is authored only when its source type already has an
//! `x:ext` owner.  The extension URI is producer-owned metadata, so creating a
//! new extension container without one would require guessing an owner and
//! could misrepresent an unrelated opaque extension; that case is rejected
//! atomically with the source left unchanged.
//!
//! [`crate::Package::edit_rich_value_refresh`] starts an [`Edit`];
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! [`Patch`] and the [`Diagnostics`], and
//! [`crate::Package::apply_rich_value_refresh`] replays a patch. The source
//! package never changes.

mod codec;
mod model;
mod package;

use crate::error::Error as XlsxError;

pub use codec::{parse_refresh_intervals, write_refresh_intervals};
pub use model::{
    REFRESH_INTERVALS_NAMESPACE, RefreshInterval, RefreshIntervals, TypeRefreshIntervals,
};
pub use package::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Snapshot, load};

pub(crate) const MAX_TYPES: usize = super::MAX_ITEMS;
pub(crate) const MAX_INTERVALS: usize = super::MAX_ITEMS;

pub(crate) fn limit(kind: LimitKind, observed: usize, maximum: usize) -> XlsxError {
    XlsxError::limit(kind, observed, maximum)
}

/// XML 1.0 fifth-edition `Char` production.
pub(crate) const fn valid_xml10(value: char) -> bool {
    matches!(value as u32, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}
