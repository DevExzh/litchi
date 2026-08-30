//! Archive-free Numbers table appearance semantics.
//!
//! [`crate::table::appearance::Appearance`] is the complete effective
//! appearance value for one table.
//! It contains only typed style settings; native style objects, inheritance
//! graphs, object identifiers, protobuf messages, and package bytes remain
//! private to the Numbers package adapter.
//!
//! Read and edit a rooted table through [`crate::Package::table_appearance`]
//! and [`crate::Package::edit_table_appearance`]. The exact-source
//! transaction types are available below
//! [`crate::table::appearance::transaction`]. A staged edit uses
//! [`crate::table::appearance::transaction::Edit::set`] to replace the
//! complete appearance value and
//! [`crate::table::appearance::transaction::Edit::commit`] to publish it.
//! Apply a retained exact-source patch with
//! [`crate::Package::apply_table_appearance`].

/// Exact-source transactions for one rooted Numbers table's appearance.
pub mod transaction {
    pub use crate::package::table_appearance::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
    };
}

pub use litchi_iwa_common::table::appearance::{
    Appearance, Banding, GridlineVisibility, Gridlines, RowSizing,
};

#[cfg(test)]
mod tests {
    use super::{Appearance, Banding, GridlineVisibility, Gridlines, RowSizing};
    use litchi_iwa_common::table::appearance::Appearance as CommonAppearance;

    #[test]
    fn appearance_is_the_shared_archive_free_value() {
        fn accepts_common_appearance(_: CommonAppearance) {}

        accepts_common_appearance(Appearance::default());
        assert_eq!(Banding::default(), Banding::Disabled);
        assert_eq!(RowSizing::default(), RowSizing::Fixed);
        assert_eq!(GridlineVisibility::default(), GridlineVisibility::Visible);
        assert_eq!(Gridlines::default(), Appearance::default().gridlines);
    }
}
