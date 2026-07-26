//! Native Numbers row/column identity-map constants.

/// Modern `TST.ColumnRowUIDMapArchive` message type.
///
/// Numbers rewrites the legacy `6200` representation on first open and assigns
/// new axis UUIDs. Emitting the current type preserves identities referenced by
/// hidden-state, formula, merge, and other table topology graphs.
pub(crate) const COLUMN_ROW_UID_MAP_MESSAGE_TYPE: u32 = 6_267;
