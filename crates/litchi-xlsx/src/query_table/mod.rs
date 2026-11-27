//! Typed `SpreadsheetML` query-table metadata.
//!
//! The owner is layered by responsibility: semantic declarations in
//! `model`, bounded XML/MCE conversion in `codec`, and OPC relationship
//! operations in `package`.

mod codec;
mod edit;
mod model;
mod package;

pub use codec::{parse_query_table, write_query_table};
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};
pub use model::{
    Conformance, ExtensionAttribute, ExtensionList, Field, GrowShrinkType, IconSet, Part, Refresh,
    SortBy, SortCondition, SortMethod, SortState, Table,
};
pub use package::{
    QUERY_TABLE_CONTENT_TYPE, QUERY_TABLE_RELATIONSHIP_TYPE, STRICT_QUERY_TABLE_RELATIONSHIP_TYPE,
    find_worksheet_query_table, is_query_table_relationship_type, load_worksheet_query_tables,
};
