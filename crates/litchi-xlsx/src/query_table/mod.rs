//! Typed SpreadsheetML query-table metadata.
//!
//! The owner is layered by responsibility: semantic declarations in
//! `model`, bounded XML/MCE conversion in `codec`, and OPC relationship
//! operations in `package`.

mod codec;
mod model;
mod package;

pub use codec::{parse_query_table, write_query_table};
pub use model::{
    Conformance, ExtensionAttribute, ExtensionList, Field, GrowShrinkType, IconSet, Part, Refresh,
    SortBy, SortCondition, SortMethod, SortState, Table,
};
pub use package::{
    QUERY_TABLE_CONTENT_TYPE, QUERY_TABLE_RELATIONSHIP_TYPE, STRICT_QUERY_TABLE_RELATIONSHIP_TYPE,
    add_worksheet_query_table, find_worksheet_query_table, is_query_table_relationship_type,
    load_worksheet_query_tables, remove_worksheet_query_table, reorder_worksheet_query_tables,
    replace_worksheet_query_table, update_worksheet_query_table,
};
