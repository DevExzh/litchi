//! ODF spreadsheet database ranges, filters, sorting, and subtotals.

mod codec;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

pub use semantic::{
    Condition, ConditionSource, DataType, EmbeddedNumberBehavior, Expression, Field, Filter, Key,
    Order, Orientation, Range, Rule, Rules, Sort, SortGroups, Source,
};

pub use codec::{
    parse_database_ranges, parse_filter, parse_source_query, parse_source_sql, parse_source_table,
    write_database_range_fragment, write_database_ranges, write_database_source, write_filter,
};
pub use validation::{validate_database_range_collection, validate_filter};
