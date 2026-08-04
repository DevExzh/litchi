//! Numbers semantic value models.
//!
//! Archive parsing, protobuf decoding, and package mutation remain owned by
//! the Numbers implementation. This crate starts the downward migration with
//! the dependency-free cell vocabulary used by Numbers, Pages table editing,
//! and the shared structured extractor.

#![forbid(unsafe_code)]

/// Cell-level Numbers vocabulary.
pub mod cell;
/// Semantic sheet containers.
pub mod sheet;
/// Sparse semantic table vocabulary.
pub mod table;

pub use sheet::{Builder as SheetBuilder, Sheet};
pub use table::{
    Builder as TableBuilder, Cell, Dimensions, Error as TableError, Grid, GridBudget, InsertError,
    InsertResult, Position, Range, Table, View,
};
