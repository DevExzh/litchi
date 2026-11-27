//! BIFF8 `QUERYTABLE` record sequence (MS-XLS 2.1.7.20.5): typed, inert
//! reading of worksheet query tables and their external data connections.
//!
//! A query table is described by the record sequence
//! `Qsi DBQUERY QsiSXTag DBQUERYEXT [SXADDLQSI] [QSIR] [SORTDATA12]`.
//! Connection strings, SQL command text, Web query URLs, post data, and
//! text-source file paths are stored verbatim and are never opened, resolved,
//! contacted, refreshed, or executed.
//!
//! The owner is intentionally layered: [`model`] contains the typed public
//! representation, [`codec`] decodes individual BIFF records, and
//! [`validation`] assembles the ordered sequence while enforcing its inert,
//! bounded parsing rules.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::*;

pub(crate) use validation::QueryTableCollector;
