//! Worksheet-level package assembly for list objects.
//!
//! Collection state, future-record retention, record emission, and package
//! validation live in focused submodules. This facade keeps the crate-visible
//! owner names stable for worksheet readers and writers.

mod collector;
mod future;
mod records;
mod validation;

pub(crate) use collector::ListObjectCollector;
pub(crate) use records::feature_header_record;
