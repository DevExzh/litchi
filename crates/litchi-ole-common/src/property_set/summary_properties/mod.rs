//! Document properties authored into a new compound file.
//!
//! The legacy writers (DOC, PPT and XLS) create files from scratch, so they
//! have no source `\005SummaryInformation` or `\005DocumentSummaryInformation`
//! stream to edit. [`SummaryProperties`] is the property state such a writer
//! serializes: the typed [MS-OLEPS] 2.25.1 `SummaryInformation` section, the
//! typed [MS-OSHARED] 2.3.3.2.2 `DocumentSummaryInformation` section and the
//! optional [MS-OSHARED] 2.3.3.2.3 `UserDefinedProperties` section, all held
//! by the same owners the property-set transactions use
//! ([`summary_information::Snapshot`](super::summary_information::Snapshot),
//! [`document_summary::Snapshot`](super::document_summary::Snapshot) and a
//! generic [`Section`](super::Section)). A file the writer creates therefore
//! reads back through the existing readers, and its streams can be edited by
//! the existing property-set transactions afterwards.
//!
//! Edits mirror the transactions: [`SummaryProperties::edit_summary_information`],
//! [`SummaryProperties::edit_document_summary_information`] and
//! [`SummaryProperties::edit_user_defined_properties`] take the same closures,
//! validate the whole section before anything changes, and leave the value
//! untouched when they fail. [`SummaryProperties::set`] is the by-name entry
//! the writers' `set_property` uses; [`SummaryProperties::set_user_defined`]
//! sets one typed custom property.
//!
//! # Names accepted by [`SummaryProperties::set`]
//!
//! Names compare ASCII-case-insensitively with ASCII spaces ignored, so both
//! Office's built-in property names (`"Last Author"`) and the unified
//! metadata keys (`"lastModifiedBy"`) are accepted:
//!
//! | Names | Property |
//! |---|---|
//! | `Title` | PIDSI `TITLE` |
//! | `Subject` | PIDSI `SUBJECT` |
//! | `Author` | PIDSI `AUTHOR` |
//! | `Keywords` | PIDSI `KEYWORDS` |
//! | `Comments`, `Description` | PIDSI `COMMENTS` |
//! | `Template` | PIDSI `TEMPLATE` |
//! | `Last Author`, `Last Modified By` | PIDSI `LASTAUTHOR` |
//! | `Revision Number` | PIDSI `REVNUMBER` (a nonnegative whole number) |
//! | `Application Name` | PIDSI `APPNAME` |
//! | `Creation Date`, `Created` | PIDSI `CREATE_DTM` (RFC 3339 timestamp) |
//! | `Last Save Time`, `Modified` | PIDSI `LASTSAVE_DTM` (RFC 3339 timestamp) |
//! | `Last Print Date`, `Last Printed` | PIDSI `LASTPRINTED` (RFC 3339 timestamp) |
//! | `Category` | PIDDSI `CATEGORY` |
//! | `Format`, `Presentation Format` | PIDDSI `PRESFORMAT` |
//! | `Manager` | PIDDSI `MANAGER` |
//! | `Company` | PIDDSI `COMPANY` |
//! | `Content Type` | PIDDSI `CONTENTTYPE` |
//! | `Content Status` | PIDDSI `CONTENTSTATUS` |
//! | `Language` | PIDDSI `LANGUAGE` |
//! | `Document Version` | PIDDSI `DOCVERSION` |
//!
//! Office's other built-in names (the counts, `Security`, `Total Editing
//! Time`, `Hyperlink base`) are refused: they are not text, and the typed
//! edits set them. A name that begins with `_` is refused too, because
//! [MS-OSHARED] 2.3.3.2.3.2 reserves such names for system properties with
//! fixed types; [`SummaryProperties::set_user_defined`] sets them with their
//! type. Every other name sets a text property in the user-defined section.
//!
//! # Code pages
//!
//! Every section starts with the Windows-1252 code page, as Office writes
//! for Western text. A text value or property name that code page cannot
//! represent switches its section to UTF-16 (`CP_WINUNICODE`) before it is
//! stored, so [`SummaryProperties::set`] and
//! [`SummaryProperties::set_user_defined`] never refuse text for its
//! characters. The typed edit closures keep the caller's code page: they
//! refuse unrepresentable text unless the closure changes the code page
//! itself.
//!
//! # Determinism
//!
//! Nothing here reads the clock or the environment: properties are written
//! in the order they were first set, and a creation or save time is written
//! only when the caller sets one.

mod codec;
mod metadata;
mod model;
mod names;
mod source;

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "tests use concise assertions while exercising fallible authoring paths"
)]
mod tests;

pub use codec::{SYSTEM_IDENTIFIER, Streams};
pub use metadata::MetadataError;
pub use model::SummaryProperties;
