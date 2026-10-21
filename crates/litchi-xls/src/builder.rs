//! Fresh XLS workbook authoring.
//!
//! [`Builder::build`](crate::builder::Builder::build) returns an immutable [`crate::Package`]. The package owns
//! memory, stream and atomic path publication.
//!
//! Drafts reached from authoring live here without a prefix (record 0351
//! §2.1 rule 7): [`Chart`](crate::builder::Chart) is the BIFF8 chart draft that the `chart` family
//! (`Package::edit_charts`) places and returns.
//!
//! ```
//! let mut builder = litchi_xls::Builder::new();
//! let sheet = builder.add_worksheet("Data")?;
//! builder.write_number(sheet, 0, 0, 42.0)?;
//! let package = builder.build()?;
//! let bytes = package.to_bytes()?;
//! assert!(!bytes.is_empty());
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub use crate::chart::model::Chart;
pub use crate::writer::Builder;
