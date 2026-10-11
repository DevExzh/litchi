//! Fresh XLS workbook authoring.
//!
//! [`Builder::build`] returns an immutable [`crate::Package`]. The package owns
//! memory, stream and atomic path publication.
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

pub use crate::writer::Builder;
