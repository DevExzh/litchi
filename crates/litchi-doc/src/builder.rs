//! Fresh legacy Word package authoring.
//!
//! [`Builder`] stages semantic DOC content and [`Builder::build`] publishes a
//! validated [`crate::Package`].  Publication and byte output belong to the
//! package owner; the builder deliberately has no save or stream methods.

//! ```
//! let mut builder = litchi_doc::Builder::new();
//! builder.add_paragraph("Hello from DOC")?;
//! let mut package = builder.build()?;
//! let bytes = package.to_bytes()?;
//! assert!(!bytes.is_empty());
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub use crate::writer::Builder;
