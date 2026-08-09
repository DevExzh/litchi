//! Typed anchors and opaque payloads for `WordprocessingML` alternative-format imports.
//!
//! Payloads are deliberately never parsed, executed, fetched, or opened as nested
//! packages. The supported authoring media types follow the Microsoft Word notes
//! in `[MS-OI29500]` §2.1.527 and `[MS-OE376]` §2.1.558.
//!
//! ```
//! use litchi_docx::alt::{Data, Import};
//!
//! let embedded = Import::data(Data::Html(b"<p>opaque</p>".to_vec()));
//! let linked = Import::link("https://example.invalid/import.html")?;
//! # let _ = (embedded, linked);
//! # Ok::<(), litchi_docx::Error>(())
//! ```

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use codec::{active, is_relationship, scan};
pub use model::{
    Chunk, Conformance, Data, Import, Kind, MAX_CHUNKS, MAX_DATA_BYTES, MAX_XML_BYTES,
    MAX_XML_DEPTH, Rel, Uri,
};
pub use package::{Part, Target};
