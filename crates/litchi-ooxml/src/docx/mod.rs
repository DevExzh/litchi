//! Canonical WordprocessingML facade.
//!
//! DOCX package-graph orchestration, document models, codecs, and writer APIs
//! are owned by [`litchi_docx`]. This module only exposes that standalone
//! surface to the umbrella host while the other OOXML families migrate.
//!
//! ```rust,no_run
//! use litchi_docx::{Cell, Document, Element, Package, Paragraph, Row, Run, Table};
//!
//! let package = Package::open("document.docx")?;
//! let document: Document<'_> = package.document()?;
//! for element in document.elements()? {
//!     match element {
//!         Element::Paragraph(paragraph) => {
//!             let _: &Paragraph = &paragraph;
//!             for run in paragraph.runs()? {
//!                 let _: &Run = &run;
//!             }
//!         }
//!         Element::Table(table) => {
//!             let _: &Table = &table;
//!             for row in table.rows()? {
//!                 let _: &Row = &row;
//!                 for cell in row.cells()? {
//!                     let _: &Cell = &cell;
//!                 }
//!             }
//!         }
//!     }
//! }
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

pub use litchi_docx::*;
pub use litchi_drawingml::diagram::{DiagramNode, DiagramType, SmartArt, SmartArtBuilder};
pub use litchi_drawingml::geom::{Preset, TextPreset};
pub use litchi_opc::FontEmbedding;
