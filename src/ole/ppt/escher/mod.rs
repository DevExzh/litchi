//! Escher (Office Drawing) record parsing module.
//!
//! Escher is Microsoft's drawing layer format used across Office applications.
//! It contains shapes, text boxes, images, and other graphical elements.
//!
//! # Architecture
//!
//! - Zero-copy parsing with lifetime-based borrowing
//! - Iterator-based container traversal
//! - Lazy shape evaluation
//! - Minimal allocations
//!
//! # Performance
//!
//! - Direct byte slice access (no intermediate buffers)
//! - Pre-allocated capacity estimation
//! - Functional iterator chains
//! - Efficient bit manipulation

pub mod types;
pub mod record;
pub mod parser;
pub mod container;
pub mod text;
pub mod shape;
pub mod shape_factory;
pub mod properties;

pub use types::EscherRecordType;
pub use record::EscherRecord;
pub use parser::EscherParser;
pub use container::EscherContainer;
pub use text::extract_text_from_escher;
pub use shape::{EscherShape, EscherShapeType};
pub use shape_factory::EscherShapeFactory;
pub use properties::{EscherProperties, EscherPropertyId, ShapeAnchor};

