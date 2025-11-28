//! PPT file writing module
//!
//! This module provides comprehensive support for creating and modifying
//! Microsoft PowerPoint presentations in the legacy binary format (.ppt files).

/// Core PPT writer implementation
mod core;

/// PPT record generation system
pub mod records;

/// Escher (Office Drawing) records
pub mod escher;

/// PersistPtr offset mapping
pub mod persist;

/// Atom record builders
pub mod atoms;

/// MS-PPT specification types and constants
pub mod spec;

/// TxMasterStyleAtom data constants
pub mod tx_style;

/// Environment container data constants
pub mod env_data;

/// Master slide PPDrawing types and constants
pub mod master_drawing;

// Re-export public types
pub use core::{PptWriteError, PptWriter, ShapeProperties, ShapeType, TextAlignment};
pub use escher::{EscherBuilder, create_dgg_container, create_shape_container};
pub use persist::{PersistPtrBuilder, UserEditAtom};
pub use records::{RecordBuilder, RecordHeader};
