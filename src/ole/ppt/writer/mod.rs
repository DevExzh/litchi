//! PPT file writing module
//!
//! This module provides comprehensive support for creating and modifying
//! Microsoft PowerPoint presentations in the legacy binary format (.ppt files).
//!
//! # Features
//!
//! - **Shapes**: Rectangles, ellipses, lines, arrows, and more
//! - **Text formatting**: Bold, italic, underline, font sizes, colors
//! - **Shape styling**: Fill colors, gradients, line styles, shadows
//! - **Pictures**: JPEG, PNG, and other image format support
//! - **Hyperlinks**: URL and slide navigation links
//! - **Notes**: Full speaker notes support

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

/// BLIP (picture) support
pub mod blip;

/// Text formatting support (bold, italic, colors, fonts)
pub mod text_format;

/// Shape styling (fill, line, shadow)
pub mod shape_style;

/// Extended shape types (lines, ellipses, arrows, etc.)
pub mod shapes;

/// Hyperlink support
pub mod hyperlink;

/// Notes slide support
pub mod notes;

/// Sound collection for animations
mod sound_collection;

// Re-export public types from core
pub use core::{PptWriteError, PptWriter, ShapeProperties, ShapeType, TextAlignment};

// Re-export commonly used types from submodules
pub use blip::{BlipStoreBuilder, BlipType, PictureData};
pub use escher::{EscherBuilder, create_dgg_container, create_shape_container};
pub use hyperlink::{Hyperlink, HyperlinkCollection, HyperlinkTarget};
pub use notes::{NotesCollection, NotesPage};
pub use persist::{PersistPtrBuilder, UserEditAtom};
pub use records::{RecordBuilder, RecordHeader};
pub use shape_style::{FillStyle, LineStyleConfig, ShadowStyle, ShapeColor, ShapeStyle};
pub use shapes::{Shape, ShapeCollection, ShapeKind};
pub use sound_collection::build_sound_collection;
pub use text_format::{FontEntity, FontStyle, Paragraph, TextAlign, TextColor, TextRun};
