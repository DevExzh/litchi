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

// Migration-only OfficeArt encoding internals. The typed public substrate is
// `litchi-odraw`; numeric record builders are not part of the writer facade.
mod escher;

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

/// Slide comment support
pub mod comments;

/// Custom slide show (named show) support
pub mod custom_shows;

/// Per-slide timing support
pub mod slide_timing;

/// Table authoring (OfficeArt table groups with text cells)
pub mod table;

/// Native chart request model; binary emission is currently refused safely.
pub mod chart;

/// PowerPoint 11 smart-tag authoring.
pub mod smart_tags;

// Re-export public types from core
pub use crate::encryption::EncryptionProfile;
pub use crate::view_info::{
    Guide, GuideOrientation, Ratio, SlideViewInfo, SlideViewPreferences, ViewKind, ViewOrigin,
    ZoomViewInfo,
};
pub use core::{ShapeProperties, ShapeType, TextAlignment, WriteError, Writer};

// Re-export commonly used types from submodules
pub use crate::shapes::geometry::{GeometryRect, ShapePathType};
pub use blip::{Id as PictureId, Kind as PictureKind, PictureData, Pictures};
pub use chart::{Chart, ChartKind, ChartSeries};
pub use comments::{CommentDateTime, SlideComment};
pub use custom_shows::CustomShow;
pub use escher::FreeformGeometry;
pub use hyperlink::{Hyperlink, HyperlinkCollection, HyperlinkTarget};
pub use notes::{NotesCollection, NotesPage};
pub use persist::{PersistPtrBuilder, UserEditAtom};
pub use records::{RecordBuilder, RecordHeader};
pub use shape_style::{FillStyle, LineStyleConfig, ShadowStyle, ShapeColor, ShapeStyle};
pub use shapes::{Shape, ShapeCollection, ShapeKind};
pub use slide_timing::SlideTiming;
pub use smart_tags::{SmartTagDefinition, SmartTagIndex};
pub use sound_collection::build_sound_collection;
pub use table::{DEFAULT_COLUMN_WIDTH_PT, DEFAULT_ROW_HEIGHT_PT, MAX_TABLE_DIMENSION, Table};
pub use text_format::{
    FontEntity, FontStyle, Paragraph, TabAlign, TabStop, TextAlign, TextColor, TextDirection,
    TextFontAlign, TextRun,
};
