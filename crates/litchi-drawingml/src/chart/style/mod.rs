//! Contextual codecs and values for `DrawingML` chart-style companion XML.
//!
//! The `chartStyle` and `colorStyle` parts are XML-level `DrawingML`
//! capabilities. Their OPC content types, relationships, and physical part
//! locations remain owned by the format crate that contains them.
//!
//! Specification basis: `[MS-ODRAWXML]` §§2.6 (Chart), 2.8 (chartStyle), and
//! 2.24 (`ChartEx`). `[MS-OI29500]` package and relationship rules are applied by
//! the owning format crate, as are host-specific `[MS-PPTX]` and `[MS-XLSX]`
//! resource and relationship constraints.

pub mod codec;
pub mod model;

pub use codec::{parse, parse_color};
pub use model::{
    Color, ColorDocument, ColorInfo, ColorKind, ColorMethod, ColorValue, Document, Entry,
    EntryKind, FontIndex, FontReference, Info, MarkerLayout, MarkerSymbol, Payload, Reference,
    Transform, TransformKind, Variation,
};
