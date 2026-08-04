//! Layered XLSB styles-part model and codec.
//!
//! The model contains package-neutral values and the codec owns the checked-in
//! Brt* record layouts. Host adapters may convert neutral alignment and border
//! values to legacy public types.

mod codec;
mod model;

pub use codec::{
    Error, Result, parse_border, parse_cell_format, parse_direct_color, parse_fill, parse_font,
    parse_num_fmt, read,
};
pub use model::{
    Alignment, Border, BorderSide, BorderStyle, CellFormat, Fill, Font, HorizontalAlignment,
    NumberFormat, Table, VerticalAlignment,
};
