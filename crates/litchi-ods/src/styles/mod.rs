//! ODS style vocabulary and layered style codecs.

pub mod table_template;

pub use table_template::{Axis, Region, Style, Template, parse, parse_parts};
