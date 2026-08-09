//! Layered XML codec for `DrawingML` charts.
//!
//! The public façade remains [`super::read`]. Input handling is isolated in
//! [`xml`], typed chart construction in [`semantic`], and value constraints in
//! [`validation`].

mod semantic;
mod validation;
mod xml;

pub use xml::read;
