//! OpenDocument Chart support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod package;

pub use facade::{Builder, Chart};
pub use litchi_odf_common::chart;
