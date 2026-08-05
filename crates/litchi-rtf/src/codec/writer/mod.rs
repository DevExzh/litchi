//! RTF writer owner.
//!
//! Document serialization lives in [`codec`], while focused writer regression
//! coverage lives in [`tests`].

mod codec;

#[cfg(test)]
mod tests;

pub use codec::{
    Charset, DEFAULT_TAB_WIDTH_TWIPS, DefaultTabWidthPolicy, MAX_DEFAULT_TAB_WIDTH_TWIPS,
    RtfWriter, WriterOptions,
};
