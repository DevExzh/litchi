//! Strict, inert metadata for binary PowerPoint headers and footers.
//!
//! This module implements [MS-PPT] sections 2.4.15 and 2.5.16. It parses and
//! serializes only the relevant record family; it does not format dates, modify
//! an OLE compound file, or activate presentation content.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{
    DateTimeFormatId, HeaderFooter, HeaderFooterDisplayText, HeaderFooterOptions,
    HeaderFooterParent, HeaderFooterParentOrdinal, HeaderFooterScope, HeaderFooters,
    ScopedHeaderFooterDisplayText,
};
