//! Layered `WordprocessingML` header and footer stories.
//!
//! The owner keeps the semantic story view, the bounded XML/MCE codec, and
//! package relationship resolution separate. `Story` retains the original
//! XML allocation for lossless access while semantic traversal uses one
//! cached, bounded MCE view when required.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::{Kind, MAX_XML_BYTES, MAX_XML_DEPTH, MAX_XML_NODES, Role, Story};

pub(crate) use package::{footers as load_footers, headers as load_headers, image_watermarks};
