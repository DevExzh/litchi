//! Strict PowerPoint 10 presentation-comment support.
//!
//! The public model is kept separate from the record codec. The codec accepts
//! only the ordered, bounded records defined by [MS-PPT], while the model
//! exposes contextual author metadata to callers. Unknown records remain in
//! the surrounding `Record` tree and are not executed or resolved.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub(crate) use codec::parse_slide_comments;
pub use model::{Author, Authors};
