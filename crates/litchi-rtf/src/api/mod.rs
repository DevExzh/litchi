//! Ordinary immutable document API.

mod facade;
pub(crate) mod story;

pub use facade::Document;
pub use story::{
    Break, Format, Inline, Inlines, Paragraph, ParagraphFormat, Paragraphs, Run, Runs, Story,
};
