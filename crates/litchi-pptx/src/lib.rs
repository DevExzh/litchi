//! Typed PowerPoint Open XML documents.
//!
//! The crate is being extracted one semantic capability at a time. The
//! backgrounds module owns package-independent slide-background values and
//! its XML fill codec. The transition module owns slide-transition values and
//! its bounded codec. The
//! laser module owns inert laser-trace values and their bounded codec. The
//! font module owns embedded-font values and atomic package CRUD. The tag
//! module owns inert programmable tag lists and package CRUD. The notes module
//! owns bounded speaker-notes graphs, text encoding, and transactional package
//! mutation.
//! [`table::style`] owns typed table-style catalogs and their package graph.

#![forbid(unsafe_code)]

pub mod animations;
pub mod backgrounds;
pub mod comments;
mod error;
pub mod font;
pub mod format;
pub mod hyperlinks;
pub mod laser;
pub mod notes;
pub mod shape;
pub mod table;
pub mod tag;
pub mod time;
pub mod transition;

pub use animations::*;
pub use backgrounds::{GradientStop, GradientType, PatternType, PictureStyle, SlideBackground};
pub use comments::{
    PresentationComment, PresentationCommentAuthor, PresentationCommentConformance,
    PresentationComments, SlideCommentList, add_presentation_comment,
    add_presentation_comment_author, find_presentation_comment, find_presentation_comment_author,
    load_presentation_comments, parse_comment_authors, parse_slide_comments,
    remove_presentation_comment, remove_presentation_comment_author,
    reorder_presentation_comment_authors, reorder_presentation_comments,
    replace_presentation_comment, replace_presentation_comment_author, store_presentation_comments,
    update_presentation_comment, update_presentation_comment_author, write_comment_authors,
    write_slide_comments,
};
pub use error::{Error, Result};
pub use format::{ImageFormat, TextFormat};
pub use hyperlinks::Hyperlink;
