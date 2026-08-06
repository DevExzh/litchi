//! Layered WordprocessingML paragraph and run models.
//!
//! The semantic XML-backed values live in [`model`], streaming WordprocessingML
//! parsing lives in [`codec`], and relationship-backed hyperlink resolution
//! lives in [`package`]. This module remains the historical `crate::paragraph`
//! facade.

mod codec;
pub mod collapsed;
pub mod extensions;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use crate::run_effects;
pub use crate::run_effects::Effects;
pub use crate::run_symbols::{Symbol, Symbols};
pub use collapsed::{Collapsed, Commit, Patch, Snapshot, Transaction};
pub use extensions::{Extensions, Id, Ids};
pub use model::{
    LineSpacingRule, Paragraph, ParagraphSpacing, Run, RunBreak, RunBreakClear, RunBreakType,
    RunProperties, RunUnderline, RunUnderlineColor,
};

pub(crate) use codec::{extract_word_text, is_fragment_word_name};
