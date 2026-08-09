//! Detached construction for new family packages.

mod builder;

pub use builder::Builder;
pub(crate) use builder::{
    render_forms, render_fragment, render_inline, render_metadata, render_styles,
};
