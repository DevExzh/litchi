//! Contextual semantic helpers shared by the web model and package layers.

mod constants;

/// Protocol constants and defaults remain private to the web owner while
/// semantic submodules consume them through this facade.
#[allow(unused_imports)]
pub(in crate::web) use constants::*;
