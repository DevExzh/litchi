//! RTF control-word vocabulary and specification dispatch.

mod dispatch;
mod model;

pub(super) use dispatch::match_control_word;
pub use model::ControlWord;
