//! RTF control-word vocabulary and specification dispatch.

mod dispatch;
mod model;
mod repair;

pub(super) use dispatch::match_control_word;
pub(crate) use model::ControlWord;
