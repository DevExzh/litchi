//! Inert, lossless [MS-OSHARED] toolbar customization structures.
//!
//! The module exposes the common wire objects used by DOC, PPT, and XLS
//! without interpreting commands, macros, icons, or UI behavior.  Names are
//! intentionally contextual and prefix-free: [`Header`] is the `TB` structure,
//! while [`ControlHeader`] is `TBCHeader`.

mod codec;
mod control;
mod model;
mod patch;
mod snapshot;
mod transaction;
mod validation;

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "tests use concise assertions while exercising fallible malformed-input paths"
)]
mod tests;

// Keep the public surface contextual and prefix-free even though the model is
// split into semantic modules internally.
pub use self::control::{Body, Control};
pub use self::model::{
    ButtonFlags, ButtonState, ControlFlags, ControlHeader, ControlType, Data, Dimensions, Error,
    ExtraInfo, Flags, GeneralFlags, GeneralInfo, Header, HyperlinkType, MenuMerge, MergeMode,
    Restrictions, SpecificFlags, TextIcon, Type, WString,
};
pub use self::patch::{Change, Patch};
pub use self::snapshot::{Revision, Snapshot};
pub use self::transaction::{Commit, Transaction};
