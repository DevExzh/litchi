//! Inert, lossless [MS-OSHARED] toolbar customization structures.
//!
//! The module exposes the common wire objects used by DOC, PPT, and XLS
//! without interpreting commands, macros, icons, or UI behavior.  Names are
//! intentionally contextual and prefix-free: [`Header`] is the `TB` structure,
//! while [`ControlHeader`] is `TBCHeader`.

mod codec;
mod model;

#[cfg(test)]
mod tests;

// Keep the public surface contextual and prefix-free even though the model is
// split into semantic modules internally.
pub use self::model::{
    ButtonFlags, ButtonState, ControlFlags, ControlHeader, ControlType, Data, Dimensions, Error,
    ExtraInfo, Flags, GeneralFlags, GeneralInfo, Header, HyperlinkType, MenuMerge, MergeMode,
    Restrictions, SpecificFlags, TextIcon, Type, WString,
};
