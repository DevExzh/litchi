//! Typed inert access to the MS-OSHARED hyperlink properties stored in a
//! `UserDefinedProperties` section.
//!
//! `_PID_LINKBASE` and `_PID_HLINKS` are names in the user-defined dictionary,
//! not PIDDSI 0x14 or 0x15. Their target and location strings remain opaque
//! text: this module never resolves, opens, or executes them.
//!
//! The generic property codec owns the `VT_BLOB` envelope, including
//! `VtHyperlinkValue.cbData` (MS-OSHARED 2.3.3.1.20-21). The reserved overlay
//! receives/emits only `VecVtHyperlink`, starting with `cElements`; it must
//! not add another byte-count word inside `Value::Blob`.

mod codec;
mod model;

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::drop_non_drop,
    reason = "tests use concise assertions while exercising fallible malformed-input paths"
)]
mod tests;

pub use model::{Edit, Hyperlink, Hyperlinks, Limits, LimitsBuilder, LinkBase, Properties};

/// Case-insensitive reserved user-defined property name for a relative-link base.
pub const LINK_BASE: &str = "_PID_LINKBASE";
/// Case-insensitive reserved user-defined property name for the inert hyperlink list.
pub const HYPERLINKS: &str = "_PID_HLINKS";
