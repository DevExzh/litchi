//! Word document and range-level protection metadata and edit policy.
//!
//! The `protection` context models the bookmark-delimited editable ranges
//! described by [MS-DOC]. Usernames are never authenticated and decoding never
//! changes document content. Changed publication is enforced by default; an
//! explicit [`ProtectionAuthorization`] caller capability is required for a
//! bypass. The capability records caller metadata; it does not authenticate
//! a user or provide cryptographic audit.

mod codec;
mod model;
mod policy;

#[cfg(test)]
mod tests;

pub use model::{Mode, Range, Ranges, Reserved, Role, Selector, User};
pub(crate) use policy::classify;
pub use policy::{
    AuthorizationError, EditProtection, PackagePatch, ProtectionAuthorization, ProtectionPolicy,
};
