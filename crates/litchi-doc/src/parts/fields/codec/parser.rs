//! Field instruction grammar and typed parser parts.

mod buttons;
mod external;
mod formatting;
mod indexing;
mod limits;
mod mail_merge;
mod merge;
mod metadata;
mod numbering;
mod parts;
mod prelude;
mod references;
mod syntax;

// Preserve the codec's crate-internal parser facade while keeping each semantic
// field family in its own bounded parser owner.
pub(in crate::parts::fields) use buttons::*;
pub(in crate::parts::fields) use external::*;
pub(in crate::parts::fields) use formatting::*;
pub(in crate::parts::fields) use indexing::*;
#[allow(unused_imports)]
pub(in crate::parts::fields) use limits::*;
pub(in crate::parts::fields) use mail_merge::*;
pub(in crate::parts::fields) use merge::*;
pub(in crate::parts::fields) use metadata::*;
pub(in crate::parts::fields) use numbering::*;
#[allow(unused_imports)]
pub(in crate::parts::fields) use parts::*;
pub(in crate::parts::fields) use references::*;
