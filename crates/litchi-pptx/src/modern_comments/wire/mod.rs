//! Bounded XML wire codecs for the versioned modern-comment families.

mod changes;
mod extensions;
mod monikers;
mod reactions;
mod tasks;
mod xml;

pub(crate) use changes::{collect_change_commands, replace_change_commands};
pub(crate) use extensions::{parse_extensions, write_extensions};
