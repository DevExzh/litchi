//! Package-independent values for the versioned modern-comment extensions.
//!
//! These values describe storage only. They do not model presence, identity,
//! synchronization, notification, or any other collaboration behavior.

mod changes;
mod extensions;
mod monikers;
mod reactions;
mod tasks;

pub use changes::{ChangeMetadata, CommentChange, CommentChanges, ReplyChange, ReplyChanges};
pub use extensions::{ExtensionEntry, ExtensionList, ExtensionPayload, OpaqueXml};
pub use monikers::{MonikerKind, MonikerList, MonikerNode};
pub use reactions::{Reaction, ReactionInstance, Reactions};
pub use tasks::{
    TaskAction, TaskAnchor, TaskAssign, TaskDetails, TaskEvent, TaskHistory, TaskSchedule,
    TaskTitle, TaskUndo, TaskUser,
};
