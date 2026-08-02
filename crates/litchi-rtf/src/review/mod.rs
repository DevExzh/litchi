//! Comments, notes, permissions, protection, and tracked revisions.

pub(crate) mod annotation;
pub(crate) mod bookmark;
pub(crate) mod editable_region;
pub(crate) mod note_options;
pub(crate) mod note_separator;
pub(crate) mod protection_range;
pub(crate) mod protection_user;
pub(crate) mod review_display;
pub(crate) mod revision_save;

pub use crate::section::Note;
pub use annotation::{
    Annotation, AnnotationType as AnnotationKind, Revision, RevisionAuthor,
    RevisionType as RevisionKind,
};
