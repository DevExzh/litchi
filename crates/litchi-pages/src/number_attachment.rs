//! Page-number and page-count attachments of the body and of header/footer
//! slots.

pub use crate::package::body_number_attachments::{
    Attachment, Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal, Target,
};
pub use litchi_iwa_text::number_attachment::{
    TextNumberAttachmentFormat, TextNumberAttachmentKind, TextNumberAttachmentSettings,
    TextNumberAttachmentText,
};
