//! Explicit package-local metadata of DrawingML text.
use super::{Field, Hyperlink, Picture};

/// Native field identifier as authored.
#[must_use]
pub fn field_id(field: &Field) -> &str {
    &field.id
}

/// Native relationship referenced by a hyperlink.
#[must_use]
pub fn hyperlink_relationship_id(link: &Hyperlink) -> Option<&str> {
    link.relationship_id.as_deref()
}

/// Native relationship referenced by a picture bullet.
#[must_use]
pub fn picture_relationship_id(picture: &Picture) -> Option<&str> {
    picture.relationship_id.as_deref()
}
