//! Shared semantic state for native drawable titles and captions.

/// Text attached to a drawable through iWork's native title and caption controls.
///
/// `None` means that the corresponding control is absent. An empty string is a
/// valid, present title or caption.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DrawableTitleCaption {
    /// Optional title displayed with the drawable.
    pub title: Option<String>,
    /// Optional caption displayed with the drawable.
    pub caption: Option<String>,
}
