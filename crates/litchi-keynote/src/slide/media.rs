//! Keynote slide media classifications.

/// The semantic role of a movie drawable owned directly by a Keynote slide.
///
/// This value deliberately contains no archive, package, or native identifier
/// state. The IWA adapter owns the graph and media records and uses this type
/// only for the product-level classification exposed in movie information.
#[repr(u8)]
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MovieKind {
    /// An ordinary file-backed movie inserted by the user.
    File,
    /// An independently positioned audio clip stored in a movie archive.
    Audio,
    /// A file-backed replacement target materialized from a slide layout.
    Placeholder,
    /// A camera-backed live-video drawable.
    LiveVideo,
}

#[cfg(test)]
mod tests {
    use super::MovieKind;
    use std::mem::size_of;

    #[test]
    fn classification_is_a_compact_copyable_value() {
        assert_eq!(size_of::<MovieKind>(), 1);
        assert_ne!(MovieKind::File, MovieKind::Audio);
        assert_ne!(MovieKind::Placeholder, MovieKind::LiveVideo);
    }
}
