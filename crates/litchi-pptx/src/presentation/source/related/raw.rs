//! Explicit escape hatches for related-part identities and bytes.
//!
//! The ordinary related-part API intentionally exposes only semantic values.
//! Callers that need the package identity or original payload must opt into
//! this module explicitly.

use litchi_opc::PackURI;

/// Raw information retained by a related-part snapshot.
pub trait Metadata: sealed::Sealed {
    /// The package part name of the related part.
    fn raw_part_uri(&self) -> &PackURI;

    /// The relationship identifier used to select the related part.
    fn raw_relationship_id(&self) -> &str;

    /// The original part payload, before semantic projection.
    fn raw_data(&self) -> &[u8];
}

/// Return a related snapshot's package part name.
pub fn part_uri<T: Metadata + ?Sized>(value: &T) -> &PackURI {
    value.raw_part_uri()
}

/// Return a related snapshot's selecting relationship identifier.
pub fn relationship_id<T: Metadata + ?Sized>(value: &T) -> &str {
    value.raw_relationship_id()
}

/// Return a related snapshot's original payload.
pub fn data<T: Metadata + ?Sized>(value: &T) -> &[u8] {
    value.raw_data()
}

/// Return the slide's package part name.
pub fn slide_part_uri(slide: &super::super::SourceSlide) -> &PackURI {
    &slide.data.part_uri
}

/// Return the presentation relationship identifier that selected the slide.
pub fn slide_relationship_id(slide: &super::super::SourceSlide) -> &str {
    &slide.data.binding.presentation_relationship.id
}

impl Metadata for super::Layout {
    fn raw_part_uri(&self) -> &PackURI {
        &self.data.related.partname
    }

    fn raw_relationship_id(&self) -> &str {
        &self.data.related.relationship_id
    }

    fn raw_data(&self) -> &[u8] {
        self.data.related.data.as_bytes()
    }
}

impl Metadata for super::Notes {
    fn raw_part_uri(&self) -> &PackURI {
        &self.data.related.partname
    }

    fn raw_relationship_id(&self) -> &str {
        &self.data.related.relationship_id
    }

    fn raw_data(&self) -> &[u8] {
        self.data.related.data.as_bytes()
    }
}

impl Metadata for super::Comments {
    fn raw_part_uri(&self) -> &PackURI {
        &self.data.related.partname
    }

    fn raw_relationship_id(&self) -> &str {
        &self.data.related.relationship_id
    }

    fn raw_data(&self) -> &[u8] {
        self.data.related.data.as_bytes()
    }
}

/// Return the package part name of the selected legacy comment-author part,
/// when that part was acquired for the snapshot.
pub fn author_part_uri(value: &super::Comments) -> Option<&PackURI> {
    value
        .data
        .author_related
        .as_ref()
        .map(|part| &part.partname)
}

/// Return the relationship identifier that selected the legacy comment-author
/// part, when that part was acquired for the snapshot.
pub fn author_relationship_id(value: &super::Comments) -> Option<&str> {
    value
        .data
        .author_related
        .as_ref()
        .map(|part| part.relationship_id.as_ref())
}

/// Return the retained legacy comment-author part bytes, when acquired.
pub fn author_data(value: &super::Comments) -> Option<&[u8]> {
    value
        .data
        .author_related
        .as_ref()
        .map(|part| part.data.as_bytes())
}

/// Return the producer-supplied native author identifier for one comment.
pub fn comment_author_id(value: &super::Comment) -> u32 {
    value.author_id
}

/// Return the producer-supplied native comment index for one comment.
pub fn comment_index(value: &super::Comment) -> u32 {
    value.index
}

/// Short alias for [`comment_author_id`].
pub fn author_id(value: &super::Comment) -> u32 {
    comment_author_id(value)
}

/// Short alias for [`comment_index`].
pub fn index(value: &super::Comment) -> u32 {
    comment_index(value)
}

impl Metadata for super::Chart {
    fn raw_part_uri(&self) -> &PackURI {
        &self.data.related.partname
    }

    fn raw_relationship_id(&self) -> &str {
        &self.data.related.relationship_id
    }

    fn raw_data(&self) -> &[u8] {
        self.data.related.data.as_bytes()
    }
}

mod sealed {
    pub trait Sealed {}
    impl Sealed for super::super::Layout {}
    impl Sealed for super::super::Master {}
    impl Sealed for super::super::Notes {}
    impl Sealed for super::super::Comments {}
    impl Sealed for super::super::Chart {}
}

impl Metadata for super::Master {
    fn raw_part_uri(&self) -> &PackURI {
        &self.data.related.partname
    }
    fn raw_relationship_id(&self) -> &str {
        &self.data.related.relationship_id
    }
    fn raw_data(&self) -> &[u8] {
        self.data.related.data.as_bytes()
    }
}
