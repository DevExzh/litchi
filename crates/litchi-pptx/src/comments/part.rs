//! Borrowed low-level views for legacy PresentationML comment parts.

use litchi_opc::part::Part as OpcPart;

use super::{AUTHORS_CONTENT_TYPE, Author, COMMENTS_CONTENT_TYPE, Comment, Conformance};
use crate::Result;
use crate::parts::validate_content_type;

/// Borrowed `/ppt/comments/commentN.xml` view.
pub struct ListPart<'a> {
    part: &'a dyn OpcPart,
}

impl<'a> ListPart<'a> {
    pub fn from_part(part: &'a dyn OpcPart) -> Result<Self> {
        validate_content_type(part, COMMENTS_CONTENT_TYPE)?;
        Ok(Self { part })
    }

    pub fn comments(&self) -> Result<Vec<Comment>> {
        super::parse_slide_comments(self.part.blob())
    }

    pub fn to_xml(&self, conformance: Conformance) -> Result<Vec<u8>> {
        super::write_slide_comments(&self.comments()?, conformance)
    }

    #[inline]
    pub fn part(&self) -> &'a dyn OpcPart {
        self.part
    }
}

/// Borrowed `/ppt/commentAuthors.xml` view.
pub struct AuthorListPart<'a> {
    part: &'a dyn OpcPart,
}

impl<'a> AuthorListPart<'a> {
    pub fn from_part(part: &'a dyn OpcPart) -> Result<Self> {
        validate_content_type(part, AUTHORS_CONTENT_TYPE)?;
        Ok(Self { part })
    }

    pub fn authors(&self) -> Result<Vec<Author>> {
        super::parse_comment_authors(self.part.blob())
    }

    pub fn to_xml(&self, conformance: Conformance) -> Result<Vec<u8>> {
        super::write_comment_authors(&self.authors()?, conformance)
    }

    #[inline]
    pub fn part(&self) -> &'a dyn OpcPart {
        self.part
    }
}
