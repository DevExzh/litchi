//! Private migration adapter for the Pages document reader.
//!
//! The canonical section model lives in [`litchi_pages`]. This module remains
//! private until the concurrent document-reader migration can replace its
//! legacy local names; it contains no model implementation or public alias.

pub(crate) type PagesSection = litchi_pages::Section;
pub(crate) type PagesSectionType = litchi_pages::SectionType;
