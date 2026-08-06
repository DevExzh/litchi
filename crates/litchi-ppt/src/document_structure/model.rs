//! Semantic model for a legacy PPT `DocumentContainer`.

use crate::package::{Error, Result};

/// Bounded resource policy retained by a document-structure snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum serialized `DocumentContainer` size.
    pub max_bytes: usize,
    /// Maximum number of records in the owned record tree.
    pub max_records: usize,
    /// Maximum record nesting depth.
    pub max_depth: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_bytes: 64 * 1024 * 1024,
            max_records: 1_000_000,
            max_depth: 64,
        }
    }
}

impl Limits {
    pub(super) fn validate(self) -> Result<Self> {
        if self.max_bytes < 8 || self.max_records == 0 || self.max_depth == 0 {
            return Err(Error::InvalidFormat(
                "document-structure limits must be positive and permit a record header".into(),
            ));
        }
        Ok(self)
    }
}

/// Position of the optional PowerPoint 12 custom table-style package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CustomTableStylesPlacement {
    BeforeEndDocument,
    AfterEndDocument,
}

/// One typed `MasterPersistAtom` from the document's master list.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Master {
    child_index: usize,
    persist_id: u32,
    master_id: u32,
    flags: u32,
}

impl Master {
    pub(super) const fn new(
        child_index: usize,
        persist_id: u32,
        master_id: u32,
        flags: u32,
    ) -> Self {
        Self {
            child_index,
            persist_id,
            master_id,
            flags,
        }
    }

    /// Position of this reference in the master-list container.
    pub const fn child_index(self) -> usize {
        self.child_index
    }

    /// Persist-object identifier of the referenced master container.
    pub const fn persist_id(self) -> u32 {
        self.persist_id
    }

    /// Semantic master identifier carried by the reference.
    pub const fn master_id(self) -> u32 {
        self.master_id
    }

    /// The preserved `MasterPersistAtom` flags.
    pub const fn flags(self) -> u32 {
        self.flags
    }
}

/// One typed `SlidePersistAtom` from the presentation slide list.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Slide {
    child_index: usize,
    persist_id: u32,
    slide_id: u32,
    flags: u32,
    text_count: u32,
}

impl Slide {
    pub(super) const fn new(
        child_index: usize,
        persist_id: u32,
        slide_id: u32,
        flags: u32,
        text_count: u32,
    ) -> Self {
        Self {
            child_index,
            persist_id,
            slide_id,
            flags,
            text_count,
        }
    }

    /// Position of this reference in the slide-list container.
    pub const fn child_index(self) -> usize {
        self.child_index
    }

    /// Persist-object identifier of the referenced slide container.
    pub const fn persist_id(self) -> u32 {
        self.persist_id
    }

    /// Semantic slide identifier used for source-order selection.
    pub const fn slide_id(self) -> u32 {
        self.slide_id
    }

    /// The preserved `SlidePersistAtom` flags.
    pub const fn flags(self) -> u32 {
        self.flags
    }

    /// Number of text placeholders declared by the slide reference.
    pub const fn text_count(self) -> u32 {
        self.text_count
    }
}

/// Strictly validated structure and typed reference inventories of a
/// `DocumentContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentStructure {
    pub end_document_child_index: usize,
    pub custom_table_styles: Option<CustomTableStylesPlacement>,
    document_atom_child_index: usize,
    master_list_child_index: Option<usize>,
    slide_list_child_index: Option<usize>,
    notes_list_child_index: Option<usize>,
    masters: Vec<Master>,
    slides: Vec<Slide>,
}

impl DocumentStructure {
    pub(super) fn new(
        end_document_child_index: usize,
        custom_table_styles: Option<CustomTableStylesPlacement>,
        document_atom_child_index: usize,
        master_list_child_index: Option<usize>,
        slide_list_child_index: Option<usize>,
        notes_list_child_index: Option<usize>,
        masters: Vec<Master>,
        slides: Vec<Slide>,
    ) -> Self {
        Self {
            end_document_child_index,
            custom_table_styles,
            document_atom_child_index,
            master_list_child_index,
            slide_list_child_index,
            notes_list_child_index,
            masters,
            slides,
        }
    }

    /// Position of the required `DocumentAtom` in source order.
    pub const fn document_atom_child_index(&self) -> usize {
        self.document_atom_child_index
    }

    /// Position of the optional master reference list.
    pub const fn master_list_child_index(&self) -> Option<usize> {
        self.master_list_child_index
    }

    /// Position of the optional presentation slide reference list.
    pub const fn slide_list_child_index(&self) -> Option<usize> {
        self.slide_list_child_index
    }

    /// Position of the optional notes reference list.
    pub const fn notes_list_child_index(&self) -> Option<usize> {
        self.notes_list_child_index
    }

    /// Typed master references in their stored order.
    pub fn masters(&self) -> &[Master] {
        &self.masters
    }

    /// Typed presentation-slide references in their stored order.
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Look up a master by its semantic identifier.
    pub fn master(&self, master_id: u32) -> Option<Master> {
        self.masters
            .iter()
            .copied()
            .find(|master| master.master_id == master_id)
    }

    /// Look up a presentation slide by its semantic identifier.
    pub fn slide(&self, slide_id: u32) -> Option<Slide> {
        self.slides
            .iter()
            .copied()
            .find(|slide| slide.slide_id == slide_id)
    }
}
