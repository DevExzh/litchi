//! Typed semantic values and package-side run state for tracked revisions.

use crate::DateTime;
use crate::parts::fkp::ParagraphHeight;

/// A revision representation supported by binary DOC.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum RevisionKind {
    Insertion,
    Deletion,
    /// The deletion half of a move, paired by `revision_save_id`.
    MoveFrom,
    /// The insertion half of a move, paired by `revision_save_id`.
    MoveTo,
    CharacterFormatting,
    ParagraphFormatting,
    TableRowFormatting,
}

/// Metadata used to author or replace a revision mark.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RevisionMetadata {
    pub author: String,
    pub timestamp: Option<DateTime>,
    pub reason: Option<u16>,
    pub revision_save_id: Option<u32>,
}

impl RevisionMetadata {
    pub fn new(author: impl Into<String>) -> Self {
        Self {
            author: author.into(),
            timestamp: None,
            reason: None,
            revision_save_id: None,
        }
    }

    #[must_use]
    pub fn with_timestamp(mut self, timestamp: DateTime) -> Self {
        self.timestamp = Some(timestamp);
        self
    }

    #[must_use]
    pub fn with_reason(mut self, reason: u16) -> Self {
        self.reason = Some(reason);
        self
    }

    #[must_use]
    pub fn with_revision_save_id(mut self, revision_save_id: u32) -> Self {
        self.revision_save_id = Some(revision_save_id);
        self
    }
}

/// A tracked range in main-story CP coordinates.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Revision {
    pub kind: RevisionKind,
    pub start_cp: u32,
    pub end_cp: u32,
    pub author_index: u16,
    pub author: String,
    pub timestamp: Option<DateTime>,
    pub reason: Option<u16>,
    pub revision_save_id: Option<u32>,
    /// Move pair identity when binary insertion/deletion marks share an RSID.
    pub move_pair_id: Option<u32>,
}

#[derive(Clone, Debug)]
pub(super) struct RawPiece {
    pub(super) start: u32,
    pub(super) end: u32,
    pub(super) fc: u32,
    pub(super) unicode: bool,
    pub(super) prefix: [u8; 2],
    pub(super) prm: [u8; 2],
}

#[derive(Clone, Debug)]
pub(super) struct FcRun {
    pub(super) start: u32,
    pub(super) end: u32,
    pub(super) grpprl: Vec<u8>,
}

/// Direct table-membership and top-level row-terminator flags cached from PAPX.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[repr(transparent)]
pub(super) struct PapxTableState(u8);
impl PapxTableState {
    pub(super) fn new(in_table: bool, row_end: bool) -> Self {
        Self(u8::from(in_table) | (u8::from(row_end) << 1))
    }
    pub(super) const fn in_table(self) -> bool {
        self.0 & 1 != 0
    }
    pub(super) const fn row_end(self) -> bool {
        self.0 & 2 != 0
    }
}

#[derive(Clone, Debug)]
pub(super) struct PapxRun {
    pub(super) start: u32,
    pub(super) end: u32,
    pub(super) grpprl: Vec<u8>,
    /// Final `sprmPFInTable` and `sprmPFTtp` states. The one-byte cache
    /// keeps terminator queries allocation-free and is refreshed on edits.
    table_state: PapxTableState,
    pub(super) phe: ParagraphHeight,
}

impl PapxRun {
    pub(super) fn new(
        start: u32,
        end: u32,
        grpprl: Vec<u8>,
        table_state: PapxTableState,
        phe: ParagraphHeight,
    ) -> Self {
        Self {
            start,
            end,
            grpprl,
            table_state,
            phe,
        }
    }

    #[inline]
    #[must_use]
    pub(super) const fn in_table(&self) -> bool {
        self.table_state.in_table()
    }

    pub(super) const fn row_end(&self) -> bool {
        self.table_state.row_end()
    }

    pub(super) fn replace_grpprl(&mut self, grpprl: Vec<u8>, table_state: PapxTableState) {
        self.grpprl = grpprl;
        self.table_state = table_state;
    }
}

#[derive(Clone, Debug)]
pub(super) struct CpTable {
    pub(super) index: usize,
    pub(super) cps: Vec<u32>,
    pub(super) records: Vec<u8>,
}
