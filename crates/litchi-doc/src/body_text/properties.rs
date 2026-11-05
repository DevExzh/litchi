//! Paragraph style and tab stops, `HYPERLINK` fields, standard bookmarks
//! and section page setup on the primary body transaction (record 0324).
//!
//! Each operation resolves its target against the staged text, refuses
//! tracked or structurally ambiguous content, stages one revision-editor
//! primitive and records a reversible [`Change::Mark`]. Setting the value a
//! target already has is an exact no-op.

use litchi_core::Position;

use super::{
    Change, Edit, Error, Refusal, Result, SourceSpan, TextTarget, cp_overflow, source_paragraphs,
    text_revision_intersects,
};
use crate::package::Error as PackageError;
use crate::parts::pap::{ParagraphProperties, TabAlignment, TabLeader, TabStop};
use crate::parts::styles::StyleKind;
use crate::section::{BreakKind, VerticalJustification};
use crate::sprm_operations::{
    SPRM_C_CV, SPRM_C_F_OBJ, SPRM_C_F_OLE2, SPRM_C_F_RMARK, SPRM_C_F_RMARK_DEL, SPRM_C_F_SPEC,
    SPRM_C_ICO, SPRM_C_ISTD, SPRM_C_KUL, SPRM_C_PIC_LOCATION, SPRM_P_CHG_TABS,
    SPRM_P_CHG_TABS_PAPX, SPRM_P_PROP_RMARK, SPRM_P_PROP_RMARK_CURRENT, SPRM_P_PROP_RMARK90,
};
use crate::tracked_revision::BookmarkRange;

/// Which part of a field a [`TextTarget::Field`] addresses.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum FieldPart {
    /// The instruction between the field's begin and separator marks.
    Instruction,
    /// The cached result between the field's separator and end marks.
    Result,
}

impl FieldPart {
    pub(super) const fn wire_name(self) -> &'static str {
        match self {
            Self::Instruction => "instruction",
            Self::Result => "result",
        }
    }
}

/// One section page-setup value set through `sprmS*` SEPX properties.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SectionValue {
    /// `sprmSBkc`: how the section starts.
    Break(BreakKind),
    /// `sprmSXaPage`: page width in twips.
    PageWidth(u16),
    /// `sprmSYaPage`: page height in twips.
    PageHeight(u16),
    /// `sprmSBOrientation`: `true` for landscape.
    Landscape(bool),
    /// `sprmSDyaTop`: signed top margin in twips (negative is exact).
    MarginTop(i16),
    /// `sprmSDyaBottom`: signed bottom margin in twips.
    MarginBottom(i16),
    /// `sprmSDxaLeft`: left margin in twips.
    MarginLeft(u16),
    /// `sprmSDxaRight`: right margin in twips.
    MarginRight(u16),
    /// `sprmSDyaHdrTop`: header distance in twips.
    HeaderDistance(u16),
    /// `sprmSDyaHdrBottom`: footer distance in twips.
    FooterDistance(u16),
    /// `sprmSDzaGutter`: gutter width in twips.
    Gutter(u16),
    /// `sprmSCcolumns`: the column count (1 through 45).
    Columns(u16),
    /// `sprmSDxaColumns`: the space between even columns in twips.
    ColumnSpace(u16),
    /// `sprmSFTitlePage`: a distinct first-page header and footer.
    TitlePage(bool),
    /// `sprmSVjc`: vertical alignment of the page content.
    VerticalAlignment(VerticalJustification),
}

const SPRM_S_BKC: u16 = 0x3009;
const SPRM_S_F_TITLE_PAGE: u16 = 0x300A;
const SPRM_S_CCOLUMNS: u16 = 0x500B;
const SPRM_S_DXA_COLUMNS: u16 = 0x900C;
const SPRM_S_DYA_HDR_TOP: u16 = 0xB017;
const SPRM_S_DYA_HDR_BOTTOM: u16 = 0xB018;
const SPRM_S_VJC: u16 = 0x301A;
const SPRM_S_B_ORIENTATION: u16 = 0x301D;
const SPRM_S_XA_PAGE: u16 = 0xB01F;
const SPRM_S_YA_PAGE: u16 = 0xB020;
const SPRM_S_DXA_LEFT: u16 = 0xB021;
const SPRM_S_DXA_RIGHT: u16 = 0xB022;
const SPRM_S_DYA_TOP: u16 = 0x9023;
const SPRM_S_DYA_BOTTOM: u16 = 0x9024;
const SPRM_S_DZA_GUTTER: u16 = 0xB025;

/// `flt` of a `HYPERLINK` field (MS-DOC 2.9.90).
const FLT_HYPERLINK: u8 = 0x58;

impl SectionValue {
    /// The SPRM that stores the value, and its encoded operand.
    fn sprm(self) -> Result<(u16, Vec<u8>)> {
        let word = |value: u16| value.to_le_bytes().to_vec();
        Ok(match self {
            Self::Break(kind) => (
                SPRM_S_BKC,
                vec![match kind {
                    BreakKind::Continuous => 0,
                    BreakKind::NewColumn => 1,
                    BreakKind::NewPage => 2,
                    BreakKind::EvenPage => 3,
                    BreakKind::OddPage => 4,
                }],
            ),
            Self::PageWidth(value) => (SPRM_S_XA_PAGE, word(value)),
            Self::PageHeight(value) => (SPRM_S_YA_PAGE, word(value)),
            Self::Landscape(value) => (SPRM_S_B_ORIENTATION, vec![if value { 2 } else { 1 }]),
            Self::MarginTop(value) => (SPRM_S_DYA_TOP, value.to_le_bytes().to_vec()),
            Self::MarginBottom(value) => (SPRM_S_DYA_BOTTOM, value.to_le_bytes().to_vec()),
            Self::MarginLeft(value) => (SPRM_S_DXA_LEFT, word(value)),
            Self::MarginRight(value) => (SPRM_S_DXA_RIGHT, word(value)),
            Self::HeaderDistance(value) => (SPRM_S_DYA_HDR_TOP, word(value)),
            Self::FooterDistance(value) => (SPRM_S_DYA_HDR_BOTTOM, word(value)),
            Self::Gutter(value) => (SPRM_S_DZA_GUTTER, word(value)),
            Self::Columns(value) => {
                if !(1..=45).contains(&value) {
                    return Err(Error::Refused(Refusal::InvalidProperty));
                }
                (SPRM_S_CCOLUMNS, word(value - 1))
            },
            Self::ColumnSpace(value) => {
                if value > 31_680 {
                    return Err(Error::Refused(Refusal::InvalidProperty));
                }
                (SPRM_S_DXA_COLUMNS, word(value))
            },
            Self::TitlePage(value) => (SPRM_S_F_TITLE_PAGE, vec![u8::from(value)]),
            Self::VerticalAlignment(value) => (
                SPRM_S_VJC,
                vec![match value {
                    VerticalJustification::Top => 0,
                    VerticalJustification::Center => 1,
                    VerticalJustification::Justified => 2,
                    VerticalJustification::Bottom => 3,
                }],
            ),
        })
    }
}

/// A `HYPERLINK` field to insert into a paragraph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HyperlinkInsert {
    /// The field instruction, `HYPERLINK "url"` and its switches.
    pub instruction: String,
    /// The displayed text.
    pub text: String,
}

/// A reversible record-0324 change, described for history and three-way
/// plans; patches still invert through their exact byte snapshots.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) enum MarkChange {
    ParagraphStyle {
        position: Position,
        before: u16,
        after: u16,
    },
    ParagraphTabs {
        position: Position,
        before: Vec<TabStop>,
        after: Vec<TabStop>,
    },
    Field {
        start_cp: u32,
        /// The paragraph and UTF-16 offset of an inserted field, with its
        /// contents.
        inserted: Option<(Position, u32, HyperlinkInsert)>,
        removed: bool,
    },
    Bookmark {
        name: String,
        before: Option<(u32, u32)>,
        after: Option<(u32, u32)>,
    },
    Section {
        section: Position,
        before: Vec<u8>,
        after: Vec<u8>,
    },
}

impl MarkChange {
    pub(super) fn effect(&self) -> String {
        match self {
            Self::ParagraphStyle { position, .. } => {
                format!("story/main/paragraph:{}/style", position.get())
            },
            Self::ParagraphTabs { position, .. } => {
                format!("story/main/paragraph:{}/tabs", position.get())
            },
            Self::Field { .. } => "story/main/structure".to_string(),
            Self::Bookmark { name, .. } => format!("resource/bookmark:{name}"),
            Self::Section { section, .. } => format!("section:{}", section.get()),
        }
    }

    pub(super) fn inverse(&self) -> Self {
        match self {
            Self::ParagraphStyle {
                position,
                before,
                after,
            } => Self::ParagraphStyle {
                position: *position,
                before: *after,
                after: *before,
            },
            Self::ParagraphTabs {
                position,
                before,
                after,
            } => Self::ParagraphTabs {
                position: *position,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Field {
                start_cp,
                inserted,
                removed,
            } => Self::Field {
                start_cp: *start_cp,
                inserted: inserted.clone(),
                removed: !*removed,
            },
            Self::Bookmark {
                name,
                before,
                after,
            } => Self::Bookmark {
                name: name.clone(),
                before: *after,
                after: *before,
            },
            Self::Section {
                section,
                before,
                after,
            } => Self::Section {
                section: *section,
                before: after.clone(),
                after: before.clone(),
            },
        }
    }

    pub(super) fn as_ref(&self) -> super::ChangeRef<'_> {
        match self {
            Self::ParagraphStyle {
                position,
                before,
                after,
            } => super::ChangeRef::ParagraphStyle {
                position: *position,
                before: *before,
                after: *after,
            },
            Self::ParagraphTabs {
                position,
                before,
                after,
            } => super::ChangeRef::ParagraphTabs {
                position: *position,
                before,
                after,
            },
            Self::Field {
                start_cp, removed, ..
            } => super::ChangeRef::Field {
                start_cp: *start_cp,
                before_present: *removed,
                after_present: !*removed,
            },
            Self::Bookmark {
                name,
                before,
                after,
            } => super::ChangeRef::Bookmark {
                name,
                before: *before,
                after: *after,
            },
            Self::Section {
                section,
                before,
                after,
            } => super::ChangeRef::Section {
                section: *section,
                before,
                after,
            },
        }
    }

    /// Stage this change's after state again (three-way plans).
    pub(super) fn apply_after(&self, edit: &mut Edit) -> Result<()> {
        match self {
            Self::ParagraphStyle {
                position, after, ..
            } => edit.set_paragraph_style(*position, *after),
            Self::ParagraphTabs {
                position, after, ..
            } => edit.set_paragraph_tabs(*position, after),
            Self::Field {
                start_cp,
                inserted,
                removed,
            } => match (removed, inserted) {
                (true, _) => edit.remove_field(*start_cp),
                (false, Some((paragraph, offset, link))) => {
                    edit.insert_hyperlink(*paragraph, *offset, link)
                },
                (false, None) => Err(Error::Refused(Refusal::FieldNotFound)),
            },
            Self::Bookmark { name, after, .. } => match after {
                Some((start, end)) => edit.add_bookmark_range(name, *start, *end),
                None => edit.remove_bookmark(name),
            },
            Self::Section { section, after, .. } => edit.set_section_grpprl(*section, after),
        }
    }
}

fn invalid(message: &str) -> Error {
    Error::Invalid(PackageError::Corrupted(message.to_string()))
}

/// The style index of a PAPX.
fn papx_style(grpprl: &[u8]) -> Result<u16> {
    grpprl
        .get(..2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .ok_or_else(|| invalid("PAPX has no style index"))
}

/// The SPRMs of a PAPX body, its FKP padding byte allowed.
fn papx_body_sprms(body: &[u8]) -> Result<Vec<crate::sprm::Sprm>> {
    match crate::sprm::parse_sprms(body) {
        Ok(sprms) => Ok(sprms),
        Err(_) if body.last() == Some(&0) => crate::sprm::parse_sprms(&body[..body.len() - 1])
            .map_err(|_error| invalid("malformed PAPX")),
        Err(_) => Err(invalid("malformed PAPX")),
    }
}

/// The body of a PAPX without its FKP padding byte and without the SPRMs
/// named by `remove`.
fn papx_body_without(body: &[u8], remove: &[u16]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(body.len());
    for sprm in papx_body_sprms(body)? {
        if !remove.contains(&sprm.opcode) {
            output.extend_from_slice(&body[sprm.offset..sprm.offset + sprm.size]);
        }
    }
    Ok(output)
}

fn grpprl_without(grpprl: &[u8], remove: &[u16]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(grpprl.len());
    for sprm in crate::sprm::parse_sprms(grpprl).map_err(|_error| invalid("malformed grpprl"))? {
        if !remove.contains(&sprm.opcode) {
            output.extend_from_slice(&grpprl[sprm.offset..sprm.offset + sprm.size]);
        }
    }
    Ok(output)
}

const PARAGRAPH_REVISIONS: [u16; 3] = [
    SPRM_P_PROP_RMARK,
    SPRM_P_PROP_RMARK90,
    SPRM_P_PROP_RMARK_CURRENT,
];

/// The TBD byte of a tab stop.
const fn tab_descriptor(stop: &TabStop) -> u8 {
    let alignment = match stop.alignment {
        TabAlignment::Left => 0,
        TabAlignment::Center => 1,
        TabAlignment::Right => 2,
        TabAlignment::Decimal => 3,
        TabAlignment::Bar => 4,
        TabAlignment::List => 6,
    };
    let leader = match stop.leader {
        TabLeader::None => 0,
        TabLeader::Dots => 1,
        TabLeader::Hyphens => 2,
        TabLeader::Underline => 3,
        TabLeader::Heavy => 4,
        TabLeader::MiddleDot => 5,
        TabLeader::DefaultLeader => 7,
    };
    alignment | (leader << 3)
}

impl Edit {
    /// The ordinary main-story paragraph at `position`, and the PAPX of its
    /// mark, refusing tracked paragraph marks and tracked paragraph
    /// formatting.
    fn untracked_paragraph(&self, position: Position) -> Result<(u32, u32, Vec<u8>)> {
        let paragraphs = source_paragraphs(&self.editor)?;
        let paragraph = paragraphs
            .get(position.get())
            .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
        let (start, mark) = (paragraph.start_cp, paragraph.end_cp);
        let after_mark = mark.checked_add(1).ok_or_else(cp_overflow)?;
        if text_revision_intersects(&self.editor, mark, after_mark)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        let grpprl = self.editor.paragraph_grpprl(mark).map_err(Error::Invalid)?;
        if papx_body_sprms(grpprl.get(2..).unwrap_or_default())?
            .iter()
            .any(|sprm| PARAGRAPH_REVISIONS.contains(&sprm.opcode))
        {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        Ok((start, mark, grpprl))
    }

    /// Apply paragraph style `istd` to the ordinary main-story paragraph at
    /// `position`, keeping its direct paragraph properties. The style must
    /// be a paragraph style of the document's stylesheet.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::StyleNotFound`] for an empty or non-paragraph
    /// style slot, and a typed refusal for an absent, tracked or huge-PAPX
    /// paragraph.
    pub(crate) fn set_paragraph_style(&mut self, position: Position, istd: u16) -> Result<()> {
        let sheet = self.editor.stylesheet().map_err(Error::Invalid)?;
        if sheet
            .get(istd)
            .is_none_or(|style| style.kind != StyleKind::Paragraph)
        {
            return Err(Error::Refused(Refusal::StyleNotFound));
        }
        let (start, mark, grpprl) = self.untracked_paragraph(position)?;
        let before = papx_style(&grpprl)?;
        if before == istd {
            return Ok(());
        }
        self.ensure_operation_capacity()?;
        let mut after = grpprl;
        after[..2].copy_from_slice(&istd.to_le_bytes());
        self.editor
            .set_paragraph_grpprl(start, mark, after)
            .map_err(Error::Invalid)?;
        self.changes.push(Change::Mark(MarkChange::ParagraphStyle {
            position,
            before,
            after: istd,
        }));
        Ok(())
    }

    /// The effective tab stops of the paragraph at `position`: its style's
    /// stops changed by its own.
    pub(crate) fn paragraph_tabs(&self, position: Position) -> Result<Vec<TabStop>> {
        let (_, _, grpprl) = self.untracked_paragraph(position)?;
        self.effective_tabs(&grpprl)
    }

    fn effective_tabs(&self, grpprl: &[u8]) -> Result<Vec<TabStop>> {
        let sheet = self.editor.stylesheet().map_err(Error::Invalid)?;
        let (_, mut sprms, _) = sheet
            .resolve_paragraph_style_sprms(papx_style(grpprl)?)
            .map_err(Error::Invalid)?;
        sprms.extend_from_slice(&papx_body_without(
            grpprl.get(2..).unwrap_or_default(),
            &[],
        )?);
        Ok(ParagraphProperties::from_sprm(&sprms)
            .map_err(Error::Invalid)?
            .tab_stops)
    }

    /// Set the tab stops of the ordinary main-story paragraph at `position`
    /// to `stops`, through one `sprmPChgTabsPapx` that clears every stop the
    /// paragraph has and adds `stops`.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidProperty`] for a position outside ±31680
    /// twips, duplicate positions or more than 64 stops, and a typed refusal
    /// for an absent or tracked paragraph or one whose direct properties use
    /// the style-only `sprmPChgTabs`.
    pub(crate) fn set_paragraph_tabs(
        &mut self,
        position: Position,
        stops: &[TabStop],
    ) -> Result<()> {
        let mut wanted = stops.to_vec();
        wanted.sort_by_key(|stop| stop.position);
        if wanted.len() > 64
            || wanted
                .windows(2)
                .any(|pair| pair[0].position == pair[1].position)
            || wanted
                .iter()
                .any(|stop| !(-31_680..=31_680).contains(&stop.position))
        {
            return Err(Error::Refused(Refusal::InvalidProperty));
        }
        let (start, mark, grpprl) = self.untracked_paragraph(position)?;
        let body = grpprl.get(2..).unwrap_or_default();
        let sprms = papx_body_sprms(body)?;
        if sprms.iter().any(|sprm| sprm.opcode == SPRM_P_CHG_TABS) {
            return Err(Error::Refused(Refusal::FormattingDependency));
        }
        let before = self.effective_tabs(&grpprl)?;
        if before == wanted {
            return Ok(());
        }
        // Clear every stop the paragraph shows, and every stop its old
        // direct change deleted (the style may still define those).
        let mut deleted = before.iter().map(|stop| stop.position).collect::<Vec<_>>();
        for sprm in sprms
            .iter()
            .filter(|sprm| sprm.opcode == SPRM_P_CHG_TABS_PAPX)
        {
            let operand = &sprm.operand;
            let count = usize::from(*operand.first().unwrap_or(&0));
            for index in 0..count {
                let at = 1 + index * 2;
                if let Some(bytes) = operand.get(at..at + 2) {
                    deleted.push(i32::from(i16::from_le_bytes([bytes[0], bytes[1]])));
                }
            }
        }
        deleted.sort_unstable();
        deleted.dedup();
        if deleted.len() > 64 {
            return Err(Error::Refused(Refusal::InvalidProperty));
        }
        let mut operand = Vec::new();
        operand.push(u8::try_from(deleted.len()).map_err(|_error| invalid("tab count overflow"))?);
        for position in &deleted {
            let position =
                i16::try_from(*position).map_err(|_error| invalid("tab position overflow"))?;
            operand.extend_from_slice(&position.to_le_bytes());
        }
        operand.push(u8::try_from(wanted.len()).map_err(|_error| invalid("tab count overflow"))?);
        for stop in &wanted {
            let position =
                i16::try_from(stop.position).map_err(|_error| invalid("tab position overflow"))?;
            operand.extend_from_slice(&position.to_le_bytes());
        }
        operand.extend(wanted.iter().map(tab_descriptor));
        let size = u8::try_from(operand.len())
            .map_err(|_error| Error::Refused(Refusal::InvalidProperty))?;
        let mut after = grpprl[..2].to_vec();
        after.extend_from_slice(&papx_body_without(body, &[SPRM_P_CHG_TABS_PAPX])?);
        if !(deleted.is_empty() && wanted.is_empty()) {
            after.extend_from_slice(&SPRM_P_CHG_TABS_PAPX.to_le_bytes());
            after.push(size);
            after.extend_from_slice(&operand);
        }
        self.ensure_operation_capacity()?;
        self.editor
            .set_paragraph_grpprl(start, mark, after)
            .map_err(Error::Invalid)?;
        self.changes.push(Change::Mark(MarkChange::ParagraphTabs {
            position,
            before,
            after: wanted,
        }));
        Ok(())
    }

    /// The main-story fields as (begin, separator, end) CPs, in begin order.
    fn main_fields(&self) -> Result<Vec<(u32, Option<u32>, u32)>> {
        let mut fields = Vec::new();
        let mut stack: Vec<(u32, Option<u32>)> = Vec::new();
        for (cp, record) in self.editor.main_field_markers() {
            match record[0] & 0x1F {
                0x13 => stack.push((cp, None)),
                0x14 => {
                    if let Some(open) = stack.last_mut() {
                        open.1 = Some(cp);
                    }
                },
                0x15 => {
                    let Some((begin, separator)) = stack.pop() else {
                        return Err(Error::Refused(Refusal::ComplexField));
                    };
                    fields.push((begin, separator, cp));
                },
                _ => return Err(Error::Refused(Refusal::ComplexField)),
            }
        }
        if !stack.is_empty() {
            return Err(Error::Refused(Refusal::ComplexField));
        }
        fields.sort_unstable();
        Ok(fields)
    }

    /// Insert a `HYPERLINK` field into the ordinary main-story paragraph at
    /// `position`, before UTF-16 offset `offset` of its text. The field's
    /// marks carry `fSpec`; its result is underlined in the hyperlink
    /// colour (`#0563C1`) over the formatting of the text it is inserted
    /// before.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal when the paragraph is absent, the offset is
    /// past its text or inside another field, the text there is tracked, or
    /// the instruction or text holds structural characters.
    pub(crate) fn insert_hyperlink(
        &mut self,
        position: Position,
        offset: u32,
        link: &HyperlinkInsert,
    ) -> Result<()> {
        if super::has_structural_content(&link.instruction)
            || super::has_structural_content(&link.text)
        {
            return Err(Error::Refused(
                Refusal::ReplacementContainsStructuralContent,
            ));
        }
        if super::keyword_of(&link.instruction).as_deref() != Some("HYPERLINK") {
            return Err(Error::Refused(Refusal::InvalidProperty));
        }
        let paragraphs = source_paragraphs(&self.editor)?;
        let paragraph = paragraphs
            .get(position.get())
            .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
        let length =
            u32::try_from(paragraph.text.encode_utf16().count()).map_err(|_error| cp_overflow())?;
        if offset > length {
            return Err(Error::Refused(Refusal::TargetNotFound));
        }
        let at = paragraph
            .start_cp
            .checked_add(offset)
            .ok_or_else(cp_overflow)?;
        if self
            .main_fields()?
            .iter()
            .any(|(begin, _, end)| *begin < at && at <= *end)
        {
            return Err(Error::Refused(Refusal::ComplexField));
        }
        let next = at.checked_add(1).ok_or_else(cp_overflow)?;
        if text_revision_intersects(&self.editor, at, next)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        self.length_change_admitted(at, at, 1)?;
        // The formatting of the text the link is inserted before, else of
        // the text before it; a producer whose CHPX leaves both uncovered
        // gets the style's formatting.
        let base = self
            .editor
            .character_grpprl_at(at)
            .or_else(|_error| match at.checked_sub(1) {
                Some(previous) => self.editor.character_grpprl_at(previous),
                None => Ok(Vec::new()),
            })
            .unwrap_or_default();
        let base = grpprl_without(
            &base,
            &[
                SPRM_C_F_SPEC,
                SPRM_C_F_OBJ,
                SPRM_C_F_OLE2,
                SPRM_C_PIC_LOCATION,
                SPRM_C_F_RMARK,
                SPRM_C_F_RMARK_DEL,
            ],
        )?;
        let mut special = base.clone();
        special.extend_from_slice(&SPRM_C_F_SPEC.to_le_bytes());
        special.push(1);
        let mut result = grpprl_without(&base, &[SPRM_C_KUL, SPRM_C_ICO, SPRM_C_CV, SPRM_C_ISTD])?;
        result.extend_from_slice(&SPRM_C_KUL.to_le_bytes());
        result.push(1);
        result.extend_from_slice(&SPRM_C_CV.to_le_bytes());
        result.extend_from_slice(&[0x05, 0x63, 0xC1, 0x00]);
        let units = |text: &str| text.encode_utf16().collect::<Vec<_>>();
        let instruction = units(&format!(" {} ", link.instruction.trim()));
        let text = units(&link.text);
        let total = 3 + instruction.len() + text.len();
        self.ensure_replacement_capacity(total)?;
        self.ensure_operation_capacity()?;
        let instruction_len = u32::try_from(instruction.len()).map_err(|_error| cp_overflow())?;
        let text_len = u32::try_from(text.len()).map_err(|_error| cp_overflow())?;
        let separator = at + 1 + instruction_len;
        let end = separator + 1 + text_len;
        let runs = vec![
            (vec![0x13], special.clone()),
            (instruction, base),
            (vec![0x14], special.clone()),
            (text, result),
            (vec![0x15], special),
        ];
        let runs = runs
            .into_iter()
            .filter(|(units, _)| !units.is_empty())
            .collect::<Vec<_>>();
        let fields = [
            (at, [0x13, FLT_HYPERLINK]),
            (separator, [0x14, 0x00]),
            (end, [0x15, 0x80]),
        ];
        self.editor
            .insert_runs_with_fields(at, &runs, &fields)
            .map_err(Error::Invalid)?;
        self.replacement_units = self.replacement_units.saturating_add(total);
        self.changes.push(Change::Mark(MarkChange::Field {
            start_cp: at,
            inserted: Some((position, offset, link.clone())),
            removed: false,
        }));
        Ok(())
    }

    /// Refuse a main-story length change that an unmodeled CP-indexed table
    /// or an ambiguous saved selection depends on, as text replacement does.
    fn length_change_admitted(&self, start: u32, end: u32, added: usize) -> Result<()> {
        if let Some(&fib_index) = self.editor.unmodeled_length_dependencies().first() {
            return Err(Error::Refused(Refusal::PositionDependency { fib_index }));
        }
        self.prepare_saved_selection_length_change(start, end, added)
    }

    /// Remove the whole main-story field that begins at `start_cp`, with
    /// its instruction, its result and every field nested in it.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::FieldNotFound`] when no field begins there, and a
    /// typed refusal when the field holds tracked text, a paragraph mark or
    /// another anchored range.
    pub(crate) fn remove_field(&mut self, start_cp: u32) -> Result<()> {
        let (_, _, end) = self
            .main_fields()?
            .into_iter()
            .find(|(begin, _, _)| *begin == start_cp)
            .ok_or(Error::Refused(Refusal::FieldNotFound))?;
        let after = end.checked_add(1).ok_or_else(cp_overflow)?;
        if text_revision_intersects(&self.editor, start_cp, after)? {
            return Err(Error::Refused(Refusal::TrackedText));
        }
        self.length_change_admitted(start_cp, after, 0)?;
        self.ensure_operation_capacity()?;
        self.editor
            .delete_range_with_fields(start_cp, after)
            .map_err(|error| match error {
                PackageError::Corrupted(message)
                    if message.contains("anchored range") || message.contains("paragraph") =>
                {
                    Error::Refused(Refusal::AnchoredContent)
                },
                other => Error::Invalid(other),
            })?;
        self.changes.push(Change::Mark(MarkChange::Field {
            start_cp,
            inserted: None,
            removed: true,
        }));
        Ok(())
    }

    /// The standard bookmarks, as (name, start CP, end CP) in start order.
    pub(crate) fn bookmarks(&self) -> Result<Vec<(String, u32, u32)>> {
        Ok(self
            .editor
            .bookmark_ranges()
            .map_err(Error::Invalid)?
            .into_iter()
            .map(|bookmark| (bookmark.name, bookmark.start, bookmark.end))
            .collect())
    }

    /// Add a standard bookmark named `name` over the UTF-16 range
    /// `start..end` of the ordinary main-story paragraph at `position`.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BookmarkName`] for a name that is empty, longer
    /// than 39 UTF-16 units, holds characters other than letters, digits,
    /// `_`, `.` and `-`, or is already used (names compare ASCII
    /// case-insensitively, as Word compares them).
    pub(crate) fn add_bookmark(
        &mut self,
        name: &str,
        position: Position,
        start: u32,
        end: u32,
    ) -> Result<()> {
        let paragraphs = source_paragraphs(&self.editor)?;
        let paragraph = paragraphs
            .get(position.get())
            .ok_or(Error::Refused(Refusal::ParagraphNotFound))?;
        let length =
            u32::try_from(paragraph.text.encode_utf16().count()).map_err(|_error| cp_overflow())?;
        if start > end || end > length {
            return Err(Error::Refused(Refusal::TargetNotFound));
        }
        let start = paragraph
            .start_cp
            .checked_add(start)
            .ok_or_else(cp_overflow)?;
        let end = paragraph
            .start_cp
            .checked_add(end)
            .ok_or_else(cp_overflow)?;
        self.add_bookmark_range(name, start, end)
    }

    fn add_bookmark_range(&mut self, name: &str, start: u32, end: u32) -> Result<()> {
        let units = name.encode_utf16().count();
        if units == 0
            || units >= 40
            || !name.chars().all(|character| {
                character.is_alphanumeric() || matches!(character, '_' | '.' | '-')
            })
        {
            return Err(Error::Refused(Refusal::BookmarkName));
        }
        let mut bookmarks = self.editor.bookmark_ranges().map_err(Error::Invalid)?;
        if bookmarks
            .iter()
            .any(|bookmark| bookmark.name.eq_ignore_ascii_case(name))
        {
            return Err(Error::Refused(Refusal::BookmarkName));
        }
        self.ensure_operation_capacity()?;
        bookmarks.push(BookmarkRange {
            name: name.to_string(),
            start,
            end,
        });
        self.editor
            .set_bookmarks(&bookmarks)
            .map_err(Error::Invalid)?;
        self.changes.push(Change::Mark(MarkChange::Bookmark {
            name: name.to_string(),
            before: None,
            after: Some((start, end)),
        }));
        Ok(())
    }

    /// Rename the standard bookmark named exactly `from` to `to`, keeping
    /// its range.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BookmarkNotFound`] when no bookmark is named
    /// `from`, and [`Refusal::BookmarkName`] for an invalid or used `to`.
    pub(crate) fn rename_bookmark(&mut self, from: &str, to: &str) -> Result<()> {
        if from == to {
            let exists = self
                .editor
                .bookmark_ranges()
                .map_err(Error::Invalid)?
                .iter()
                .any(|bookmark| bookmark.name == from);
            return if exists {
                Ok(())
            } else {
                Err(Error::Refused(Refusal::BookmarkNotFound))
            };
        }
        let bookmarks = self.editor.bookmark_ranges().map_err(Error::Invalid)?;
        let range = bookmarks
            .iter()
            .find(|bookmark| bookmark.name == from)
            .map(|bookmark| (bookmark.start, bookmark.end))
            .ok_or(Error::Refused(Refusal::BookmarkNotFound))?;
        self.remove_bookmark(from)?;
        self.add_bookmark_range(to, range.0, range.1)
    }

    /// Remove the standard bookmark named exactly `name`; the text it
    /// covered stays.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BookmarkNotFound`] when no bookmark has the name.
    pub(crate) fn remove_bookmark(&mut self, name: &str) -> Result<()> {
        let mut bookmarks = self.editor.bookmark_ranges().map_err(Error::Invalid)?;
        let index = bookmarks
            .iter()
            .position(|bookmark| bookmark.name == name)
            .ok_or(Error::Refused(Refusal::BookmarkNotFound))?;
        self.ensure_operation_capacity()?;
        let removed = bookmarks.remove(index);
        self.editor
            .set_bookmarks(&bookmarks)
            .map_err(Error::Invalid)?;
        self.changes.push(Change::Mark(MarkChange::Bookmark {
            name: removed.name,
            before: Some((removed.start, removed.end)),
            after: None,
        }));
        Ok(())
    }

    /// Set section page-setup values on section `section` (in `PlcfSed`
    /// order), replacing the section's SPRMs for those properties.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::SectionNotFound`] for a section past the last and
    /// [`Refusal::InvalidProperty`] for an out-of-range value.
    pub(crate) fn set_section_values(
        &mut self,
        section: Position,
        values: &[SectionValue],
    ) -> Result<()> {
        let sections = self.editor.sections().map_err(Error::Invalid)?;
        let (_, _, before) = sections
            .get(section.get())
            .ok_or(Error::Refused(Refusal::SectionNotFound))?;
        let mut encoded = Vec::with_capacity(values.len());
        for value in values {
            encoded.push(value.sprm()?);
        }
        let opcodes = encoded
            .iter()
            .map(|(opcode, _)| *opcode)
            .collect::<Vec<_>>();
        let mut after = grpprl_without(before, &opcodes)?;
        for (opcode, operand) in &encoded {
            after.extend_from_slice(&opcode.to_le_bytes());
            after.extend_from_slice(operand);
        }
        // Equal properties with the same encoding are an exact no-op, even
        // when the stored order differs.
        if same_sprms(before, &after)? {
            return Ok(());
        }
        self.set_section_grpprl(section, &after)
    }

    fn set_section_grpprl(&mut self, section: Position, after: &[u8]) -> Result<()> {
        let sections = self.editor.sections().map_err(Error::Invalid)?;
        let (_, _, before) = sections
            .get(section.get())
            .ok_or(Error::Refused(Refusal::SectionNotFound))?;
        if before.as_slice() == after {
            return Ok(());
        }
        self.ensure_operation_capacity()?;
        self.editor
            .set_section_grpprl(section.get(), after)
            .map_err(Error::Invalid)?;
        self.changes.push(Change::Mark(MarkChange::Section {
            section,
            before: before.clone(),
            after: after.to_vec(),
        }));
        Ok(())
    }
}

/// Whether two grpprls hold the same SPRMs, in any order.
fn same_sprms(left: &[u8], right: &[u8]) -> Result<bool> {
    let collect = |grpprl: &[u8]| -> Result<Vec<Vec<u8>>> {
        let mut sprms = crate::sprm::parse_sprms(grpprl)
            .map_err(|_error| invalid("malformed grpprl"))?
            .into_iter()
            .map(|sprm| grpprl[sprm.offset..sprm.offset + sprm.size].to_vec())
            .collect::<Vec<_>>();
        sprms.sort();
        Ok(sprms)
    };
    Ok(collect(left)? == collect(right)?)
}

/// The span of one part of the main-story field beginning at `start_cp`.
pub(super) fn field_span(
    editor: &crate::tracked_revision::RevisionEditor,
    start_cp: u32,
    part: FieldPart,
) -> Result<SourceSpan> {
    let text = editor.main_story_text().map_err(Error::Invalid)?;
    let mut cp = 0u32;
    let mut found = None;
    // (begin CP, start byte) of each open field, the target's depth.
    let mut depth = 0usize;
    let mut target_depth = None;
    let mut separator = None;
    for (byte, character) in text.char_indices() {
        let width = if character.len_utf16() == 1 { 1 } else { 2 };
        match character {
            '\u{13}' => {
                if cp == start_cp {
                    target_depth = Some(depth);
                    found = Some(byte + character.len_utf8());
                }
                depth += 1;
            },
            '\u{14}' => {
                if target_depth == Some(depth.saturating_sub(1)) && separator.is_none() {
                    separator = Some((cp, byte));
                }
            },
            '\u{15}' => {
                depth = depth.saturating_sub(1);
                if target_depth == Some(depth) {
                    let instruction_start = found.ok_or(Error::Refused(Refusal::FieldNotFound))?;
                    let (span_start, span_end, start_byte, end_byte) = match (part, separator) {
                        (FieldPart::Instruction, Some((separator_cp, separator_byte))) => (
                            start_cp + 1,
                            separator_cp,
                            instruction_start,
                            separator_byte,
                        ),
                        (FieldPart::Instruction, None) => {
                            (start_cp + 1, cp, instruction_start, byte)
                        },
                        (FieldPart::Result, Some((separator_cp, separator_byte))) => {
                            (separator_cp + 1, cp, separator_byte + 1, byte)
                        },
                        (FieldPart::Result, None) => {
                            return Err(Error::Refused(Refusal::ComplexField));
                        },
                    };
                    let slice = text[start_byte..end_byte].to_string();
                    if slice.contains('\u{13}') {
                        return Err(Error::Refused(Refusal::ComplexField));
                    }
                    return Ok(SourceSpan {
                        target: TextTarget::Field {
                            start: start_cp,
                            part,
                        },
                        start_cp: span_start,
                        end_cp: span_end,
                        text: slice,
                    });
                }
            },
            _ => {},
        }
        cp = cp.checked_add(width).ok_or_else(cp_overflow)?;
    }
    Err(Error::Refused(Refusal::FieldNotFound))
}
