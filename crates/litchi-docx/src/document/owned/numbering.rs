//! List definitions staged with a main-document edit (record 0251e).
//!
//! A new list is one abstract definition and one numbering instance spliced
//! into the numbering part (`crate::numbering::insert_lists`). The part is
//! replaced exactly, or created with its relationship when the document has
//! none, as a graph change (0251b), so the definition and the paragraphs
//! that join it publish in one commit.

use super::super::{Error, Refusal, Result};
use super::Edit;
use crate::Package;
use crate::numbering::{AddedList, ListKind};

const NUMBERING_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const STRICT_NUMBERING_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/numbering";
const NUMBERING_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

impl Edit<Package> {
    /// Add a list definition of `kind` and return its identifiers; a
    /// paragraph joins it by naming [`AddedList::num_id`] in its numbering
    /// properties. `start` restarts one level (0–8) at a value.
    ///
    /// The definition is OfficeCLI's `listStyle` template: nine levels,
    /// bullets `•`, `◦`, `▪` or numbers `1.`, `a.`, `i.` by level, indented
    /// 720 twips per level with a 360-twip hanging indent.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::Package;
    /// use litchi_docx::document::{ElementPath, ElementStep};
    /// use litchi_docx::numbering::ListKind;
    /// use litchi_docx::properties::Numbering;
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("First item");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let list = edit.add_list(ListKind::Decimal, None)?;
    /// edit.format_paragraph_at(
    ///     &ElementPath::new([ElementStep::Paragraph(Position::new(0))]),
    ///     |properties| {
    ///         let mut numbering = Numbering::default();
    ///         numbering.id = Some(list.num_id());
    ///         numbering.level = Some(0);
    ///         properties.numbering = Some(numbering);
    ///     },
    /// )?;
    /// let mut next = edit.commit()?.into_package();
    /// let reopened = Package::from_bytes(&next.to_bytes()?)?;
    /// let paragraph = &reopened.document()?.paragraphs()?[0];
    /// assert_eq!(paragraph.numbering()?.map(|n| n.num_id), Some(list.num_id()));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] for a start level above 8,
    /// or a typed error for a malformed numbering part.
    pub fn add_list(&mut self, kind: ListKind, start: Option<(u8, u32)>) -> Result<AddedList> {
        if start.is_some_and(|(level, _)| level > crate::numbering::MAX_LIST_LEVEL) {
            return Err(Error::Refused {
                position: 0,
                reason: Refusal::InvalidPropertyValue,
            });
        }
        let strict = self.inner.is_strict();
        let relationship = if strict {
            STRICT_NUMBERING_RELATIONSHIP
        } else {
            NUMBERING_RELATIONSHIP
        };
        let part = self.related_part(relationship)?;
        let current = match &part {
            Some((name, _)) => self
                .part_payload(name)?
                .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?
                .to_vec(),
            None => crate::numbering::empty_numbering_xml(strict),
        };
        let (abstract_num_id, num_id) = crate::numbering::next_ids(&current)?;
        let list = AddedList {
            kind,
            abstract_num_id,
            num_id,
            start,
        };
        let updated = crate::numbering::insert_lists(&current, &[list])?;
        match part {
            Some((name, _)) => {
                self.replace_part(&name, updated)?;
            },
            None => {
                self.stage_related_part(
                    relationship,
                    ("", "numbering", ".xml"),
                    true,
                    NUMBERING_CONTENT_TYPE,
                    updated,
                )?;
            },
        }
        Ok(list)
    }

    /// The kind of the list `num_id` names, from its level-0 number format
    /// (staged lists included): [`ListKind::Bullet`] for bullets,
    /// [`ListKind::Decimal`] for every numbered format, `None` when the
    /// document has no such list.
    ///
    /// # Errors
    ///
    /// Returns a typed error for a malformed numbering part.
    pub fn list_kind(&self, num_id: u32) -> Result<Option<ListKind>> {
        let relationship = if self.inner.is_strict() {
            STRICT_NUMBERING_RELATIONSHIP
        } else {
            NUMBERING_RELATIONSHIP
        };
        let Some((name, _)) = self.related_part(relationship)? else {
            return Ok(None);
        };
        let Some(xml) = self.part_payload(&name)? else {
            return Ok(None);
        };
        let collection = crate::numbering::parse_numbering(&xml)?;
        let Some(instance) = collection.get_num(num_id) else {
            return Ok(None);
        };
        Ok(collection
            .get_abstract_num(instance.abstract_num_id())
            .and_then(|definition| definition.level(0))
            .map(|level| {
                if level.format == crate::numbering::Format::Bullet {
                    ListKind::Bullet
                } else {
                    ListKind::Decimal
                }
            }))
    }
}
