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
use crate::numbering::catalog::{
    AbstractNumChange, Catalog, LevelChange, NewAbstractNum, NewLevel, NewNum,
};
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
        let strict = self.is_strict()?;
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
        let relationship = if self.is_strict()? {
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

    /// The numbering part as the staged changes leave it, or `None` when the
    /// document has none.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error, or a malformed-part error.
    pub fn numbering_catalog(&self) -> Result<Option<Catalog>> {
        Ok(self.numbering_part()?.map(|(_, catalog)| catalog))
    }

    /// Apply `change` to level `ilvl` of abstract definition `abstract_num_id`.
    ///
    /// ```
    /// use litchi_docx::numbering::ListKind;
    /// use litchi_docx::numbering::catalog::LevelChange;
    ///
    /// let package = litchi_docx::Builder::new()?.build()?;
    /// let mut edit = package.edit()?;
    /// let list = edit.add_list(ListKind::Decimal, None)?;
    /// edit.set_level(
    ///     list.abstract_num_id(),
    ///     0,
    ///     &LevelChange::default().with_format(Some("upperRoman".into())),
    /// )?;
    /// let next = edit.commit()?.into_package();
    /// let catalog = next.numbering_catalog()?.unwrap();
    /// assert_eq!(catalog.format_of(list.num_id(), 0), Some("upperRoman"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown definition or level,
    /// or a typed error for a property change that cannot be written.
    pub fn set_level(
        &mut self,
        abstract_num_id: u32,
        ilvl: u8,
        change: &LevelChange,
    ) -> Result<&mut Self> {
        let (name, catalog) = self.numbering_part()?.ok_or_else(not_found)?;
        if catalog
            .abstract_num(abstract_num_id)
            .and_then(|definition| definition.level(ilvl))
            .is_none()
        {
            return Err(not_found());
        }
        if let Some(updated) = catalog
            .with_level_changed(abstract_num_id, ilvl, change)
            .map_err(Error::Document)?
        {
            self.replace_part(&name, updated)?;
        }
        Ok(self)
    }

    /// Apply `change` to abstract definition `id`.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown definition.
    pub fn set_abstract_num(&mut self, id: u32, change: &AbstractNumChange) -> Result<&mut Self> {
        let (name, catalog) = self.numbering_part()?.ok_or_else(not_found)?;
        if catalog.abstract_num(id).is_none() {
            return Err(not_found());
        }
        if let Some(updated) = catalog
            .with_abstract_num_changed(id, change)
            .map_err(Error::Document)?
        {
            self.replace_part(&name, updated)?;
        }
        Ok(self)
    }

    /// Point numbering instance `id` at abstract definition `abstract_num_id`.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown instance or
    /// definition.
    pub fn set_num_definition(&mut self, id: u32, abstract_num_id: u32) -> Result<&mut Self> {
        let (name, catalog) = self.numbering_part()?.ok_or_else(not_found)?;
        if catalog.num(id).is_none() || catalog.abstract_num(abstract_num_id).is_none() {
            return Err(not_found());
        }
        if let Some(updated) = catalog
            .with_num_retargeted(id, abstract_num_id)
            .map_err(Error::Document)?
        {
            self.replace_part(&name, updated)?;
        }
        Ok(self)
    }

    /// Add a numbering instance and return its identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown definition, and
    /// [`Refusal::InvalidPropertyValue`] for an identifier in use or 0.
    pub fn add_num(&mut self, num: &NewNum) -> Result<u32> {
        let (name, catalog) = self.numbering_part()?.ok_or_else(not_found)?;
        let (updated, id) = catalog.with_num_added(num).map_err(refusal_of)?;
        self.replace_part(&name, updated)?;
        Ok(id)
    }

    /// Add an abstract definition, creating the numbering part when the
    /// document has none, and return its identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] for an identifier in use or
    /// a level that cannot be written.
    pub fn add_abstract_num(&mut self, definition: &NewAbstractNum) -> Result<u32> {
        match self.numbering_part()? {
            Some((name, catalog)) => {
                let (updated, id) = catalog
                    .with_abstract_num_added(definition)
                    .map_err(refusal_of)?;
                self.replace_part(&name, updated)?;
                Ok(id)
            },
            None => {
                let strict = self.is_strict()?;
                let empty = crate::numbering::empty_numbering_xml(strict);
                let catalog =
                    Catalog::from_xml(std::sync::Arc::new(empty)).map_err(Error::Document)?;
                let (updated, id) = catalog
                    .with_abstract_num_added(definition)
                    .map_err(refusal_of)?;
                self.stage_related_part(
                    numbering_relationship(strict),
                    ("", "numbering", ".xml"),
                    true,
                    NUMBERING_CONTENT_TYPE,
                    updated,
                )?;
                Ok(id)
            },
        }
    }

    /// Add `level` to abstract definition `id`, replacing the level of the
    /// same index when there is one.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown definition.
    pub fn add_level(&mut self, id: u32, level: &NewLevel) -> Result<&mut Self> {
        let (name, catalog) = self.numbering_part()?.ok_or_else(not_found)?;
        if catalog.abstract_num(id).is_none() {
            return Err(not_found());
        }
        let updated = catalog.with_level_added(id, level).map_err(refusal_of)?;
        self.replace_part(&name, updated)?;
        Ok(self)
    }

    /// Remove numbering instance `id`. Paragraphs that name it keep the
    /// reference, which Word reads as no numbering.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown instance.
    pub fn remove_num(&mut self, id: u32) -> Result<&mut Self> {
        let (name, catalog) = self.numbering_part()?.ok_or_else(not_found)?;
        if catalog.num(id).is_none() {
            return Err(not_found());
        }
        let updated = catalog.with_num_removed(id).map_err(Error::Document)?;
        self.replace_part(&name, updated)?;
        Ok(self)
    }

    /// Remove abstract definition `id` and the numbering instances of it.
    /// Paragraphs that name a removed instance keep the reference, which
    /// Word reads as no numbering.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown definition.
    pub fn remove_abstract_num(&mut self, id: u32) -> Result<&mut Self> {
        let (name, catalog) = self.numbering_part()?.ok_or_else(not_found)?;
        if catalog.abstract_num(id).is_none() {
            return Err(not_found());
        }
        let updated = catalog
            .with_abstract_num_removed(id)
            .map_err(Error::Document)?;
        self.replace_part(&name, updated)?;
        Ok(self)
    }

    /// Remove level `ilvl` of abstract definition `id`.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown definition or level.
    pub fn remove_level(&mut self, id: u32, ilvl: u8) -> Result<&mut Self> {
        let (name, catalog) = self.numbering_part()?.ok_or_else(not_found)?;
        if catalog
            .abstract_num(id)
            .and_then(|definition| definition.level(ilvl))
            .is_none()
        {
            return Err(not_found());
        }
        let updated = catalog
            .with_level_removed(id, ilvl)
            .map_err(Error::Document)?;
        self.replace_part(&name, updated)?;
        Ok(self)
    }

    /// The numbering part's name and catalog, staged changes included.
    fn numbering_part(&self) -> Result<Option<(String, Catalog)>> {
        let Some((name, _)) = self.related_part(numbering_relationship(self.is_strict()?))? else {
            return Ok(None);
        };
        let payload = self
            .part_payload(&name)?
            .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?;
        Ok(Some((
            name,
            Catalog::from_xml(payload).map_err(Error::Document)?,
        )))
    }
}

const fn numbering_relationship(strict: bool) -> &'static str {
    if strict {
        STRICT_NUMBERING_RELATIONSHIP
    } else {
        NUMBERING_RELATIONSHIP
    }
}

const fn not_found() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::BlockNotFound,
    }
}

/// A catalog refusal: an identifier in use, an unknown definition, or a
/// level that cannot be written.
fn refusal_of(error: crate::Error) -> Error {
    match error {
        crate::Error::InvalidFormat(message) if message.contains("no abstract definition") => {
            not_found()
        },
        crate::Error::InvalidFormat(message)
            if message.contains("taken") || message.contains("cannot be written") =>
        {
            Error::Refused {
                position: 0,
                reason: Refusal::InvalidPropertyValue,
            }
        },
        other => Error::Document(other),
    }
}
