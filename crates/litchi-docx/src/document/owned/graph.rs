//! Package-graph changes staged with a main-document edit (record 0251b).
//!
//! An eager package-bound [`Edit`] can stage relationships of the main part,
//! new parts related from it, and exact replacements of existing parts. They
//! are recorded as [`super::super::Operation::ChangeGraph`] beside the
//! main-document operations, so one commit publishes both atomically, the
//! patch replays them under exact preconditions, and its inverse removes
//! them. Source-backed owners refuse them: their publication overlays the
//! main document only.

use std::collections::BTreeSet;
use std::sync::Arc;

use litchi_core::Position;
use litchi_opc::PackURI;

use super::super::transaction::{GraphChange, ReplacedPart};
use super::super::{ElementPath, NewRun, Result};
use super::Edit;
use crate::Package;
use crate::document::{TransferPart, TransferRelationship};

/// The target of a hyperlink: an external address, published as a main-part
/// relationship, or a bookmark in the document.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum HyperlinkTarget {
    /// An external URL (`r:id` to an external `hyperlink` relationship).
    Url(String),
    /// A bookmark name (`w:anchor`).
    Anchor(String),
}

const HYPERLINK_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const STRICT_HYPERLINK_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink";

impl Edit<Package> {
    /// Stage an external relationship of the main document part and return
    /// its new identifier.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for an unreadable main part, or a
    /// resource-limit error.
    pub fn add_external_relationship(
        &mut self,
        relationship_type: &str,
        target: &str,
    ) -> Result<String> {
        let id = self.next_relationship_id()?;
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: Arc::new([TransferRelationship {
                    id: id.clone(),
                    relationship_type: relationship_type.to_owned(),
                    target: target.to_owned(),
                    external: true,
                }]),
                parts: Arc::new([]),
                replaced: Arc::new([]),
            }),
            true,
        )?;
        Ok(id)
    }

    /// Stage a new part related from the main document part and return its
    /// part name and relationship identifier.
    ///
    /// The part is named `{directory}{stem}{n}{extension}` with the first
    /// `n` (from 1) no part or staged part uses, under the main part's
    /// folder; `directory` is relative to it (`media/`).
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for an invalid name, or a
    /// resource-limit error.
    pub fn add_related_part(
        &mut self,
        relationship_type: &str,
        directory: &str,
        stem: &str,
        extension: &str,
        content_type: &str,
        payload: Vec<u8>,
    ) -> Result<(String, String)> {
        self.stage_related_part(
            relationship_type,
            (directory, stem, extension),
            false,
            content_type,
            payload,
        )
    }

    /// As [`Self::add_related_part`]; with `unnumbered_first`, the plain
    /// `{directory}{stem}{extension}` is tried first, as Word names its
    /// singleton parts (`comments.xml`, `numbering.xml`).
    pub(super) fn stage_related_part(
        &mut self,
        relationship_type: &str,
        (directory, stem, extension): (&str, &str, &str),
        unnumbered_first: bool,
        content_type: &str,
        payload: Vec<u8>,
    ) -> Result<(String, String)> {
        let base = self.main_folder()?;
        let used = self.part_names()?;
        let plain = format!("{directory}{stem}{extension}");
        let (name, target) = if unnumbered_first && !used.contains(&format!("{base}{plain}")) {
            (format!("{base}{plain}"), plain)
        } else {
            let mut index = 1usize;
            loop {
                let target = format!("{directory}{stem}{index}{extension}");
                let name = format!("{base}{target}");
                if !used.contains(&name) {
                    break (name, target);
                }
                index += 1;
            }
        };
        PackURI::new(name.clone()).map_err(crate::Error::Invalid)?;
        let id = self.next_relationship_id()?;
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: Arc::new([TransferRelationship {
                    id: id.clone(),
                    relationship_type: relationship_type.to_owned(),
                    target,
                    external: false,
                }]),
                parts: Arc::new([TransferPart {
                    name: name.clone(),
                    content_type: content_type.to_owned(),
                    blob: Arc::new(payload),
                    relationships: Arc::new([]),
                }]),
                replaced: Arc::new([]),
            }),
            true,
        )?;
        Ok((name, id))
    }

    /// Stage an exact replacement of an existing part's payload.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for a missing part.
    pub fn replace_part(&mut self, name: &str, payload: Vec<u8>) -> Result<&mut Self> {
        let before = self
            .part_payload(name)?
            .ok_or_else(|| crate::Error::PartNotFound(name.to_owned()))?;
        if before.as_slice() == payload.as_slice() {
            return Ok(self);
        }
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: Arc::new([]),
                parts: Arc::new([]),
                replaced: Arc::new([ReplacedPart {
                    name: name.to_owned(),
                    before,
                    after: Arc::new(payload),
                }]),
            }),
            true,
        )?;
        Ok(self)
    }

    /// The payload a part has once the staged changes apply, or `None` when
    /// it does not exist.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for an unreadable part.
    pub fn part_payload(&self, name: &str) -> Result<Option<Arc<Vec<u8>>>> {
        let mut current = None;
        let uri = PackURI::new(name.to_owned()).map_err(crate::Error::Invalid)?;
        if let Ok(part) = self.package.opc_package().get_part(&uri) {
            current = Some(part.blob_arc());
        }
        self.inner_ref(|inner| {
            for (change, forward) in inner.staged_graph_changes() {
                if !forward {
                    continue;
                }
                for part in change.parts.iter().filter(|part| part.name == name) {
                    current = Some(Arc::clone(&part.blob));
                }
                for part in change.replaced.iter().filter(|part| part.name == name) {
                    current = Some(Arc::clone(&part.after));
                }
            }
            current
        })
    }

    /// The name and relationship identifier of the main part's first related
    /// part of `relationship_type`, staged parts included.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for an unreadable main part.
    pub fn related_part(&self, relationship_type: &str) -> Result<Option<(String, String)>> {
        let main = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?;
        if let Some(relationship) = main.rels().iter().find(|relationship| {
            relationship.reltype() == relationship_type && !relationship.is_external()
        }) {
            let name = relationship
                .target_partname()
                .map_err(crate::Error::from)?
                .as_str()
                .to_owned();
            return Ok(Some((name, relationship.r_id().to_owned())));
        }
        self.inner_ref(|inner| {
            for (change, forward) in inner.staged_graph_changes() {
                if !forward {
                    continue;
                }
                for relationship in change
                    .relationships
                    .iter()
                    .filter(|relationship| relationship.relationship_type == relationship_type)
                {
                    if let Some(part) = change.parts.first() {
                        return Some((part.name.clone(), relationship.id.clone()));
                    }
                }
            }
            None
        })
    }

    /// Insert a hyperlink holding `run` into the paragraph at `paragraph`,
    /// before its direct run `at` (after its last child for `None`), and
    /// return the hyperlink's position among the paragraph's direct
    /// hyperlinks.
    ///
    /// A URL target stages an external `hyperlink` relationship; the
    /// hyperlink declares the relationships namespace itself. A bookmark
    /// target is written as `w:anchor`.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::Package;
    /// use litchi_docx::document::{ElementPath, ElementStep, HyperlinkTarget, NewRun};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("See ");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// edit.insert_hyperlink_at(
    ///     &paragraph,
    ///     None,
    ///     &HyperlinkTarget::Url("https://example.test/".to_owned()),
    ///     &NewRun::new("the site"),
    /// )?;
    /// let mut next = edit.commit()?.into_package();
    /// let reopened = Package::from_bytes(&next.to_bytes()?)?;
    /// assert_eq!(reopened.document()?.text()?, "See the site");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the refusals of [`Self::insert_run_at`], or a resource-limit
    /// error.
    pub fn insert_hyperlink_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        target: &HyperlinkTarget,
        run: &NewRun,
    ) -> Result<Position> {
        let relationship = match target {
            HyperlinkTarget::Url(url) => {
                Some(self.add_external_relationship(self.hyperlink_relationship()?, url)?)
            },
            HyperlinkTarget::Anchor(_) => None,
        };
        let anchor = match target {
            HyperlinkTarget::Anchor(anchor) => Some(anchor.as_str()),
            HyperlinkTarget::Url(_) => None,
        };
        self.inner_mut()?.insert_hyperlink_markup_at(
            paragraph,
            at,
            relationship.as_deref(),
            anchor,
            run,
        )
    }

    /// Point the hyperlink at `hyperlink` to `target`, replacing its
    /// relationship reference or bookmark; its runs are unchanged.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a missing hyperlink, or a resource-limit
    /// error.
    pub fn set_hyperlink_target_at(
        &mut self,
        hyperlink: &ElementPath,
        target: &HyperlinkTarget,
    ) -> Result<&mut Self> {
        let relationship = match target {
            HyperlinkTarget::Url(url) => {
                Some(self.add_external_relationship(self.hyperlink_relationship()?, url)?)
            },
            HyperlinkTarget::Anchor(_) => None,
        };
        let anchor = match target {
            HyperlinkTarget::Anchor(anchor) => Some(anchor.as_str()),
            HyperlinkTarget::Url(_) => None,
        };
        self.inner_mut()?
            .retarget_hyperlink_at(hyperlink, relationship.as_deref(), anchor)?;
        Ok(self)
    }

    /// The hyperlink relationship type of the document's conformance.
    fn hyperlink_relationship(&self) -> Result<&'static str> {
        Ok(if self.is_strict()? {
            STRICT_HYPERLINK_RELATIONSHIP
        } else {
            HYPERLINK_RELATIONSHIP
        })
    }

    /// The folder of the main document part, with a trailing slash.
    pub(super) fn main_folder(&self) -> Result<String> {
        let main = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?;
        let name = main.partname().as_str();
        Ok(name[..=name.rfind('/').unwrap_or(0)].to_owned())
    }

    /// Every part name, staged parts included.
    pub(super) fn part_names(&self) -> Result<BTreeSet<String>> {
        let mut names: BTreeSet<String> = self
            .package
            .opc_package()
            .iter_parts()
            .map(|part| part.partname().as_str().to_owned())
            .collect();
        self.inner_ref(|inner| {
            for (change, _) in inner.staged_graph_changes() {
                names.extend(change.parts.iter().map(|part| part.name.clone()));
            }
            names
        })
    }

    /// The first `rIdN` the main part and the staged changes do not use.
    pub(super) fn next_relationship_id(&self) -> Result<String> {
        // One identifier was asked for, so one is returned.
        Ok(self.next_relationship_ids(1)?.swap_remove(0))
    }

    /// The first `count` `rIdN`s the main part and the staged changes do
    /// not use, ascending.
    pub(super) fn next_relationship_ids(&self, count: usize) -> Result<Vec<String>> {
        let main = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?;
        let mut used: BTreeSet<String> = main
            .rels()
            .iter()
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        self.inner_ref(|inner| {
            for (change, _) in inner.staged_graph_changes() {
                used.extend(
                    change
                        .relationships
                        .iter()
                        .map(|relationship| relationship.id.clone()),
                );
            }
        })?;
        let mut index = used.len() + 1;
        let mut ids = Vec::with_capacity(count);
        while ids.len() < count {
            let candidate = format!("rId{index}");
            if !used.contains(&candidate) {
                ids.push(candidate);
            }
            index += 1;
        }
        Ok(ids)
    }
}
