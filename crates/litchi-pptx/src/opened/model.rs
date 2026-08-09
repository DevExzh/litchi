//! Immutable opened-package state and finite resource policy.

use std::fmt;
use std::sync::Arc;

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, PackURI};
use sha2::{Digest, Sha256};

use crate::parts::PresentationPart;
use crate::{Error, Result};

/// Finite limits for opened-presentation transactions and durable patches.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_parts: usize,
    max_patch_bytes: usize,
    max_text_bytes: usize,
    max_history_entries: usize,
    max_history_bytes: usize,
}

impl Limits {
    /// Conservative defaults suitable for ordinary presentations.
    pub const DEFAULT: Self = Self {
        max_parts: 4_096,
        max_patch_bytes: 128 * 1024 * 1024,
        max_text_bytes: 8 * 1024 * 1024,
        max_history_entries: 64,
        max_history_bytes: 256 * 1024 * 1024,
    };

    /// Construct a finite, nonzero policy.
    #[must_use]
    pub const fn new(
        max_parts: usize,
        max_patch_bytes: usize,
        max_text_bytes: usize,
        max_history_entries: usize,
        max_history_bytes: usize,
    ) -> Option<Self> {
        if max_parts == 0
            || max_patch_bytes == 0
            || max_text_bytes == 0
            || max_history_entries == 0
            || max_history_bytes == 0
        {
            None
        } else {
            Some(Self {
                max_parts,
                max_patch_bytes,
                max_text_bytes,
                max_history_entries,
                max_history_bytes,
            })
        }
    }

    /// Maximum number of changed parts in one patch.
    #[must_use]
    pub const fn max_parts(self) -> usize {
        self.max_parts
    }

    /// Maximum encoded durable-patch size.
    #[must_use]
    pub const fn max_patch_bytes(self) -> usize {
        self.max_patch_bytes
    }

    /// Maximum replacement text size.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Maximum retained history entries.
    #[must_use]
    pub const fn max_history_entries(self) -> usize {
        self.max_history_entries
    }

    /// Maximum aggregate encoded history size.
    #[must_use]
    pub const fn max_history_bytes(self) -> usize {
        self.max_history_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Stable identity of one slide in current presentation order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Slide {
    pub(crate) id: u32,
    pub(crate) relationship_id: String,
    pub(crate) part_name: PackURI,
    pub(crate) name: String,
}

impl Slide {
    /// Stable `p:sldId@id` identity.
    #[must_use]
    pub const fn id(&self) -> u32 {
        self.id
    }

    /// Exact producer-visible name, with the part name as fallback.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Physical slide part name retained across ordering edits.
    #[must_use]
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }
}

/// Immutable root snapshot for an opened presentation.
#[derive(Clone)]
pub struct Snapshot {
    pub(crate) package: Arc<OpcPackage>,
    pub(crate) presentation_name: PackURI,
    pub(crate) slides: Vec<Slide>,
    pub(crate) revision: [u8; 32],
    pub(crate) limits: Limits,
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("presentation_name", &self.presentation_name)
            .field("slides", &self.slides)
            .field("revision", &self.revision)
            .field("limits", &self.limits)
            .finish_non_exhaustive()
    }
}

impl Snapshot {
    /// Slides in current presentation order.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Fingerprint of the complete captured OPC graph.
    #[must_use]
    pub const fn revision(&self) -> [u8; 32] {
        self.revision
    }

    /// Start one detached atomic transaction.
    #[must_use]
    pub fn edit(&self) -> super::Transaction {
        super::Transaction::new(self.clone())
    }

    /// Resource policy inherited by edits and patches.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }
}

pub(crate) fn capture(package: &OpcPackage, limits: Limits) -> Result<Snapshot> {
    let presentation = PresentationPart::from_package(package)?;
    let presentation_name = presentation.part().partname().clone();
    let references = presentation.slide_references()?;
    if references.len() > limits.max_parts {
        return Err(Error::Limit {
            resource: "opened-presentation slides",
            limit: limits.max_parts,
        });
    }
    let view = crate::presentation::Presentation::new(presentation, package);
    let contextual = view.slides()?;
    if references.len() != contextual.len() {
        return Err(invalid(
            "opened-presentation slide references do not resolve one-to-one",
        ));
    }
    let mut slides = Vec::new();
    slides
        .try_reserve_exact(references.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide identities",
            source,
        })?;
    let mut ids = std::collections::HashSet::new();
    let mut relationship_ids = std::collections::HashSet::new();
    let mut part_names = std::collections::HashSet::new();
    for (reference, slide) in references.iter().zip(contextual) {
        let relationship = presentation
            .part()
            .rels()
            .get(reference.relationship_id())
            .ok_or_else(|| invalid("opened-presentation slide relationship is missing"))?;
        if relationship.is_external()
            || !crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide")
        {
            return Err(invalid(
                "opened-presentation slide relationship is unsupported",
            ));
        }
        let target = relationship.target_partname()?;
        let part_name = slide.part().part().partname().clone();
        if target != part_name {
            return Err(invalid(
                "opened-presentation slide relationship target changed during capture",
            ));
        }
        if !ids.insert(reference.id())
            || !relationship_ids.insert(reference.relationship_id().to_owned())
            || !part_names.insert(part_name.clone())
        {
            return Err(invalid(
                "opened-presentation slide identities are not one-to-one",
            ));
        }
        slides.push(Slide {
            id: reference.id(),
            relationship_id: reference.relationship_id().to_owned(),
            name: slide.name()?,
            part_name,
        });
    }
    let _notes = crate::notes::load_snapshot(package, &presentation_name)?;
    let revision = package_fingerprint(package)?;
    Ok(Snapshot {
        package: Arc::new(package.clone()),
        presentation_name,
        slides,
        revision,
        limits,
    })
}

pub(crate) fn package_fingerprint(package: &OpcPackage) -> Result<[u8; 32]> {
    let mut parts: Vec<_> = package.iter_parts().collect();
    parts.sort_unstable_by(|left, right| left.partname().as_str().cmp(right.partname().as_str()));
    let mut digest = Sha256::new();
    feed(&mut digest, b"litchi-pptx-opened-v1");
    let mut root_relationships: Vec<_> = package.rels().iter().collect();
    root_relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    feed_relationships(&mut digest, &root_relationships)?;
    for part in parts {
        feed(&mut digest, part.partname().as_str().as_bytes());
        feed(&mut digest, part.content_type().as_bytes());
        feed(&mut digest, part.blob());
        let mut relationships: Vec<_> = part.rels().iter().collect();
        relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
        feed_relationships(&mut digest, &relationships)?;
    }
    Ok(digest.finalize().into())
}

pub(crate) fn part_context(part: &dyn litchi_opc::Part) -> Result<[u8; 32]> {
    let mut digest = Sha256::new();
    feed(&mut digest, b"litchi-pptx-opened-part-v1");
    feed(&mut digest, part.partname().as_str().as_bytes());
    feed(&mut digest, part.content_type().as_bytes());
    let mut relationships: Vec<_> = part.rels().iter().collect();
    relationships.sort_unstable_by(|left, right| left.r_id().cmp(right.r_id()));
    feed_relationships(&mut digest, &relationships)?;
    Ok(digest.finalize().into())
}

fn feed_relationships(
    digest: &mut Sha256,
    relationships: &[&litchi_opc::Relationship],
) -> Result<()> {
    let count = u32::try_from(relationships.len())
        .map_err(|_err| invalid("opened-presentation relationship count exceeds u32"))?;
    digest.update(count.to_le_bytes());
    for relationship in relationships {
        feed(digest, relationship.r_id().as_bytes());
        feed(digest, relationship.reltype().as_bytes());
        feed(digest, relationship.target_ref().as_bytes());
        digest.update([u8::from(relationship.is_external())]);
    }
    Ok(())
}

fn feed(digest: &mut Sha256, value: &[u8]) {
    digest.update(u64::try_from(value.len()).unwrap_or(u64::MAX).to_le_bytes());
    digest.update(value);
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
