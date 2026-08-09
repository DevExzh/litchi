//! Inert DOC context for the reserved user-defined `_PID_HLINKS` property.
//!
//! Numeric `dwApp` values can also be `FcCompressed` picture offsets. This
//! module exposes matching field markers only as candidates and requires an
//! explicit caller resolution before treating one as a field. It never parses,
//! normalizes, resolves, opens, or executes targets, locations, instructions,
//! or field results.

use crate::package::{Error as PackageError, Result};
use crate::parts::fields::{Field, FieldMarkerValue, FieldStory, FieldType, FieldsTable};
use litchi_ole_common::property_set::{Section, user_defined};
use std::cmp::Reverse;

pub use litchi_ole_common::property_set::user_defined::{Hyperlink, Hyperlinks, Limits};

/// One exact, story-local hyperlink field-begin candidate for a `dwApp` value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldCandidate {
    /// The document story that owns the `Plcfld`.
    pub story: FieldStory,
    /// The exact zero-based `aFld` index, including non-begin markers.
    pub plcfld_index: u32,
    /// The story-relative character position of the begin marker.
    pub begin_cp: u32,
    /// The paired structural field and its unevaluated instruction/result ranges.
    pub field: Field,
}

/// The inert DOC content associated with one stored hyperlink entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum HyperlinkAssociation {
    /// Matching field markers that still require caller proof.
    FieldCandidates(Vec<FieldCandidate>),
    /// A caller-proven exact field-begin marker.
    Field(FieldCandidate),
    /// The MS-DOC `OfficeArt` sentinel `dwApp == 0xFFFF_FFFF`.
    OfficeArtShape,
    /// Application data with no matching `HYPERLINK` field-begin marker.
    UnassociatedApplicationData,
}

/// A failure to promote one candidate to a caller-proven DOC field association.
#[derive(Debug, thiserror::Error, Clone, PartialEq, Eq)]
pub enum ResolutionError {
    #[error("hyperlink entry has no unresolved field candidates")]
    NoCandidates,
    #[error("selected DOC field candidate is not associated with this hyperlink entry")]
    NotCandidate,
    #[error("hyperlink entry is already resolved to a field")]
    AlreadyResolved,
}

/// A typed failure while staging a DOC hyperlink-property replacement.
#[derive(Debug, thiserror::Error)]
pub enum MutationError {
    #[error("cannot serialize changed _PID_HLINKS metadata with unresolved field candidates")]
    UnresolvedFieldCandidates,
    #[error(transparent)]
    Package(#[from] PackageError),
}

/// One inert stored `_PID_HLINKS` entry contextualized for a DOC package.
#[derive(Debug, Clone)]
pub struct UserDefinedHyperlink {
    raw: Hyperlink,
    association: HyperlinkAssociation,
}

impl UserDefinedHyperlink {
    fn from_raw(raw: Hyperlink, fields: Option<&FieldsTable>) -> Self {
        let association = association_for(raw.app(), fields);
        Self { raw, association }
    }

    /// Returns the stored hash without correcting a producer mismatch.
    #[must_use]
    pub const fn stored_hash(&self) -> u32 {
        self.raw.stored_hash()
    }
    /// Returns the recalculated shared hash without changing stored metadata.
    #[must_use]
    pub fn calculated_hash(&self) -> u32 {
        self.raw.calculated_hash()
    }
    /// Returns whether the producer's stored hash matches the shared algorithm.
    #[must_use]
    pub fn hash_matches(&self) -> bool {
        self.raw.hash_matches()
    }
    /// Returns the raw `dwApp` value.
    #[must_use]
    pub const fn app(&self) -> i32 {
        self.raw.app()
    }
    /// Returns the raw `dwOfficeArt` value.
    #[must_use]
    pub const fn office_art(&self) -> i32 {
        self.raw.office_art()
    }
    /// Returns the raw `dwInfo` value.
    #[must_use]
    pub const fn info(&self) -> i32 {
        self.raw.info()
    }
    /// Returns the stored target without URI parsing or normalization.
    #[must_use]
    pub fn target(&self) -> &str {
        self.raw.target()
    }
    /// Returns the stored location without URI parsing or normalization.
    #[must_use]
    pub fn location(&self) -> &str {
        self.raw.location()
    }
    /// Returns the structural DOC association or unresolved candidates.
    #[must_use]
    pub const fn association(&self) -> &HyperlinkAssociation {
        &self.association
    }

    /// Promotes one exact stored candidate to a caller-proven field association.
    ///
    /// The caller must independently establish that `dwApp` is a `Plcfld`
    /// index rather than an `FcCompressed` picture offset.
    pub fn resolve_field(
        &mut self,
        story: FieldStory,
        plcfld_index: u32,
    ) -> std::result::Result<(), ResolutionError> {
        let HyperlinkAssociation::FieldCandidates(candidates) = &self.association else {
            return match self.association {
                HyperlinkAssociation::Field(_) => Err(ResolutionError::AlreadyResolved),
                HyperlinkAssociation::OfficeArtShape
                | HyperlinkAssociation::UnassociatedApplicationData => {
                    Err(ResolutionError::NoCandidates)
                },
                HyperlinkAssociation::FieldCandidates(_) => unreachable!(),
            };
        };
        let candidate = candidates
            .iter()
            .find(|candidate| candidate.story == story && candidate.plcfld_index == plcfld_index)
            .cloned()
            .ok_or(ResolutionError::NotCandidate)?;
        self.association = HyperlinkAssociation::Field(candidate);
        Ok(())
    }
}

/// All stored `_PID_HLINKS` entries in property-array source order.
#[derive(Debug, Clone, Default)]
pub struct UserDefinedHyperlinks {
    entries: Vec<UserDefinedHyperlink>,
}

impl UserDefinedHyperlinks {
    /// Contextualizes immutable shared hyperlink metadata for one DOC field table.
    #[must_use]
    pub fn from_hyperlinks(hyperlinks: Hyperlinks, fields: Option<&FieldsTable>) -> Self {
        Self {
            entries: hyperlinks
                .iter()
                .cloned()
                .map(|raw| UserDefinedHyperlink::from_raw(raw, fields))
                .collect(),
        }
    }

    pub(crate) fn from_shared(hyperlinks: Hyperlinks, fields: Option<&FieldsTable>) -> Self {
        Self::from_hyperlinks(hyperlinks, fields)
    }
    /// Returns every stored entry in property-array source order.
    #[must_use]
    pub fn entries(&self) -> &[UserDefinedHyperlink] {
        &self.entries
    }
    /// Returns mutable source-order entries for explicit candidate resolution.
    pub fn entries_mut(&mut self) -> &mut [UserDefinedHyperlink] {
        &mut self.entries
    }

    /// Appends immutable shared metadata and recomputes its DOC candidates.
    pub fn push_hyperlink(&mut self, hyperlink: Hyperlink, fields: Option<&FieldsTable>) {
        self.entries
            .push(UserDefinedHyperlink::from_raw(hyperlink, fields));
    }

    /// Replaces one entry and recomputes candidates from the immutable raw value.
    ///
    /// Returns the prior contextual entry, or `None` when `index` is absent.
    pub fn replace_hyperlink(
        &mut self,
        index: usize,
        hyperlink: Hyperlink,
        fields: Option<&FieldsTable>,
    ) -> Option<UserDefinedHyperlink> {
        let replacement = UserDefinedHyperlink::from_raw(hyperlink, fields);
        let entry = self.entries.get_mut(index)?;
        Some(std::mem::replace(entry, replacement))
    }

    /// Removes and returns one source-order entry, if present.
    pub fn remove_hyperlink(&mut self, index: usize) -> Option<UserDefinedHyperlink> {
        (index < self.entries.len()).then(|| self.entries.remove(index))
    }

    /// Returns the MS-DOC section 2.4.7 ordering for explicitly resolved fields.
    /// Candidates, shapes, direct-picture offsets, and other application data
    /// retain source order relative to one another after resolved groups.
    #[must_use]
    pub fn canonicalized(&self) -> Self {
        let mut resolved: Vec<_> = self
            .entries
            .iter()
            .filter(|entry| matches!(entry.association, HyperlinkAssociation::Field(_)))
            .cloned()
            .collect();
        resolved.sort_by_key(|entry| match &entry.association {
            HyperlinkAssociation::Field(candidate) => {
                (story_rank(candidate.story), Reverse(candidate.plcfld_index))
            },
            HyperlinkAssociation::FieldCandidates(_)
            | HyperlinkAssociation::OfficeArtShape
            | HyperlinkAssociation::UnassociatedApplicationData => unreachable!(),
        });
        resolved.extend(
            self.entries
                .iter()
                .filter(|entry| !matches!(entry.association, HyperlinkAssociation::Field(_)))
                .cloned(),
        );
        Self { entries: resolved }
    }

    fn has_unresolved_candidates(&self) -> bool {
        self.entries
            .iter()
            .any(|entry| matches!(entry.association, HyperlinkAssociation::FieldCandidates(_)))
    }
    fn source_shared(&self) -> Hyperlinks {
        Hyperlinks::new(self.entries.iter().map(|entry| entry.raw.clone()).collect())
    }
    fn canonical_shared(&self) -> Hyperlinks {
        Hyperlinks::new(
            self.canonicalized()
                .entries
                .into_iter()
                .map(|entry| entry.raw)
                .collect(),
        )
    }
}

pub(crate) fn read(
    section: &Section,
    fields: Option<&FieldsTable>,
    limits: Limits,
) -> std::result::Result<Option<UserDefinedHyperlinks>, litchi_cfb::OleError> {
    user_defined::Properties::with_limits(section, limits)?
        .hyperlinks()
        .map(|value| value.map(|value| UserDefinedHyperlinks::from_shared(value, fields)))
}
pub(crate) fn put(
    section: &mut Section,
    hyperlinks: &UserDefinedHyperlinks,
    limits: Limits,
) -> std::result::Result<(), litchi_cfb::OleError> {
    user_defined::Edit::with_limits(section, limits)?.set_hyperlinks(hyperlinks.canonical_shared())
}
pub(crate) fn remove(
    section: &mut Section,
    limits: Limits,
) -> std::result::Result<bool, litchi_cfb::OleError> {
    Ok(user_defined::Edit::with_limits(section, limits)?.remove_hyperlinks())
}

/// Reads contextualized `_PID_HLINKS` metadata with default overlay limits.
pub fn from_user_defined_section(
    section: &Section,
    fields: Option<&FieldsTable>,
) -> Result<Option<UserDefinedHyperlinks>> {
    from_user_defined_section_with_limits(section, fields, Limits::default())
}
/// Reads contextualized `_PID_HLINKS` metadata with explicit overlay limits.
/// Limits apply after generic property-set parsing to the reserved BLOB overlay.
pub fn from_user_defined_section_with_limits(
    section: &Section,
    fields: Option<&FieldsTable>,
    limits: Limits,
) -> Result<Option<UserDefinedHyperlinks>> {
    read(section, fields, limits).map_err(PackageError::from)
}

pub(crate) fn unchanged_with_unresolved_candidates(
    section: Option<&Section>,
    hyperlinks: &UserDefinedHyperlinks,
    limits: Limits,
) -> std::result::Result<bool, litchi_cfb::OleError> {
    if !hyperlinks.has_unresolved_candidates() {
        return Ok(false);
    }
    let Some(section) = section else {
        return Ok(false);
    };
    let current = user_defined::Properties::with_limits(section, limits)?.hyperlinks()?;
    Ok(current.as_ref() == Some(&hyperlinks.source_shared()))
}

fn association_for(app: i32, fields: Option<&FieldsTable>) -> HyperlinkAssociation {
    if app == -1 {
        return HyperlinkAssociation::OfficeArtShape;
    }
    let Some(index) = u32::try_from(app).ok() else {
        return HyperlinkAssociation::UnassociatedApplicationData;
    };
    let Some(fields) = fields else {
        return HyperlinkAssociation::UnassociatedApplicationData;
    };
    let candidates: Vec<_> = FieldStory::ALL
        .into_iter()
        .filter_map(|story| {
            let table = fields.story(story)?;
            let marker = table.markers().get(usize::try_from(index).ok()?)?;
            if !matches!(
                marker.descriptor.value,
                FieldMarkerValue::Begin(FieldType::Hyperlink)
            ) {
                return None;
            }
            let field = table
                .fields()
                .iter()
                .find(|field| {
                    field.start_cp == marker.position && field.field_type == FieldType::Hyperlink
                })?
                .clone();
            Some(FieldCandidate {
                story,
                plcfld_index: index,
                begin_cp: marker.position,
                field,
            })
        })
        .collect();
    if candidates.is_empty() {
        HyperlinkAssociation::UnassociatedApplicationData
    } else {
        HyperlinkAssociation::FieldCandidates(candidates)
    }
}

fn story_rank(story: FieldStory) -> u8 {
    match story {
        FieldStory::Main => 0,
        FieldStory::Footnote => 1,
        FieldStory::Header => 2,
        FieldStory::Comment => 3,
        FieldStory::Endnote => 4,
        FieldStory::Textbox => 5,
        FieldStory::HeaderTextbox => 6,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parts::fields::FieldStoryTable;
    fn plcf(cps: &[u32], descriptors: &[[u8; 2]]) -> Vec<u8> {
        let mut bytes = Vec::new();
        for cp in cps {
            bytes.extend_from_slice(&cp.to_le_bytes());
        }
        for descriptor in descriptors {
            bytes.extend_from_slice(descriptor);
        }
        bytes
    }
    fn fields(stories: &[(FieldStory, u32, Vec<u8>)]) -> FieldsTable {
        FieldsTable::from_story_tables(
            stories
                .iter()
                .map(|(story, length, bytes)| {
                    FieldStoryTable::parse_plcf(*story, *length, bytes).unwrap()
                })
                .collect(),
        )
        .unwrap()
    }
    fn raw(app: i32, target: &str) -> Hyperlink {
        Hyperlink::new(app, 0, 0, target, "").unwrap()
    }

    #[test]
    fn matching_index_is_only_a_candidate_until_the_caller_resolves_it() {
        let fields = fields(&[(
            FieldStory::Main,
            30,
            plcf(
                &[2, 5, 10, 20, 30],
                &[[0x13, 0x58], [0x13, 0x58], [0x15, 0x40], [0x15, 0x00]],
            ),
        )]);
        let mut metadata = UserDefinedHyperlinks::from_shared(
            Hyperlinks::new(vec![raw(1, "nested")]),
            Some(&fields),
        );
        assert!(
            matches!(metadata.entries()[0].association(), HyperlinkAssociation::FieldCandidates(candidates) if candidates.len() == 1 && candidates[0].begin_cp == 5 && candidates[0].field.end_cp == 10)
        );
        metadata.entries_mut()[0]
            .resolve_field(FieldStory::Main, 1)
            .unwrap();
        assert!(matches!(
            metadata.entries()[0].association(),
            HyperlinkAssociation::Field(FieldCandidate {
                story: FieldStory::Main,
                plcfld_index: 1,
                ..
            })
        ));
    }
    #[test]
    fn same_index_across_stories_retains_every_candidate() {
        let story = |kind| (kind, 10, plcf(&[1, 4, 10], &[[0x13, 0x58], [0x15, 0x00]]));
        let fields = fields(&[story(FieldStory::Main), story(FieldStory::Footnote)]);
        let metadata = UserDefinedHyperlinks::from_shared(
            Hyperlinks::new(vec![raw(0, "ambiguous")]),
            Some(&fields),
        );
        assert!(
            matches!(metadata.entries()[0].association(), HyperlinkAssociation::FieldCandidates(candidates) if candidates.iter().map(|candidate| candidate.story).collect::<Vec<_>>() == vec![FieldStory::Main, FieldStory::Footnote])
        );
    }

    #[test]
    fn resolution_rejects_missing_wrong_and_repeated_candidates() {
        let fields = fields(&[(
            FieldStory::Main,
            10,
            plcf(&[1, 4, 10], &[[0x13, 0x58], [0x15, 0x00]]),
        )]);
        let mut metadata = UserDefinedHyperlinks::from_shared(
            Hyperlinks::new(vec![raw(-1, "shape"), raw(0, "field")]),
            Some(&fields),
        );
        assert_eq!(
            metadata.entries_mut()[0].resolve_field(FieldStory::Main, 0),
            Err(ResolutionError::NoCandidates)
        );
        assert_eq!(
            metadata.entries_mut()[1].resolve_field(FieldStory::Main, 1),
            Err(ResolutionError::NotCandidate)
        );
        metadata.entries_mut()[1]
            .resolve_field(FieldStory::Main, 0)
            .unwrap();
        assert_eq!(
            metadata.entries_mut()[1].resolve_field(FieldStory::Main, 0),
            Err(ResolutionError::AlreadyResolved)
        );
    }
    #[test]
    fn canonicalizes_only_resolved_entries_by_story_and_descending_index() {
        let fields = fields(&[(
            FieldStory::Main,
            20,
            plcf(
                &[1, 3, 5, 7, 20],
                &[[0x13, 0x58], [0x15, 0x00], [0x13, 0x58], [0x15, 0x00]],
            ),
        )]);
        let mut metadata = UserDefinedHyperlinks::from_shared(
            Hyperlinks::new(vec![raw(0, "first"), raw(2, "second"), raw(-1, "shape")]),
            Some(&fields),
        );
        metadata.entries_mut()[0]
            .resolve_field(FieldStory::Main, 0)
            .unwrap();
        metadata.entries_mut()[1]
            .resolve_field(FieldStory::Main, 2)
            .unwrap();
        assert_eq!(
            metadata
                .canonicalized()
                .entries()
                .iter()
                .map(UserDefinedHyperlink::target)
                .collect::<Vec<_>>(),
            vec!["second", "first", "shape"]
        );
    }
}
