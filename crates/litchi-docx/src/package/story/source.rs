//! Lazy, source-backed capture of the validated Word story ownership graph.
//!
//! The owning [`super`] inventory reads every story eagerly because it already
//! owns an `OpcPackage`.  This adapter keeps the graph pass metadata-only and
//! admits payloads only after the central-directory sizes for the complete
//! reachable story set have been checked.  The retained [`PartData`] handles
//! are deliberate: on managed source packages they carry the cache reservation
//! which makes the borrowed source bytes safe to expose.

use super::{
    StoryDialect, StoryKind, StoryLimits, StoryOwner, invalid, kind_from_content_type,
    relationship_dialect, relationship_kind, validate_main_content_type, validate_root,
};
use crate::{Error, Result};
use litchi_opc::{
    OpcError, PackURI, PartData, PartReadSession, PartView, Relationship, SourceBackedPackage,
};
use std::collections::HashSet;

/// One reachable Word story whose payload remains attached to source-backed
/// cache ownership.
#[derive(Clone, Debug)]
pub(crate) struct SourceStoryPart {
    part: PackURI,
    kind: StoryKind,
    data: PartData,
}

impl SourceStoryPart {
    #[must_use]
    pub(crate) const fn part(&self) -> &PackURI {
        &self.part
    }

    #[must_use]
    pub(crate) const fn kind(&self) -> StoryKind {
        self.kind
    }

    #[must_use]
    pub(crate) fn source(&self) -> &[u8] {
        self.data.as_bytes()
    }

    /// Borrow the retained managed payload handle, when this record has been
    /// populated by [`capture_source`].
    #[allow(
        dead_code,
        reason = "managed source-backed consumers borrow the pin as they land"
    )]
    #[must_use]
    pub(crate) const fn data(&self) -> &PartData {
        &self.data
    }
}

/// Source-backed story ownership and payload inventory.
#[derive(Clone, Debug)]
pub(crate) struct SourceStoryInventory {
    dialect: StoryDialect,
    stories: Vec<SourceStoryPart>,
}

impl SourceStoryInventory {
    #[must_use]
    pub(crate) const fn dialect(&self) -> StoryDialect {
        self.dialect
    }

    /// Main story first, followed by the remaining stories in canonical part
    /// name order, matching the owning-package inventory.
    #[must_use]
    pub(crate) fn stories(&self) -> &[SourceStoryPart] {
        &self.stories
    }
}

#[derive(Clone, Debug)]
struct Candidate {
    part: PackURI,
    kind: StoryKind,
    content_type: String,
    owner: StoryOwner,
}

/// Capture the validated story graph from a deferred source package.
///
/// The first phase traverses relationships and content-type metadata only.
/// The second phase checks every reachable story's central-directory size and
/// the aggregate before calling [`PartReadSession::read`]. A later decoded-size
/// mismatch, source change, cancellation, or XML validation error therefore
/// cannot leave a partially returned inventory.
pub(crate) fn capture_source(
    package: &SourceBackedPackage,
    limits: StoryLimits,
    session: &mut PartReadSession<'_>,
) -> Result<SourceStoryInventory> {
    let result = capture_source_inner(package, limits, session);
    // A source or execution failure racing any semantic error wins.  This
    // keeps every early-return path under the same post-operation fence as a
    // successful payload read.
    with_fence(package, result)
}

fn capture_source_inner(
    package: &SourceBackedPackage,
    limits: StoryLimits,
    session: &mut PartReadSession<'_>,
) -> Result<SourceStoryInventory> {
    let limits = limits.validate()?;
    fence(package)?;

    let part_count = package.iter_parts().count();
    if part_count > limits.max_package_parts {
        return Err(invalid(format!(
            "DOCX package part count exceeds {}",
            limits.max_package_parts
        )));
    }
    fence(package)?;

    let root = root_owner(package)?;
    let dialect = relationship_dialect(&root.relationship_type)
        .ok_or_else(|| invalid("unsupported main-document relationship dialect"))?;
    let main_part = with_fence(package, package.main_document_part().map_err(Error::from))?;
    validate_main_content_type(main_part.content_type())?;
    let main = main_part.partname().clone();

    let capacity = part_count.min(limits.max_stories);
    let mut owned = HashSet::new();
    owned
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "source-backed DOCX story ownership set",
            source,
        })?;
    let mut candidates = Vec::new();
    candidates
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "source-backed DOCX story inventory",
            source,
        })?;
    let mut relationship_count = 0usize;
    push_candidate(
        &mut candidates,
        &mut owned,
        main_part,
        StoryKind::Main,
        root,
        limits,
    )?;

    let glossary = discover_owned(
        package,
        &main,
        true,
        dialect,
        limits,
        &mut owned,
        &mut candidates,
        &mut relationship_count,
    )?;
    if let Some(glossary) = glossary {
        discover_owned(
            package,
            &glossary,
            false,
            dialect,
            limits,
            &mut owned,
            &mut candidates,
            &mut relationship_count,
        )?;
    }

    // A non-glossary subsidiary story cannot be a relationship owner for a
    // second Word story.  This is a metadata-only graph check.
    for candidate in candidates
        .iter()
        .filter(|candidate| !matches!(candidate.kind, StoryKind::Main | StoryKind::Glossary))
    {
        fence(package)?;
        let part = get_part(package, &candidate.part)?;
        if part
            .rels()
            .iter()
            .any(|relationship| relationship_kind(relationship.reltype()).is_some())
        {
            return Err(invalid(format!(
                "story '{}' cannot own another Word story",
                candidate.part
            )));
        }
    }

    // Every Word story content type must have been reached through the
    // relationship graph; directory names and untyped references do not grant
    // ownership.
    for part in package.iter_parts() {
        fence(package)?;
        if kind_from_content_type(part.content_type()).is_some() && !owned.contains(part.partname())
        {
            return Err(invalid(format!(
                "Word story part '{}' is orphaned",
                part.partname()
            )));
        }
    }
    fence(package)?;

    candidates.sort_unstable_by(|left, right| left.part.as_str().cmp(right.part.as_str()));
    let main_index = candidates
        .iter()
        .position(|candidate| candidate.part == main)
        .ok_or_else(|| invalid("resolved main document disappeared during story capture"))?;
    candidates[..=main_index].rotate_right(1);

    measure_topology(&candidates, limits.max_topology_bytes)?;

    let (stories, _) = load_stories(package, &candidates, dialect, limits, session)?;
    fence(package)?;
    Ok(SourceStoryInventory { dialect, stories })
}

fn root_owner(package: &SourceBackedPackage) -> Result<StoryOwner> {
    fence(package)?;
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::OFFICE_DOCUMENT
                | litchi_opc::constants::relationship_type::STRICT_OFFICE_DOCUMENT
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    if relationships.next().is_some() {
        return Err(invalid("package has multiple main-document relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("main-document relationship cannot be external"));
    }
    let owner = StoryOwner {
        part: None,
        relationship_id: relationship.r_id().to_owned(),
        relationship_type: relationship.reltype().to_owned(),
        target_ref: relationship.target_ref().to_owned(),
    };
    fence(package)?;
    Ok(owner)
}

#[allow(
    clippy::too_many_arguments,
    reason = "signature mirrors the corresponding OOXML ownership record"
)]
fn discover_owned(
    package: &SourceBackedPackage,
    owner: &PackURI,
    allow_glossary: bool,
    dialect: StoryDialect,
    limits: StoryLimits,
    owned: &mut HashSet<PackURI>,
    candidates: &mut Vec<Candidate>,
    relationship_count: &mut usize,
) -> Result<Option<PackURI>> {
    let owner_part = get_part(package, owner)?;
    let mut relationships: Vec<(StoryKind, &Relationship)> = Vec::new();
    for relationship in owner_part.rels().iter() {
        fence(package)?;
        let Some((edge_dialect, kind)) = relationship_kind(relationship.reltype()) else {
            continue;
        };
        if edge_dialect != dialect {
            return Err(invalid(format!(
                "story '{owner}' mixes Strict and Transitional relationship dialects"
            )));
        }
        if kind == StoryKind::Glossary && !allow_glossary {
            return Err(invalid(
                "a glossary story cannot own another glossary story",
            ));
        }
        relationships
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "source-backed DOCX story ownership relationships",
                source,
            })?;
        relationships.push((kind, relationship));
        if relationships.len() > limits.max_relationships_per_owner {
            return Err(invalid(format!(
                "story '{}' relationship count exceeds {}",
                owner, limits.max_relationships_per_owner
            )));
        }
    }
    relationships.sort_unstable_by(|left, right| {
        left.1
            .r_id()
            .cmp(right.1.r_id())
            .then_with(|| left.1.reltype().cmp(right.1.reltype()))
            .then_with(|| left.1.target_ref().cmp(right.1.target_ref()))
    });

    let mut singleton = [false; 4];
    let mut glossary = None;
    for (kind, relationship) in relationships {
        if let Some(index) = kind.singleton_index()
            && std::mem::replace(&mut singleton[index], true)
        {
            return Err(invalid(format!(
                "story '{owner}' has multiple {kind:?} relationships"
            )));
        }
        *relationship_count = relationship_count
            .checked_add(1)
            .ok_or_else(|| invalid("DOCX story relationship count overflow"))?;
        if *relationship_count > limits.max_total_relationships {
            return Err(invalid(format!(
                "DOCX story relationship count exceeds {}",
                limits.max_total_relationships
            )));
        }
        if relationship.is_external() {
            return Err(invalid(format!(
                "story '{owner}' has an external {kind:?} relationship"
            )));
        }
        let requested = relationship.target_partname()?;
        let target_part = match get_part(package, &requested) {
            Ok(part) => part,
            Err(Error::Opc(OpcError::PartNotFound(_))) => {
                return Err(invalid(format!(
                    "story '{owner}' {kind:?} relationship targets missing part '{requested}'"
                )));
            },
            Err(error) => return Err(error),
        };
        let expected = kind
            .content_type()
            .ok_or_else(|| invalid("main story cannot be owned by another story"))?;
        if target_part.content_type() != expected {
            return Err(invalid(format!(
                "story '{}' {:?} relationship targets content type '{}'",
                owner,
                kind,
                target_part.content_type()
            )));
        }
        let target = target_part.partname().clone();
        if owned.contains(&target) {
            return Err(invalid(format!(
                "Word story part '{target}' has ambiguous ownership"
            )));
        }
        let incoming = StoryOwner {
            part: Some(owner.clone()),
            relationship_id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target_ref: relationship.target_ref().to_owned(),
        };
        push_candidate(candidates, owned, target_part, kind, incoming, limits)?;
        if kind == StoryKind::Glossary {
            glossary = Some(target);
        }
    }
    fence(package)?;
    Ok(glossary)
}

fn push_candidate(
    candidates: &mut Vec<Candidate>,
    owned: &mut HashSet<PackURI>,
    part: PartView<'_>,
    kind: StoryKind,
    owner: StoryOwner,
    limits: StoryLimits,
) -> Result<()> {
    if candidates.len() >= limits.max_stories {
        return Err(invalid(format!(
            "DOCX story count exceeds {}",
            limits.max_stories
        )));
    }
    let name = part.partname().clone();
    owned.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "source-backed DOCX story ownership set",
        source,
    })?;
    candidates
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "source-backed DOCX story inventory",
            source,
        })?;
    if !owned.insert(name.clone()) {
        return Err(invalid(format!(
            "Word story part '{name}' has ambiguous ownership"
        )));
    }
    candidates.push(Candidate {
        part: name,
        kind,
        content_type: part.content_type().to_owned(),
        owner,
    });
    Ok(())
}

fn measure_topology(candidates: &[Candidate], limit: usize) -> Result<()> {
    let main = candidates
        .first()
        .ok_or_else(|| invalid("Word story inventory is empty"))?;
    let mut size = super::TOPOLOGY_MAGIC.len();
    for value in [
        main.owner.relationship_id.as_bytes(),
        main.owner.relationship_type.as_bytes(),
        main.owner.target_ref.as_bytes(),
    ] {
        size = super::add_field_size(size, value.len())?;
    }
    size = super::add_size(size, 8)?;
    for candidate in candidates {
        size = super::add_field_size(size, candidate.part.as_str().len())?;
        size = super::add_field_size(size, candidate.content_type.len())?;
    }
    size = super::add_size(size, 8)?;
    for candidate in candidates.iter().skip(1) {
        let owner = candidate
            .owner
            .part
            .as_ref()
            .ok_or_else(|| invalid("subsidiary Word story is missing its owning part"))?;
        for length in [
            owner.as_str().len(),
            candidate.owner.relationship_id.len(),
            candidate.owner.relationship_type.len(),
            candidate.owner.target_ref.len(),
            candidate.part.as_str().len(),
        ] {
            size = super::add_field_size(size, length)?;
        }
        size = super::add_size(size, 1)?;
    }
    if size > limit {
        return Err(invalid(format!(
            "DOCX story topology exceeds {limit} bytes"
        )));
    }
    Ok(())
}

fn load_stories(
    package: &SourceBackedPackage,
    candidates: &[Candidate],
    dialect: StoryDialect,
    limits: StoryLimits,
    session: &mut PartReadSession<'_>,
) -> Result<(Vec<SourceStoryPart>, usize)> {
    // Read only central-directory metadata in this pass.  No call to data()
    // occurs until every individual and aggregate declaration is admitted.
    let mut declared = Vec::new();
    declared
        .try_reserve_exact(candidates.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed DOCX declared story sizes",
            source,
        })?;
    let mut declared_total = 0usize;
    for candidate in candidates {
        fence(package)?;
        let part = get_part(package, &candidate.part)?;
        let declared_size = with_fence(
            package,
            part.declared_uncompressed_size().map_err(Error::from),
        )?;
        let declared_size = usize::try_from(declared_size)
            .map_err(|_| invalid("DOCX declared story byte count does not fit usize"))?;
        if declared_size > limits.max_story_bytes {
            return Err(invalid(format!(
                "DOCX story '{}' has {declared_size} bytes, exceeding {}",
                candidate.part, limits.max_story_bytes
            )));
        }
        declared_total = declared_total
            .checked_add(declared_size)
            .ok_or_else(|| invalid("DOCX aggregate story byte count overflow"))?;
        if declared_total > limits.max_total_story_bytes {
            return Err(invalid(format!(
                "DOCX aggregate story bytes exceed {}",
                limits.max_total_story_bytes
            )));
        }
        declared.push(declared_size);
    }
    fence(package)?;

    let mut stories = Vec::new();
    stories
        .try_reserve_exact(candidates.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed DOCX story payload inventory",
            source,
        })?;
    let mut actual_total = 0usize;
    for (candidate, expected_size) in candidates.iter().zip(declared) {
        fence(package)?;
        let part = get_part(package, &candidate.part)?;
        let data = with_fence(package, session.read(part).map_err(Error::from))?;
        let actual_size = data.as_bytes().len();
        if actual_size != expected_size {
            return Err(invalid(format!(
                "DOCX story '{}' decoded {actual_size} bytes but the central directory declared {expected_size}",
                candidate.part
            )));
        }
        if actual_size > limits.max_story_bytes {
            return Err(invalid(format!(
                "DOCX story '{}' has {actual_size} bytes, exceeding {}",
                candidate.part, limits.max_story_bytes
            )));
        }
        actual_total = actual_total
            .checked_add(actual_size)
            .ok_or_else(|| invalid("DOCX aggregate story byte count overflow"))?;
        if actual_total > limits.max_total_story_bytes {
            return Err(invalid(format!(
                "DOCX aggregate story bytes exceed {}",
                limits.max_total_story_bytes
            )));
        }
        with_fence(
            package,
            validate_root(
                data.as_bytes(),
                candidate.kind,
                dialect,
                limits.max_xml_prolog_events,
            ),
        )?;
        stories.push(SourceStoryPart {
            part: candidate.part.clone(),
            kind: candidate.kind,
            data,
        });
    }
    fence(package)?;
    Ok((stories, actual_total))
}

fn get_part<'a>(package: &'a SourceBackedPackage, part: &PackURI) -> Result<PartView<'a>> {
    with_fence(package, package.part(part).map_err(Error::from))
}

/// Run both source-freshness and execution-policy checks.  They are evaluated
/// independently so a failed source check cannot suppress cancellation
/// observation on an error path.
fn fence(package: &SourceBackedPackage) -> Result<()> {
    let source = package.source_version().map(|_| ()).map_err(Error::from);
    let execution = package.check_execution().map_err(Error::from);
    match (source, execution) {
        (Err(source), _) => Err(source),
        (Ok(()), Err(execution)) => Err(execution),
        (Ok(()), Ok(())) => Ok(()),
    }
}

fn with_fence<T>(package: &SourceBackedPackage, result: Result<T>) -> Result<T> {
    let boundary = fence(package);
    match (result, boundary) {
        (Ok(value), Ok(())) => Ok(value),
        (Ok(_), Err(boundary)) => Err(boundary),
        (Err(error), Ok(())) => Err(error),
        (Err(_), Err(boundary)) => Err(boundary),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::OwnedSource;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, OpcPackage, PackageWriter, Part};
    use std::sync::Arc;

    fn xml(kind: StoryKind) -> Vec<u8> {
        let root = std::str::from_utf8(kind.root()).expect("ASCII root");
        format!(
            r#"<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#
        )
        .into_bytes()
    }

    fn source_package() -> SourceBackedPackage {
        let mut package = OpcPackage::new();
        let main_uri = PackURI::new("/word/document.xml").unwrap();
        let mut main = BlobPart::new(
            main_uri.clone(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            xml(StoryKind::Main),
        );
        main.rels_mut().add_relationship(
            rt::HEADER.to_owned(),
            "a-header.xml".to_owned(),
            "rHeaderA".to_owned(),
            false,
        );
        main.rels_mut().add_relationship(
            rt::HEADER.to_owned(),
            "z-header.xml".to_owned(),
            "rHeaderZ".to_owned(),
            false,
        );
        package.try_add_part(Box::new(main)).unwrap();
        for name in ["/word/a-header.xml", "/word/z-header.xml"] {
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(name).unwrap(),
                    ct::WML_HEADER.to_owned(),
                    xml(StoryKind::Header),
                )))
                .unwrap();
        }
        package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        let bytes = PackageWriter::to_bytes(&package).unwrap();
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap()
    }

    #[test]
    fn captures_sorted_source_stories_and_retains_part_data() {
        let package = source_package();
        let inventory = capture_source(
            &package,
            StoryLimits::default(),
            &mut package.read_session(),
        )
        .unwrap();
        assert_eq!(inventory.dialect(), StoryDialect::Transitional);
        let stories = inventory.stories();
        assert_eq!(stories.len(), 3);
        assert_eq!(stories[0].kind(), StoryKind::Main);
        assert_eq!(stories[1].part().as_str(), "/word/a-header.xml");
        assert_eq!(stories[2].part().as_str(), "/word/z-header.xml");
        for story in stories {
            assert_eq!(story.source(), story.data().as_bytes());
            assert!(!story.source().is_empty());
        }
    }

    #[test]
    fn central_directory_story_limits_are_checked_before_materialization() {
        let package = source_package();
        let limits = StoryLimits {
            max_story_bytes: 1,
            max_total_story_bytes: 1,
            ..StoryLimits::default()
        };
        assert!(capture_source(&package, limits, &mut package.read_session()).is_err());
        assert_eq!(package.cache_diagnostics().retained_bytes, 0);
    }
}
