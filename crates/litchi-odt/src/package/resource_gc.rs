//! Explicit, source-bound garbage collection for detached embedded payloads.
//!
//! Removing a `draw:object`, `draw:object-ole`, or `draw:image` owner never
//! invokes this module. Callers must name candidate package payloads, inspect
//! the non-mutating plan, and apply that plan to the exact source snapshot.

#![deny(
    clippy::expect_used,
    clippy::panic,
    clippy::todo,
    clippy::unimplemented,
    clippy::unwrap_used
)]

use crate::{
    Part,
    constants::{ODF_CONTENT, ODF_STYLES},
    core::OwnedPackage,
    transaction::{Commit, EnvelopeKind, Snapshot},
};
use litchi_core::{Error, Result};
use litchi_odf_common::{
    embedded::{Kind as ObjectKind, Source as ObjectSource, scan_package as scan_objects},
    media::{Source as ImageSource, scan_package as scan_images},
    package::{is_linked_href, resolve_package_path},
};
use quick_xml::{
    XmlVersion, events::Event, name::Namespace, name::ResolveResult, reader::NsReader,
};
use soapberry_zip::{
    CompressionMethod, PreservationIndex, ZipArchive, office::StreamingArchiveWriter,
};
use std::{
    collections::{BTreeMap, BTreeSet},
    ops::Range,
};

const MANIFEST_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const XML_NS: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XLINK_NS: &[u8] = b"http://www.w3.org/1999/xlink";
const MANIFEST_PATH: &str = "META-INF/manifest.xml";
const MAX_CANDIDATES: usize = 256;
const MAX_PACKAGE_ENTRIES: usize = 16_384;
const MAX_REFERENCE_XML_BYTES: usize = 64 * 1024 * 1024;

/// Maximum UTF-8 bytes accepted in one embedded-resource GC candidate path.
///
/// Planning checks this bound before cloning or normalizing caller paths, and
/// durable replay enforces the same ceiling before owning decoded path text.
pub const MAX_EMBEDDED_RESOURCE_GC_PATH_BYTES: usize = 4 * 1024;

/// One explicitly requested package payload to inspect for garbage collection.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
#[non_exhaustive]
pub enum EmbeddedResourceGcCandidate {
    /// One opaque embedded payload file, such as `Pictures/Image_1.png`.
    PackageFile(String),
    /// One embedded subdocument directory, such as `Object_1/`.
    PackageSubdocument(String),
}

impl EmbeddedResourceGcCandidate {
    /// Construct an exact package-file candidate.
    #[must_use]
    pub fn package_file(path: impl Into<String>) -> Self {
        Self::PackageFile(path.into())
    }

    /// Construct a package-subdocument directory candidate.
    #[must_use]
    pub fn package_subdocument(root_path: impl Into<String>) -> Self {
        Self::PackageSubdocument(root_path.into())
    }

    /// Return the caller-provided package path.
    #[must_use]
    pub fn path(&self) -> &str {
        match self {
            Self::PackageFile(path) | Self::PackageSubdocument(path) => path,
        }
    }
}

/// A conservative reason why a candidate cannot be deleted.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EmbeddedResourceGcRefusal {
    UnsafePath,
    MissingArchiveEntry,
    MissingManifestEntry,
    InvalidSubdocument,
    OverlappingCandidate,
    PackageNamespaceCollision,
    SignedPackage,
    EncryptedPackage,
    ProtectedDocument,
    /// An XML part contains a package-local reference outside the supported
    /// embedded owner grammar.
    UnknownReference {
        part: String,
    },
}

/// The deterministic decision for one candidate.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EmbeddedResourceGcDecision {
    /// The listed archive members and manifest records are proven unreferenced.
    Delete,
    /// At least one supported owner still shares this payload.
    RetainReferenced { supported_owner_count: usize },
    /// The candidate cannot be proven safe to delete.
    Refuse(EmbeddedResourceGcRefusal),
}

/// One candidate entry in an [`EmbeddedResourceGcPlan`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedResourceGcPlanEntry {
    candidate: EmbeddedResourceGcCandidate,
    decision: EmbeddedResourceGcDecision,
    archive_paths: Vec<String>,
    manifest_paths: Vec<String>,
}

impl EmbeddedResourceGcPlanEntry {
    #[must_use]
    pub const fn candidate(&self) -> &EmbeddedResourceGcCandidate {
        &self.candidate
    }

    #[must_use]
    pub const fn decision(&self) -> &EmbeddedResourceGcDecision {
        &self.decision
    }

    /// Exact ZIP members that would be removed when the decision is `Delete`.
    #[must_use]
    pub fn archive_paths(&self) -> &[String] {
        &self.archive_paths
    }

    /// Exact manifest records that would be removed when the decision is `Delete`.
    #[must_use]
    pub fn manifest_paths(&self) -> &[String] {
        &self.manifest_paths
    }
}

/// Non-mutating, bounded embedded-resource garbage-collection plan.
pub struct EmbeddedResourceGcPlan {
    source: Snapshot,
    entries: Vec<EmbeddedResourceGcPlanEntry>,
}

impl EmbeddedResourceGcPlan {
    #[must_use]
    pub fn entries(&self) -> &[EmbeddedResourceGcPlanEntry] {
        &self.entries
    }

    /// Whether every candidate has either a safe deletion or a supported live owner.
    #[must_use]
    pub fn is_applicable(&self) -> bool {
        self.entries
            .iter()
            .all(|entry| !matches!(entry.decision, EmbeddedResourceGcDecision::Refuse(_)))
    }

    /// Apply this plan only to its exact immutable source snapshot.
    ///
    /// Every untouched ZIP member is raw-copied. The manifest is changed by
    /// deleting only the planned `manifest:file-entry` byte spans. The result
    /// is fully reopened before a reversible package commit is returned.
    pub fn apply(&self, source: &Snapshot) -> Result<Commit> {
        if source.as_bytes() != self.source.as_bytes() {
            return invalid("ODT embedded-resource GC plan source does not match");
        }
        if !self.is_applicable() {
            return Err(Error::Unsupported(
                "ODT embedded-resource GC plan contains a typed refusal".to_string(),
            ));
        }
        let package = source.document()?.transaction_package().clone();
        let bytes = self.apply_to_package(&package)?;
        let after = if bytes.as_slice() == source.as_bytes() {
            source.clone()
        } else {
            Snapshot::from_bytes(bytes)?
        };
        Ok(crate::transaction::embedded_resource_gc_commit(
            source.clone(),
            after,
            self.entries
                .iter()
                .map(|entry| entry.candidate.clone())
                .collect(),
        ))
    }

    pub(crate) fn apply_to_package(&self, package: &OwnedPackage) -> Result<Vec<u8>> {
        let mut archive_paths = BTreeSet::new();
        let mut manifest_paths = BTreeSet::new();
        for entry in &self.entries {
            if entry.decision == EmbeddedResourceGcDecision::Delete {
                archive_paths.extend(entry.archive_paths.iter().cloned());
                manifest_paths.extend(entry.manifest_paths.iter().cloned());
            }
        }
        if archive_paths.is_empty() && manifest_paths.is_empty() {
            return Ok(package.as_bytes().to_vec());
        }
        rewrite_package(package, &archive_paths, &manifest_paths)
    }
}

impl Snapshot {
    /// Inventory explicitly named detached payloads without changing the package.
    ///
    /// This operation never discovers arbitrary files as garbage and is never
    /// invoked by owner removal. Candidates are bounded, normalized, sorted,
    /// and checked against every supported owner plus other package XML refs.
    pub fn plan_embedded_resource_gc(
        &self,
        candidates: &[EmbeddedResourceGcCandidate],
    ) -> Result<EmbeddedResourceGcPlan> {
        plan(self.clone(), candidates)
    }
}

pub(crate) fn plan(
    source: Snapshot,
    candidates: &[EmbeddedResourceGcCandidate],
) -> Result<EmbeddedResourceGcPlan> {
    if candidates.len() > MAX_CANDIDATES {
        return invalid(format!(
            "ODT embedded-resource GC exceeds {MAX_CANDIDATES} candidates"
        ));
    }
    for candidate in candidates {
        validate_candidate_path_bound(candidate.path())?;
    }
    let mut normalized = candidates
        .iter()
        .cloned()
        .map(NormalizedCandidate::new)
        .collect::<Vec<_>>();
    normalized.sort_by(|left, right| left.candidate.cmp(&right.candidate));
    mark_overlaps(&mut normalized);
    if normalized.is_empty() {
        return Ok(finish_plan(source, normalized));
    }
    match source.envelope_kind()? {
        EnvelopeKind::Signed => {
            set_global_refusal(&mut normalized, EmbeddedResourceGcRefusal::SignedPackage);
            return Ok(finish_plan(source, normalized));
        },
        EnvelopeKind::Encrypted => {
            set_global_refusal(&mut normalized, EmbeddedResourceGcRefusal::EncryptedPackage);
            return Ok(finish_plan(source, normalized));
        },
        EnvelopeKind::Plain => {},
    }
    let document = source.document()?;
    let package = document.transaction_package();
    let archive = package.package()?;
    let mut files = archive.files()?;
    if files.len() > MAX_PACKAGE_ENTRIES || archive.manifest().entries.len() > MAX_PACKAGE_ENTRIES {
        return invalid("ODT embedded-resource GC package entry limit exceeded");
    }
    files.sort();
    let duplicate_files = duplicate_values(&files);
    let manifest_paths = archive
        .manifest()
        .entries
        .keys()
        .cloned()
        .collect::<BTreeSet<_>>();

    let protected = document.protection()? != crate::protection::Policy::default();
    if protected {
        set_global_refusal(
            &mut normalized,
            EmbeddedResourceGcRefusal::ProtectedDocument,
        );
        return Ok(finish_plan(source, normalized));
    }

    for item in &mut normalized {
        if item.refusal.is_some() {
            continue;
        }
        classify_namespace(item, &files, &manifest_paths, &duplicate_files);
    }

    if normalized.iter().any(|item| item.refusal.is_none()) {
        classify_references(
            package,
            document.transaction_content_xml(),
            document.transaction_styles_xml(),
            &mut normalized,
        )?;
    }

    Ok(finish_plan(source, normalized))
}

fn set_global_refusal(candidates: &mut [NormalizedCandidate], refusal: EmbeddedResourceGcRefusal) {
    for candidate in candidates {
        if candidate.refusal.is_none() {
            candidate.refusal = Some(refusal.clone());
        }
    }
}

fn finish_plan(source: Snapshot, normalized: Vec<NormalizedCandidate>) -> EmbeddedResourceGcPlan {
    let entries = normalized
        .into_iter()
        .map(|item| {
            let decision = if let Some(refusal) = item.refusal {
                EmbeddedResourceGcDecision::Refuse(refusal)
            } else if item.supported_owners != 0 {
                EmbeddedResourceGcDecision::RetainReferenced {
                    supported_owner_count: item.supported_owners,
                }
            } else {
                EmbeddedResourceGcDecision::Delete
            };
            EmbeddedResourceGcPlanEntry {
                candidate: item.candidate,
                decision,
                archive_paths: item.archive_paths,
                manifest_paths: item.manifest_paths,
            }
        })
        .collect();
    EmbeddedResourceGcPlan { source, entries }
}

struct NormalizedCandidate {
    candidate: EmbeddedResourceGcCandidate,
    target: Option<Target>,
    refusal: Option<EmbeddedResourceGcRefusal>,
    archive_paths: Vec<String>,
    manifest_paths: Vec<String>,
    supported_owners: usize,
    generic_references: usize,
}

#[derive(Clone)]
enum Target {
    File(String),
    Directory(String),
}

impl Target {
    fn matches_reference(&self, path: &str) -> bool {
        match self {
            Self::File(candidate) => path == candidate,
            Self::Directory(root) => path == root.trim_end_matches('/') || path.starts_with(root),
        }
    }

    fn overlaps(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::File(left), Self::File(right)) => left == right,
            (Self::Directory(left), Self::Directory(right)) => {
                left.starts_with(right) || right.starts_with(left)
            },
            (Self::File(path), Self::Directory(root))
            | (Self::Directory(root), Self::File(path)) => path.starts_with(root),
        }
    }
}

impl NormalizedCandidate {
    fn new(candidate: EmbeddedResourceGcCandidate) -> Self {
        let target = normalize_candidate(&candidate).ok();
        let refusal = target
            .is_none()
            .then_some(EmbeddedResourceGcRefusal::UnsafePath);
        Self {
            candidate,
            target,
            refusal,
            archive_paths: Vec::new(),
            manifest_paths: Vec::new(),
            supported_owners: 0,
            generic_references: 0,
        }
    }
}

fn normalize_candidate(candidate: &EmbeddedResourceGcCandidate) -> Result<Target> {
    let raw = candidate.path();
    if raw.is_empty() || is_linked_href(raw) {
        return invalid("unsafe embedded-resource GC candidate path");
    }
    let normalized = resolve_package_path(raw.trim_end_matches('/'))?;
    if protected_path(&normalized) {
        return invalid("embedded-resource GC candidate targets a protected package path");
    }
    match candidate {
        EmbeddedResourceGcCandidate::PackageFile(_) if !raw.ends_with('/') => {
            Ok(Target::File(normalized))
        },
        EmbeddedResourceGcCandidate::PackageSubdocument(_) => {
            Ok(Target::Directory(format!("{normalized}/")))
        },
        EmbeddedResourceGcCandidate::PackageFile(_) => {
            invalid("embedded-resource GC file candidate is a directory")
        },
    }
}

fn protected_path(path: &str) -> bool {
    matches!(
        path,
        "mimetype" | ODF_CONTENT | ODF_STYLES | "meta.xml" | "settings.xml"
    ) || path == "META-INF"
        || path.starts_with("META-INF/")
}

fn duplicate_values(values: &[String]) -> BTreeSet<String> {
    values
        .windows(2)
        .filter(|pair| pair[0] == pair[1])
        .map(|pair| pair[0].clone())
        .collect()
}

fn mark_overlaps(candidates: &mut [NormalizedCandidate]) {
    for left in 0..candidates.len() {
        for right in left + 1..candidates.len() {
            if candidates[left]
                .target
                .as_ref()
                .zip(candidates[right].target.as_ref())
                .is_some_and(|(left, right)| left.overlaps(right))
            {
                candidates[left].refusal = Some(EmbeddedResourceGcRefusal::OverlappingCandidate);
                candidates[right].refusal = Some(EmbeddedResourceGcRefusal::OverlappingCandidate);
            }
        }
    }
}

fn classify_namespace(
    candidate: &mut NormalizedCandidate,
    files: &[String],
    manifest: &BTreeSet<String>,
    duplicates: &BTreeSet<String>,
) {
    let Some(target) = &candidate.target else {
        return;
    };
    match target {
        Target::File(path) => {
            if duplicates.contains(path)
                || files
                    .iter()
                    .any(|entry| entry.starts_with(&format!("{path}/")))
                || manifest
                    .iter()
                    .any(|entry| entry.starts_with(&format!("{path}/")))
            {
                candidate.refusal = Some(EmbeddedResourceGcRefusal::PackageNamespaceCollision);
            } else if !files.iter().any(|entry| entry == path) {
                candidate.refusal = Some(EmbeddedResourceGcRefusal::MissingArchiveEntry);
            } else if !manifest.contains(path) {
                candidate.refusal = Some(EmbeddedResourceGcRefusal::MissingManifestEntry);
            } else {
                candidate.archive_paths.push(path.clone());
                candidate.manifest_paths.push(path.clone());
            }
        },
        Target::Directory(root) => {
            if files
                .iter()
                .any(|entry| entry == root.trim_end_matches('/'))
                || manifest
                    .iter()
                    .any(|entry| entry == root.trim_end_matches('/'))
            {
                candidate.refusal = Some(EmbeddedResourceGcRefusal::PackageNamespaceCollision);
                return;
            }
            let archive_paths = files
                .iter()
                .filter(|entry| entry.starts_with(root))
                .cloned()
                .collect::<Vec<_>>();
            if archive_paths.is_empty() {
                candidate.refusal = Some(EmbeddedResourceGcRefusal::MissingArchiveEntry);
            } else if !manifest.contains(root) {
                candidate.refusal = Some(EmbeddedResourceGcRefusal::MissingManifestEntry);
            } else if !archive_paths
                .iter()
                .any(|entry| entry == &format!("{root}{ODF_CONTENT}"))
                || archive_paths.iter().any(|entry| duplicates.contains(entry))
            {
                candidate.refusal = Some(EmbeddedResourceGcRefusal::InvalidSubdocument);
            } else {
                candidate.archive_paths = archive_paths;
                candidate.manifest_paths = manifest
                    .iter()
                    .filter(|entry| entry.as_str() == root || entry.starts_with(root))
                    .cloned()
                    .collect();
            }
        },
    }
}

fn classify_references(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    candidates: &mut [NormalizedCandidate],
) -> Result<()> {
    let archive = package.package()?;
    let zip_metadata =
        soapberry_zip::office::ArchiveReader::new(package.as_bytes()).map_err(zip_error)?;
    for object in scan_objects(content, styles, &archive)? {
        if !matches!(object.part, Part::Content | Part::Styles) {
            continue;
        }
        if !matches!(object.kind, ObjectKind::Object | ObjectKind::ObjectOle) {
            continue;
        }
        let path = match object.source {
            ObjectSource::PackageFile { path, .. } => path,
            ObjectSource::PackageSubdocument { root_path, .. } => root_path,
            _ => continue,
        };
        add_supported_reference(candidates, &path);
    }
    for image in scan_images(content, styles, &archive)? {
        if !matches!(image.part, Part::Content | Part::Styles) {
            continue;
        }
        if let ImageSource::PackagePart { path, .. } = image.source {
            add_supported_reference(candidates, &path);
        }
    }

    scan_reference_xml(ODF_CONTENT, content, candidates)?;
    if let Some(styles) = styles {
        scan_reference_xml(ODF_STYLES, styles, candidates)?;
    }
    let mut total_xml = content.len().saturating_add(styles.map_or(0, str::len));
    for path in archive.files()? {
        if matches!(path.as_str(), ODF_CONTENT | ODF_STYLES | MANIFEST_PATH)
            || !(path.ends_with(".xml")
                || archive
                    .manifest()
                    .get_media_type(&path)
                    .is_some_and(|value| value.ends_with("/xml") || value.ends_with("+xml")))
        {
            continue;
        }
        let declared_size = usize::try_from(
            zip_metadata
                .metadata(&path)
                .map_err(zip_error)?
                .uncompressed_size(),
        )
        .map_err(|_error| {
            Error::InvalidFormat("ODT embedded-resource GC XML size overflows usize".to_string())
        })?;
        total_xml = total_xml.checked_add(declared_size).ok_or_else(|| {
            Error::InvalidFormat("ODT embedded-resource GC XML byte count overflow".to_string())
        })?;
        if total_xml > MAX_REFERENCE_XML_BYTES {
            return invalid("ODT embedded-resource GC reference XML exceeds its byte limit");
        }
        let bytes = archive.get_file(&path)?;
        if bytes.len() != declared_size {
            return invalid("ODT embedded-resource GC XML size disagrees with ZIP metadata");
        }
        let xml = std::str::from_utf8(&bytes).map_err(|_error| {
            Error::InvalidFormat(format!(
                "ODT embedded-resource GC XML part '{path}' is not UTF-8"
            ))
        })?;
        scan_reference_xml(&path, xml, candidates)?;
    }

    for candidate in candidates {
        if candidate.refusal.is_none() && candidate.generic_references > candidate.supported_owners
        {
            candidate.refusal = Some(EmbeddedResourceGcRefusal::UnknownReference {
                part: "package XML".to_string(),
            });
        }
    }
    Ok(())
}

fn add_supported_reference(candidates: &mut [NormalizedCandidate], path: &str) {
    for candidate in candidates {
        if candidate
            .target
            .as_ref()
            .is_some_and(|target| target.matches_reference(path))
        {
            candidate.supported_owners = candidate.supported_owners.saturating_add(1);
        }
    }
}

fn scan_reference_xml(part: &str, xml: &str, candidates: &mut [NormalizedCandidate]) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (_, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid ODT embedded-resource GC reference XML '{part}': {error}"
                ))
            })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                for raw in element.attributes() {
                    let attribute = raw.map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid ODT embedded-resource GC reference attribute in '{part}': {error}"
                        ))
                    })?;
                    let (attribute_namespace, attribute_local) =
                        reader.resolver().resolve_attribute(attribute.key);
                    let value = attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid ODT embedded-resource GC reference value in '{part}': {error}"
                            ))
                        })?;
                    if matches!(attribute_namespace, ResolveResult::Bound(Namespace(uri)) if uri == XML_NS)
                        && attribute_local.as_ref() == b"base"
                    {
                        for candidate in &mut *candidates {
                            if candidate.refusal.is_none()
                                && !reference_part_is_inside_candidate(candidate, part)
                            {
                                candidate.refusal =
                                    Some(EmbeddedResourceGcRefusal::UnknownReference {
                                        part: part.to_string(),
                                    });
                            }
                        }
                        continue;
                    }
                    if !is_linked_href(&value)
                        && let Ok(path) = resolve_package_path(&value)
                    {
                        add_generic_reference(candidates, &path, part);
                        if matches!(attribute_namespace, ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NS)
                            && attribute_local.as_ref() == b"href"
                            && let Some((directory, _name)) = part.rsplit_once('/')
                            && let Ok(relative_path) =
                                resolve_package_path(&format!("{directory}/{value}"))
                        {
                            add_generic_reference(candidates, &relative_path, part);
                        }
                    }
                }
            },
            Event::Text(value) => {
                let value = String::from_utf8_lossy(value.as_ref());
                for candidate in &mut *candidates {
                    if candidate.refusal.is_none()
                        && !reference_part_is_inside_candidate(candidate, part)
                        && candidate
                            .target
                            .as_ref()
                            .is_some_and(|target| text_mentions(target, &value))
                    {
                        candidate.refusal = Some(EmbeddedResourceGcRefusal::UnknownReference {
                            part: part.to_string(),
                        });
                    }
                }
            },
            Event::CData(value) => {
                let value = String::from_utf8_lossy(value.as_ref());
                for candidate in &mut *candidates {
                    if candidate.refusal.is_none()
                        && !reference_part_is_inside_candidate(candidate, part)
                        && candidate
                            .target
                            .as_ref()
                            .is_some_and(|target| text_mentions(target, &value))
                    {
                        candidate.refusal = Some(EmbeddedResourceGcRefusal::UnknownReference {
                            part: part.to_string(),
                        });
                    }
                }
            },
            Event::DocType(_) => {
                return invalid(format!(
                    "DTDs are not allowed in ODT embedded-resource GC reference XML '{part}'"
                ));
            },
            Event::Eof => break,
            Event::End(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn add_generic_reference(candidates: &mut [NormalizedCandidate], path: &str, part: &str) {
    for candidate in candidates {
        if candidate.refusal.is_none()
            && !reference_part_is_inside_candidate(candidate, part)
            && candidate
                .target
                .as_ref()
                .is_some_and(|target| target.matches_reference(path))
        {
            candidate.generic_references = candidate.generic_references.saturating_add(1);
            if candidate.generic_references > candidate.supported_owners {
                // Retain the precise first surface; a later supported scanner
                // count cannot turn an unknown grammar use into a known owner.
                candidate.refusal = Some(EmbeddedResourceGcRefusal::UnknownReference {
                    part: part.to_string(),
                });
            }
        }
    }
}

fn reference_part_is_inside_candidate(candidate: &NormalizedCandidate, part: &str) -> bool {
    matches!(
        candidate.target.as_ref(),
        Some(Target::Directory(root)) if part.starts_with(root)
    )
}

fn text_mentions(target: &Target, text: &str) -> bool {
    match target {
        Target::File(path) => text.contains(path),
        Target::Directory(root) => text.contains(root) || text.contains(root.trim_end_matches('/')),
    }
}

fn rewrite_package(
    source: &OwnedPackage,
    archive_paths: &BTreeSet<String>,
    manifest_paths: &BTreeSet<String>,
) -> Result<Vec<u8>> {
    let manifest = source.get_file(MANIFEST_PATH)?;
    let replacement = remove_manifest_records(&manifest, manifest_paths)?;
    let archive = ZipArchive::from_slice(source.as_bytes())
        .map_err(zip_error)?
        .into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).map_err(zip_error)?;
    let mut records = archive.entries(&mut buffer);
    let mut seen = BTreeSet::new();
    let mut replaced_manifest = false;
    let mut entries = Vec::new();
    entries
        .try_reserve_exact(index.entries().len())
        .map_err(|source| Error::Allocation {
            resource: "ODT embedded-resource GC ZIP table",
            source,
        })?;
    for (central_index, preserved) in index.entries().iter().enumerate() {
        let record = records.next_entry().map_err(zip_error)?.ok_or_else(|| {
            Error::InvalidFormat("ODT embedded-resource GC ZIP index disagreement".to_string())
        })?;
        let normalized = record.file_path().try_normalize().map_err(|error| {
            Error::InvalidFormat(format!("unsafe ODT embedded-resource GC ZIP path: {error}"))
        })?;
        let name = normalized.as_ref();
        if !seen.insert(name.to_string()) {
            return invalid("duplicate ODT package path during embedded-resource GC");
        }
        let original_local = checked_range(preserved.local_span(), source.as_bytes().len())?;
        if archive_paths.contains(name) {
            entries.push(RawRewriteEntry {
                central_index,
                original_start: original_local.start,
                local: None,
                central: None,
            });
            continue;
        }
        if name == MANIFEST_PATH {
            if record.is_dir() || replaced_manifest {
                return invalid("invalid ODT manifest member during embedded-resource GC");
            }
            let compression = match record.compression_method() {
                CompressionMethod::Store => CompressionMethod::Store,
                CompressionMethod::Deflate => CompressionMethod::Deflate,
                _ => return invalid("unsupported ODT manifest compression during resource GC"),
            };
            let generated = generated_manifest_member(&replacement, compression)?;
            entries.push(RawRewriteEntry {
                central_index,
                original_start: original_local.start,
                local: Some(RawLocal::Owned(generated.0)),
                central: Some(generated.1),
            });
            replaced_manifest = true;
        } else {
            entries.push(RawRewriteEntry {
                central_index,
                original_start: original_local.start,
                local: Some(RawLocal::Source(original_local)),
                central: Some(
                    source.as_bytes()
                        [checked_range(preserved.central_record(), source.as_bytes().len())?]
                    .to_vec(),
                ),
            });
        }
    }
    if records.next_entry().map_err(zip_error)?.is_some() || !replaced_manifest {
        return invalid("ODT embedded-resource GC ZIP traversal disagreement");
    }
    publish_raw_rewrite(source.as_bytes(), entries)
}

enum RawLocal {
    Source(Range<usize>),
    Owned(Vec<u8>),
}

struct RawRewriteEntry {
    central_index: usize,
    original_start: usize,
    local: Option<RawLocal>,
    central: Option<Vec<u8>>,
}

fn generated_manifest_member(
    manifest: &[u8],
    compression: CompressionMethod,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let mut writer = StreamingArchiveWriter::new();
    match compression {
        CompressionMethod::Store => writer.write_stored(MANIFEST_PATH, manifest),
        CompressionMethod::Deflate => writer.write_deflated(MANIFEST_PATH, manifest),
        _ => return invalid("unsupported ODT manifest compression during resource GC"),
    }
    .map_err(zip_error)?;
    let bytes = writer.finish_to_bytes().map_err(zip_error)?;
    let archive = ZipArchive::from_slice(&bytes)
        .map_err(zip_error)?
        .into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).map_err(zip_error)?;
    if index.entries().len() != 1 {
        return invalid("generated ODT manifest ZIP has an unexpected entry count");
    }
    let entry = &index.entries()[0];
    Ok((
        bytes[checked_range(entry.local_span(), bytes.len())?].to_vec(),
        bytes[checked_range(entry.central_record(), bytes.len())?].to_vec(),
    ))
}

fn publish_raw_rewrite(source: &[u8], entries: Vec<RawRewriteEntry>) -> Result<Vec<u8>> {
    let retained = entries.iter().filter(|entry| entry.local.is_some()).count();
    let entry_count = u16::try_from(retained)
        .map_err(|_error| Error::InvalidFormat("ODT resource GC ZIP64 promotion".to_string()))?;
    let mut local_order = entries
        .iter()
        .enumerate()
        .filter_map(|(index, entry)| entry.local.as_ref().map(|_| (entry.original_start, index)))
        .collect::<Vec<_>>();
    local_order.sort_by_key(|(start, _)| *start);
    let mut output = Vec::new();
    output
        .try_reserve(source.len())
        .map_err(|source| Error::Allocation {
            resource: "ODT embedded-resource GC output",
            source,
        })?;
    let mut offsets = BTreeMap::new();
    for (_, index) in local_order {
        let offset = u32::try_from(output.len()).map_err(|_error| {
            Error::InvalidFormat("ODT resource GC ZIP64 local offset".to_string())
        })?;
        offsets.insert(entries[index].central_index, offset);
        match entries[index].local.as_ref().ok_or_else(|| {
            Error::InvalidFormat("ODT resource GC local ordering disagreement".to_string())
        })? {
            RawLocal::Source(range) => output.extend_from_slice(&source[range.clone()]),
            RawLocal::Owned(bytes) => output.extend_from_slice(bytes),
        }
    }
    let central_start = u32::try_from(output.len()).map_err(|_error| {
        Error::InvalidFormat("ODT resource GC ZIP64 central offset".to_string())
    })?;
    for entry in &entries {
        let (Some(mut central), Some(offset)) = (
            entry.central.clone(),
            offsets.get(&entry.central_index).copied(),
        ) else {
            continue;
        };
        if central.len() < 46 {
            return invalid("truncated ODT resource GC central record");
        }
        central[42..46].copy_from_slice(&offset.to_le_bytes());
        output.extend_from_slice(&central);
    }
    let central_size = u32::try_from(output.len())
        .ok()
        .and_then(|end| end.checked_sub(central_start))
        .ok_or_else(|| Error::InvalidFormat("ODT resource GC central size overflow".to_string()))?;
    let comment = archive_comment(source)?;
    let comment_len = u16::try_from(comment.len())
        .map_err(|_error| Error::InvalidFormat("ODT ZIP comment too large".to_string()))?;
    output.extend_from_slice(&0x0605_4b50_u32.to_le_bytes());
    output.extend_from_slice(&0_u16.to_le_bytes());
    output.extend_from_slice(&0_u16.to_le_bytes());
    output.extend_from_slice(&entry_count.to_le_bytes());
    output.extend_from_slice(&entry_count.to_le_bytes());
    output.extend_from_slice(&central_size.to_le_bytes());
    output.extend_from_slice(&central_start.to_le_bytes());
    output.extend_from_slice(&comment_len.to_le_bytes());
    output.extend_from_slice(comment);
    Ok(output)
}

fn checked_range(range: Range<u64>, length: usize) -> Result<Range<usize>> {
    let start = usize::try_from(range.start)
        .map_err(|_error| Error::InvalidFormat("ODT ZIP range overflow".to_string()))?;
    let end = usize::try_from(range.end)
        .map_err(|_error| Error::InvalidFormat("ODT ZIP range overflow".to_string()))?;
    if start > end || end > length {
        return invalid("ODT ZIP range escapes source bytes");
    }
    Ok(start..end)
}

fn archive_comment(source: &[u8]) -> Result<&[u8]> {
    if source.len() < 22 {
        return invalid("ODT ZIP has no end record");
    }
    let lower = source.len().saturating_sub(65_557);
    for start in (lower..=source.len() - 22).rev() {
        if source[start..].starts_with(&0x0605_4b50_u32.to_le_bytes()) {
            let comment_len =
                usize::from(u16::from_le_bytes([source[start + 20], source[start + 21]]));
            if start.checked_add(22 + comment_len) == Some(source.len()) {
                return Ok(&source[start + 22..]);
            }
        }
    }
    invalid("ODT ZIP end record is invalid")
}

fn remove_manifest_records(source: &[u8], targets: &BTreeSet<String>) -> Result<Vec<u8>> {
    let xml = std::str::from_utf8(source)
        .map_err(|_error| Error::InvalidFormat("ODT manifest is not UTF-8".to_string()))?;
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut active: Option<(usize, usize, String)> = None;
    let mut depth = 0usize;
    let mut spans = BTreeMap::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| Error::InvalidFormat("ODT manifest position overflow".to_string()))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODT manifest during resource GC: {error}"))
            })?;
        let is_entry =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NS);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| Error::InvalidFormat("ODT manifest position overflow".to_string()))?;
        match event {
            Event::Start(element) => {
                if is_entry && element.local_name().as_ref() == b"file-entry" {
                    if active.is_some() {
                        return invalid("nested manifest file entries during resource GC");
                    }
                    active = Some((depth, start, manifest_full_path(&reader, &element)?));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODT manifest depth overflow".to_string())
                })?;
            },
            Event::Empty(element) if is_entry && element.local_name().as_ref() == b"file-entry" => {
                let path = manifest_full_path(&reader, &element)?;
                if targets.contains(&path) && spans.insert(path, start..end).is_some() {
                    return invalid("duplicate selected ODT manifest record");
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODT manifest stack underflow".to_string())
                })?;
                if is_entry
                    && element.local_name().as_ref() == b"file-entry"
                    && let Some((entry_depth, entry_start, path)) = active.take()
                {
                    if entry_depth != depth {
                        return invalid("ODT manifest file-entry depth disagreement");
                    }
                    if targets.contains(&path) && spans.insert(path, entry_start..end).is_some() {
                        return invalid("duplicate selected ODT manifest record");
                    }
                }
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in ODT manifests"),
            Event::Eof => break,
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if active.is_some() || spans.len() != targets.len() {
        return invalid("ODT embedded-resource GC manifest record set changed");
    }
    let mut ordered_spans = spans.into_values().collect::<Vec<_>>();
    ordered_spans.sort_by_key(|span| std::cmp::Reverse(span.start));
    let mut output = source.to_vec();
    for span in ordered_spans {
        output.drain(span);
    }
    Ok(output)
}

fn manifest_full_path(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<String> {
    let mut path = None;
    for raw in element.attributes() {
        let attribute = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODT manifest attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == MANIFEST_NS)
            && local.as_ref() == b"full-path"
        {
            if path.is_some() {
                return invalid("duplicate manifest:full-path attribute");
            }
            path = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid manifest:full-path value: {error}"))
                    })?
                    .into_owned(),
            );
        }
    }
    path.ok_or_else(|| Error::InvalidFormat("manifest file entry has no full path".to_string()))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

pub(crate) fn validate_candidate_path_bound(path: &str) -> Result<()> {
    if path.len() > MAX_EMBEDDED_RESOURCE_GC_PATH_BYTES {
        return invalid(format!(
            "ODT embedded-resource GC candidate path exceeds {MAX_EMBEDDED_RESOURCE_GC_PATH_BYTES} UTF-8 bytes"
        ));
    }
    Ok(())
}

fn zip_error(error: soapberry_zip::Error) -> Error {
    Error::InvalidFormat(format!("ODT embedded-resource GC ZIP error: {error}"))
}
