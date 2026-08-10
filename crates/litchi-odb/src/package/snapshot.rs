//! Immutable database package ownership.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::OwnedPackage;
use litchi_odf_common::core::family::Package;
use litchi_odf_common::package::edit::Addition;
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{
    collections::{BTreeMap, BTreeSet, VecDeque},
    path::Path,
    sync::Arc,
};

pub(crate) const MIMETYPE: &str = litchi_odf_common::constants::ODF_DATABASE;
const BODY_MARKER: &str = "<";

struct State {
    package: Package,
}

/// An immutable, validated package snapshot.
#[derive(Clone)]
pub(crate) struct Snapshot(Arc<State>);

#[derive(Default)]
pub(crate) struct PackageDelta {
    pub(crate) additions: Vec<Addition>,
    pub(crate) directories: Vec<(String, String)>,
    pub(crate) excluded_paths: Vec<String>,
}

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        Package::open(path, MIMETYPE, BODY_MARKER, "ODB").and_then(Self::validated)
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODB").and_then(Self::validated)
    }

    pub(crate) fn from_bytes_with_password(
        bytes: Vec<u8>,
        password: impl Into<String>,
    ) -> Result<Self> {
        Package::from_bytes_with_password(bytes, password, MIMETYPE, BODY_MARKER, "ODB")
            .and_then(Self::validated)
    }

    fn validated(package: Package) -> Result<Self> {
        crate::codec::validate(package.content_xml())?;
        Ok(Self(Arc::new(State { package })))
    }

    pub(crate) fn content_xml(&self) -> &str {
        self.0.package.content_xml()
    }

    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.0.package.styles_xml()
    }

    pub(crate) fn metadata(&self) -> Option<&Metadata> {
        self.0.package.metadata()
    }

    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.0.package.as_bytes()
    }

    pub(crate) fn files(&self) -> Result<Vec<String>> {
        self.0.package.files()
    }

    pub(crate) fn component_payload(
        &self,
        source_prefix: &str,
        destination_prefix: &str,
    ) -> Result<(Vec<Addition>, Vec<(String, String)>)> {
        let archive = self.0.package.package();
        let package = archive.package()?;
        let files = package.files()?.into_iter().collect::<BTreeSet<_>>();
        let mut additions = BTreeMap::<String, Addition>::new();
        let mut queue = VecDeque::new();
        let total_bytes = collect_component_tree(
            archive,
            &files,
            source_prefix,
            destination_prefix,
            &mut additions,
            &mut queue,
        )?;
        complete_component_dependencies(
            archive,
            &files,
            source_prefix,
            destination_prefix,
            &mut additions,
            queue,
            total_bytes,
        )?;
        let mut directories = package
            .manifest()
            .entries
            .iter()
            .filter_map(|(path, entry)| {
                path.strip_prefix(source_prefix)
                    .and_then(|suffix| {
                        path.ends_with('/').then(|| {
                            (
                                format!("{destination_prefix}{suffix}"),
                                entry.media_type.clone(),
                            )
                        })
                    })
                    .or_else(|| {
                        (path.ends_with('/')
                            && additions
                                .keys()
                                .any(|dependency| dependency.starts_with(path)))
                        .then(|| (path.clone(), entry.media_type.clone()))
                    })
            })
            .collect::<Vec<_>>();
        directories.sort_by(|left, right| left.0.cmp(&right.0));
        directories.dedup_by(|left, right| left.0 == right.0);
        if additions
            .len()
            .checked_add(directories.len())
            .is_none_or(|count| count > 4_096)
        {
            return Err(Error::InvalidFormat(
                "ODB component dependency closure exceeds the entry limit".to_string(),
            ));
        }
        Ok((additions.into_values().collect(), directories))
    }

    pub(crate) fn file_matches(&self, addition: &Addition) -> Result<Option<bool>> {
        let package = self.0.package.package().package()?;
        if !package.has_file(&addition.path) {
            return Ok(None);
        }
        let media_type = package
            .manifest()
            .entries
            .get(&addition.path)
            .map_or("", |entry| entry.media_type.as_str());
        Ok(Some(
            package.get_file(&addition.path)? == addition.bytes
                && media_type == addition.media_type,
        ))
    }

    pub(crate) fn directory_media_type(&self, path: &str) -> Result<Option<String>> {
        Ok(self
            .0
            .package
            .package()
            .package()?
            .manifest()
            .entries
            .get(path)
            .filter(|_| path.ends_with('/'))
            .map(|entry| entry.media_type.clone()))
    }

    pub(crate) fn payload_active_content_count(additions: &[Addition]) -> Result<usize> {
        const MAX_FINDINGS: usize = 16_384;
        let mut findings = Vec::new();
        for addition in additions {
            let member_kind = if addition.path.starts_with("Basic/") {
                Some(crate::ActiveContentKind::BasicMacro)
            } else if addition.path.starts_with("Scripts/") {
                Some(crate::ActiveContentKind::Script)
            } else {
                None
            };
            if let Some(kind) = member_kind {
                push_finding(
                    &mut findings,
                    crate::ActiveContentEntry::package_member(kind, addition.path.clone()),
                    MAX_FINDINGS,
                )?;
            }
            if !is_xml_member(&addition.path, &addition.media_type) {
                continue;
            }
            let xml = std::str::from_utf8(&addition.bytes).map_err(|error| {
                Error::InvalidFormat(format!(
                    "ODB component payload XML member '{}' is not UTF-8: {error}",
                    addition.path
                ))
            })?;
            scan_active_xml(&addition.path, xml, &mut findings, MAX_FINDINGS)?;
        }
        Ok(findings.len())
    }

    pub(crate) fn payload_transfer_support(
        additions: &[Addition],
    ) -> crate::ComponentTransferSupport {
        let requires_provenance = additions.iter().any(|addition| {
            is_xml_member(&addition.path, &addition.media_type)
                && litchi_odf_common::compact_xml::validate(&addition.bytes).is_err()
        });
        if requires_provenance {
            crate::ComponentTransferSupport::Refused(
                crate::ComponentTransferRefusal::FormattedXmlRequiresSourceProvenance,
            )
        } else {
            crate::ComponentTransferSupport::Supported
        }
    }

    pub(crate) fn contains_package_prefix(&self, prefix: &str) -> Result<bool> {
        let package = self.0.package.package().package()?;
        Ok(package.files()?.iter().any(|path| path.starts_with(prefix))
            || package
                .manifest()
                .entries
                .keys()
                .any(|path| path.starts_with(prefix)))
    }

    pub(crate) fn active_content(&self) -> Result<crate::ActiveContentInventory> {
        const MAX_SCAN_BYTES: usize = 32 * 1024 * 1024;
        const MAX_FINDINGS: usize = 16_384;
        let package = self.0.package.package();
        let mut files = package.files()?;
        files.sort();
        let mut findings = Vec::new();
        let mut scanned = 0usize;
        for path in files {
            let member_kind = if path.starts_with("Basic/") {
                Some(crate::ActiveContentKind::BasicMacro)
            } else if path.starts_with("Scripts/") {
                Some(crate::ActiveContentKind::Script)
            } else {
                None
            };
            if let Some(kind) = member_kind {
                push_finding(
                    &mut findings,
                    crate::ActiveContentEntry::package_member(kind, path.clone()),
                    MAX_FINDINGS,
                )?;
            }
            let scan_xml = path == "content.xml"
                || ((path.starts_with("forms/")
                    || path.starts_with("Forms/")
                    || path.starts_with("reports/")
                    || path.starts_with("Reports/")
                    || path.starts_with("Basic/")
                    || path.starts_with("Scripts/"))
                    && has_xml_extension(&path));
            if !scan_xml {
                continue;
            }
            let bytes = package.get_file(&path)?;
            scanned = scanned.checked_add(bytes.len()).ok_or_else(|| {
                Error::InvalidFormat("ODB active-content scan size overflow".to_string())
            })?;
            if scanned > MAX_SCAN_BYTES {
                return Err(Error::InvalidFormat(
                    "ODB active-content XML exceeds the scan limit".to_string(),
                ));
            }
            let xml = std::str::from_utf8(&bytes).map_err(|error| {
                Error::InvalidFormat(format!(
                    "ODB active-content XML member '{path}' is not UTF-8: {error}"
                ))
            })?;
            scan_active_xml(&path, xml, &mut findings, MAX_FINDINGS)?;
        }
        Ok(crate::ActiveContentInventory::new(findings))
    }

    pub(crate) fn delta_to(&self, target: &Self) -> Result<PackageDelta> {
        let source = self.0.package.package().package()?;
        let target = target.0.package.package().package()?;
        let source_files = source.files()?.into_iter().collect::<BTreeSet<_>>();
        let target_files = target.files()?.into_iter().collect::<BTreeSet<_>>();
        let mut delta = PackageDelta::default();
        for path in &target_files {
            if publication_control_path(path) || path.ends_with('/') {
                continue;
            }
            let bytes = target.get_file(path)?;
            let media_type = target
                .manifest()
                .entries
                .get(path)
                .map_or_else(String::new, |entry| entry.media_type.clone());
            let source_media_type = source
                .manifest()
                .entries
                .get(path)
                .map_or("", |entry| entry.media_type.as_str());
            let changed = !source_files.contains(path)
                || source.get_file(path)? != bytes
                || source_media_type != media_type.as_str();
            if changed {
                delta.additions.push(Addition {
                    path: path.clone(),
                    bytes,
                    media_type,
                });
            }
        }
        for path in source_files.difference(&target_files) {
            if !publication_control_path(path) {
                delta.excluded_paths.push(path.clone());
            }
        }
        let source_manifest = source.manifest();
        let target_manifest = target.manifest();
        for (path, entry) in &target_manifest.entries {
            if path == "/" || !path.ends_with('/') {
                continue;
            }
            let changed = source_manifest
                .entries
                .get(path)
                .is_none_or(|source| source.media_type != entry.media_type);
            if changed {
                delta
                    .directories
                    .push((path.clone(), entry.media_type.clone()));
            }
        }
        for path in source_manifest.entries.keys() {
            if path != "/" && path.ends_with('/') && !target_manifest.entries.contains_key(path) {
                delta.excluded_paths.push(path.clone());
            }
        }
        delta
            .additions
            .sort_by(|left, right| left.path.cmp(&right.path));
        delta
            .directories
            .sort_by(|left, right| left.0.cmp(&right.0));
        delta.excluded_paths.sort();
        delta.excluded_paths.dedup();
        Ok(delta)
    }

    pub(crate) fn protection_status(&self) -> Result<crate::ProtectionStatus> {
        let files = self.files()?;
        let signed = files.iter().any(|path| {
            let lower = path.to_ascii_lowercase();
            lower.starts_with("meta-inf/") && lower.contains("signatures")
        });
        let encrypted = self
            .0
            .package
            .package()
            .package()?
            .manifest()
            .has_encrypted_entries();
        Ok(crate::ProtectionStatus::new(signed, encrypted))
    }

    pub(crate) fn digital_signatures(
        &self,
    ) -> Result<litchi_odf_common::signature::DigitalSignatures> {
        self.0.package.package().digital_signatures()
    }

    pub(crate) fn verify_document_signatures(
        &self,
    ) -> Result<Vec<litchi_odf_common::signature::SignatureVerification>> {
        self.0.package.package().verify_document_signatures()
    }

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.0) {
            Ok(state) => state.package.into_bytes(),
            Err(state) => state.package.as_bytes().to_vec(),
        }
    }

    pub(crate) fn rebuild_with_content(&self, content: &str) -> Result<Self> {
        // Opened producer documents are edited by byte-splicing only the
        // selected XML range.  Formatting whitespace in the unchanged source
        // is therefore lossless input, not generated output that needs the
        // fresh-authoring compactness gate.
        crate::codec::validate(content)?;
        let bytes = super::splice::rebuild_content(self.0.package.package(), content)?;
        Self::from_bytes(bytes)
    }

    pub(crate) fn rebuild_with_content_and_additions(
        &self,
        content: &str,
        additions: Vec<Addition>,
        directories: Vec<(String, String)>,
    ) -> Result<Self> {
        self.rebuild_with_content_and_mutations(content, additions, directories, Vec::new())
    }

    pub(crate) fn rebuild_with_content_and_mutations(
        &self,
        content: &str,
        additions: Vec<Addition>,
        directories: Vec<(String, String)>,
        excluded_paths: Vec<String>,
    ) -> Result<Self> {
        if additions.is_empty() && directories.is_empty() && excluded_paths.is_empty() {
            return self.rebuild_with_content(content);
        }
        crate::codec::validate(content)?;
        let spliced = super::splice::rebuild_content(self.0.package.package(), content)?;
        let intermediate = OwnedPackage::from_bytes(spliced)?;
        let bytes = litchi_odf_common::package::edit::rebuild_package(
            &intermediate,
            content,
            additions,
            directories,
            excluded_paths,
            Vec::<String>::new(),
        )?;
        Self::from_bytes(bytes)
    }
}

fn collect_component_tree(
    archive: &OwnedPackage,
    files: &BTreeSet<String>,
    source_prefix: &str,
    destination_prefix: &str,
    additions: &mut BTreeMap<String, Addition>,
    queue: &mut VecDeque<(String, String)>,
) -> Result<usize> {
    const MAX_DEPENDENCIES: usize = 4_096;
    const MAX_DEPENDENCY_BYTES: usize = 64 * 1024 * 1024;
    let package = archive.package()?;
    let mut total_bytes = 0usize;
    for source in files {
        let Some(suffix) = source.strip_prefix(source_prefix) else {
            continue;
        };
        if source.ends_with('/') {
            continue;
        }
        if additions.len() >= MAX_DEPENDENCIES {
            return Err(Error::InvalidFormat(
                "ODB component dependency closure exceeds the entry limit".to_string(),
            ));
        }
        let bytes = package.get_file(source)?;
        total_bytes = checked_dependency_size(total_bytes, bytes.len(), MAX_DEPENDENCY_BYTES)?;
        let destination = format!("{destination_prefix}{suffix}");
        let media_type = package
            .manifest()
            .entries
            .get(source)
            .map_or_else(String::new, |entry| entry.media_type.clone());
        let scan_xml = is_xml_member(source, &media_type);
        additions.insert(
            destination.clone(),
            Addition {
                path: destination,
                bytes,
                media_type,
            },
        );
        if scan_xml {
            queue.push_back((source.clone(), format!("{destination_prefix}{suffix}")));
        }
    }
    Ok(total_bytes)
}

fn complete_component_dependencies(
    archive: &OwnedPackage,
    files: &BTreeSet<String>,
    source_prefix: &str,
    destination_prefix: &str,
    additions: &mut BTreeMap<String, Addition>,
    mut queue: VecDeque<(String, String)>,
    mut total_bytes: usize,
) -> Result<()> {
    const MAX_DEPENDENCIES: usize = 4_096;
    const MAX_DEPENDENCY_BYTES: usize = 64 * 1024 * 1024;
    let package = archive.package()?;
    let mut visited = BTreeSet::new();
    while let Some((owner, destination_owner)) = queue.pop_front() {
        if !visited.insert((owner.clone(), destination_owner.clone())) {
            continue;
        }
        if visited.len() > MAX_DEPENDENCIES {
            return Err(Error::InvalidFormat(
                "ODB component dependency closure exceeds the entry limit".to_string(),
            ));
        }
        let owner_bytes = package.get_file(&owner)?;
        let xml = std::str::from_utf8(&owner_bytes).map_err(|error| {
            Error::InvalidFormat(format!(
                "ODB component dependency XML member '{owner}' is not UTF-8: {error}"
            ))
        })?;
        for (href, dependency) in linked_package_paths(&owner, xml)? {
            let destination_dependency = resolve_package_href(&destination_owner, &href)?;
            let expected_destination = dependency.strip_prefix(source_prefix).map_or_else(
                || dependency.clone(),
                |suffix| format!("{destination_prefix}{suffix}"),
            );
            if destination_dependency.as_deref() != Some(expected_destination.as_str()) {
                return Err(Error::InvalidFormat(format!(
                    "ODB component dependency href in '{owner}' changes meaning after relocation"
                )));
            }
            if dependency.starts_with(source_prefix) {
                continue;
            }
            if !files.contains(&dependency) {
                let directory = format!("{dependency}/");
                if package.manifest().entries.contains_key(&directory) {
                    continue;
                }
                return Err(Error::InvalidFormat(format!(
                    "ODB component dependency '{dependency}' referenced by '{owner}' is missing"
                )));
            }
            let bytes = package.get_file(&dependency)?;
            let media_type = package
                .manifest()
                .entries
                .get(&dependency)
                .map_or_else(String::new, |entry| entry.media_type.clone());
            if let Some(existing) = additions.get(&dependency) {
                if existing.bytes != bytes || existing.media_type != media_type {
                    return Err(Error::InvalidFormat(format!(
                        "ODB component dependency '{dependency}' conflicts with relocated payload"
                    )));
                }
                continue;
            }
            if additions.len() >= MAX_DEPENDENCIES {
                return Err(Error::InvalidFormat(
                    "ODB component dependency closure exceeds the entry limit".to_string(),
                ));
            }
            total_bytes = checked_dependency_size(total_bytes, bytes.len(), MAX_DEPENDENCY_BYTES)?;
            let scan_xml = is_xml_member(&dependency, &media_type);
            additions.insert(
                dependency.clone(),
                Addition {
                    path: dependency.clone(),
                    bytes,
                    media_type,
                },
            );
            if scan_xml {
                queue.push_back((dependency.clone(), dependency));
            }
        }
    }
    Ok(())
}

fn publication_control_path(path: &str) -> bool {
    matches!(path, "mimetype" | "content.xml" | "META-INF/manifest.xml")
}

fn push_finding(
    findings: &mut Vec<crate::ActiveContentEntry>,
    finding: crate::ActiveContentEntry,
    limit: usize,
) -> Result<()> {
    if findings.len() >= limit {
        return Err(Error::InvalidFormat(
            "ODB active-content inventory exceeds the finding limit".to_string(),
        ));
    }
    findings
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "ODB active-content inventory",
            source,
        })?;
    findings.push(finding);
    Ok(())
}

fn scan_active_xml(
    path: &str,
    xml: &str,
    findings: &mut Vec<crate::ActiveContentEntry>,
    limit: usize,
) -> Result<()> {
    const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const SCRIPT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
    const FORM: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:form:1.0";
    const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
    const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid ODB active-content XML member '{path}': {error}"
                ))
            })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let local = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(|_error| {
                        Error::InvalidFormat(
                            "ODB active-content declaration name is not UTF-8".to_string(),
                        )
                    })?
                    .to_owned();
                let namespace = match resolved {
                    ResolveResult::Bound(Namespace(uri)) => Some(uri),
                    ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
                };
                let kind = match (namespace, local.as_str()) {
                    (Some(SCRIPT | PRESENTATION), "event-listener") => {
                        Some(crate::ActiveContentKind::EventListener)
                    },
                    (Some(SCRIPT), "script") | (Some(OFFICE), "scripts") => {
                        Some(crate::ActiveContentKind::Script)
                    },
                    (Some(SCRIPT), "execute-macro") => Some(crate::ActiveContentKind::Action),
                    (Some(FORM), _) => Some(crate::ActiveContentKind::FormControl),
                    (Some(TABLE), "dde-link" | "dde-source") => {
                        Some(crate::ActiveContentKind::DdeLink)
                    },
                    (Some(DRAW), "object" | "object-ole" | "plugin" | "applet") => {
                        Some(crate::ActiveContentKind::EmbeddedObject)
                    },
                    _ => None,
                };
                if let Some(kind) = kind {
                    push_finding(
                        findings,
                        crate::ActiveContentEntry::declaration(
                            kind,
                            path.to_owned(),
                            local.clone(),
                        ),
                        limit,
                    )?;
                }
                let mut has_action = false;
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid ODB active-content XML attribute in '{path}': {error}"
                        ))
                    })?;
                    has_action |= attribute.key.local_name().as_ref() == b"action";
                }
                if has_action {
                    push_finding(
                        findings,
                        crate::ActiveContentEntry::declaration(
                            crate::ActiveContentKind::Action,
                            path.to_owned(),
                            local,
                        ),
                        limit,
                    )?;
                }
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not allowed in ODB active-content XML".to_string(),
                ));
            },
            Event::Eof => break,
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn has_xml_extension(path: &str) -> bool {
    path.rsplit_once('.')
        .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("xml"))
}

fn is_xml_member(path: &str, media_type: &str) -> bool {
    has_xml_extension(path)
        || media_type.eq_ignore_ascii_case("text/xml")
        || media_type.eq_ignore_ascii_case("application/xml")
        || media_type
            .rsplit_once('+')
            .is_some_and(|(_, suffix)| suffix.eq_ignore_ascii_case("xml"))
}

fn checked_dependency_size(current: usize, added: usize, limit: usize) -> Result<usize> {
    let total = current.checked_add(added).ok_or_else(|| {
        Error::InvalidFormat("ODB component dependency size overflow".to_string())
    })?;
    if total > limit {
        return Err(Error::InvalidFormat(
            "ODB component dependency closure exceeds the byte limit".to_string(),
        ));
    }
    Ok(total)
}

fn linked_package_paths(owner: &str, xml: &str) -> Result<Vec<(String, String)>> {
    const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
    const MAX_LINKS: usize = 4_096;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut paths = Vec::new();
    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid ODB component dependency XML member '{owner}': {error}"
                ))
            })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if matches!(resolved, ResolveResult::Unknown(_)) {
                    return Err(Error::InvalidFormat(format!(
                        "unresolved namespace in ODB component dependency XML member '{owner}'"
                    )));
                }
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid ODB component dependency attribute in '{owner}': {error}"
                        ))
                    })?;
                    let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
                    if matches!(namespace, ResolveResult::Unknown(_)) {
                        return Err(Error::InvalidFormat(format!(
                            "unresolved attribute namespace in ODB component dependency XML member '{owner}'"
                        )));
                    }
                    if !matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == XLINK)
                        || local.as_ref() != b"href"
                    {
                        continue;
                    }
                    let href = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Explicit1_0,
                            reader.decoder(),
                        )
                        .map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid ODB component dependency href in '{owner}': {error}"
                            ))
                        })?;
                    if let Some(path) = resolve_package_href(owner, &href)? {
                        if paths.len() >= MAX_LINKS {
                            return Err(Error::InvalidFormat(
                                "ODB component XML exceeds the link limit".to_string(),
                            ));
                        }
                        paths.try_reserve(1).map_err(|source| Error::Allocation {
                            resource: "ODB component dependency links",
                            source,
                        })?;
                        paths.push((href.into_owned(), path));
                    }
                }
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DOCTYPE is not allowed in ODB component dependency XML".to_string(),
                ));
            },
            Event::Eof => break,
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    paths.sort();
    paths.dedup();
    Ok(paths)
}

fn resolve_package_href(owner: &str, href: &str) -> Result<Option<String>> {
    let href = href.split_once('#').map_or(href, |(path, _)| path);
    if href.is_empty() {
        return Ok(None);
    }
    let first_segment = href.split('/').next().unwrap_or("");
    if first_segment.contains(':') || href.starts_with("//") {
        return Ok(None);
    }
    if href
        .chars()
        .any(|character| matches!(character, '\\' | '?'))
    {
        return Err(Error::InvalidFormat(
            "ODB component dependency href is not a safe package path".to_string(),
        ));
    }
    let mut segments = if href.starts_with('/') {
        Vec::new()
    } else {
        owner.rsplit_once('/').map_or(Vec::new(), |(parent, _)| {
            parent.split('/').map(str::to_owned).collect()
        })
    };
    for segment in href.trim_start_matches('/').split('/') {
        match segment {
            "" | "." => {},
            ".." => {
                if segments.pop().is_none() {
                    return Err(Error::InvalidFormat(
                        "ODB component dependency escapes the package root".to_string(),
                    ));
                }
            },
            value => segments.push(value.to_owned()),
        }
    }
    if segments.is_empty() {
        Ok(None)
    } else {
        Ok(Some(segments.join("/")))
    }
}
