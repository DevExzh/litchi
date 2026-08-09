//! Immutable database package ownership.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::family::Package;
use litchi_odf_common::package::edit::Addition;
use quick_xml::{
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = litchi_odf_common::constants::ODF_DATABASE;
const BODY_MARKER: &str = "<";

struct State {
    package: Package,
}

/// An immutable, validated package snapshot.
#[derive(Clone)]
pub(crate) struct Snapshot(Arc<State>);

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

    pub(crate) fn file(&self, path: &str) -> Result<Vec<u8>> {
        self.0.package.package().get_file(path)
    }

    pub(crate) fn component_payload(
        &self,
        source_prefix: &str,
        destination_prefix: &str,
    ) -> Result<(Vec<Addition>, Vec<(String, String)>)> {
        let package = self.0.package.package().package()?;
        let mut additions = Vec::new();
        for path in package.files()? {
            if path.ends_with('/') {
                continue;
            }
            let Some(suffix) = path.strip_prefix(source_prefix) else {
                continue;
            };
            let media_type = package
                .manifest()
                .entries
                .get(&path)
                .map_or_else(String::new, |entry| entry.media_type.clone());
            additions.push(Addition {
                path: format!("{destination_prefix}{suffix}"),
                bytes: package.get_file(&path)?,
                media_type,
            });
        }
        let mut directories = package
            .manifest()
            .entries
            .iter()
            .filter_map(|(path, entry)| {
                path.strip_prefix(source_prefix).and_then(|suffix| {
                    path.ends_with('/').then(|| {
                        (
                            format!("{destination_prefix}{suffix}"),
                            entry.media_type.clone(),
                        )
                    })
                })
            })
            .collect::<Vec<_>>();
        additions.sort_by(|left, right| left.path.cmp(&right.path));
        directories.sort_by(|left, right| left.0.cmp(&right.0));
        Ok((additions, directories))
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
        if additions.is_empty() && directories.is_empty() {
            return self.rebuild_with_content(content);
        }
        crate::codec::validate(content)?;
        let spliced = super::splice::rebuild_content(self.0.package.package(), content)?;
        let intermediate = litchi_odf_common::core::OwnedPackage::from_bytes(spliced)?;
        let bytes = litchi_odf_common::package::edit::rebuild_package(
            &intermediate,
            content,
            additions,
            directories,
            Vec::<String>::new(),
            Vec::<String>::new(),
        )?;
        Self::from_bytes(bytes)
    }
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
