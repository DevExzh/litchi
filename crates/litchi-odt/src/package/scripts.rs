//! Inert package authoring for document scripts and script resources.
//!
//! This module never resolves a URI, loads an external resource, or executes a payload.

use crate::core::OwnedPackage;
use crate::document_scripts::{EventListener, ScriptBinding, Scripts, parse_scripts};
use crate::package::charts::{Addition, rebuild_package};
use litchi_core::{Error, Result};
use litchi_odf_common::XmlStreamLimits;
use litchi_odf_common::package::resolve_package_path;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::collections::HashSet;

const MAX_RESOURCE_COUNT: usize = 16_384;
const MAX_RESOURCE_BYTES: usize = 16 * 1024 * 1024;
const MAX_TOTAL_RESOURCE_BYTES: usize = 64 * 1024 * 1024;
const MAX_VALUE_BYTES: usize = 64 * 1024;
const MAX_XML_DEPTH: usize = 256;
const MAX_XML_EVENTS: usize = 1_000_000;

/// The inert role of a package-contained script resource.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScriptResourceKind {
    BasicLibrary,
    BasicModule,
    Dialog,
    Opaque,
}

/// One discovered package-contained script resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScriptResource {
    pub kind: ScriptResourceKind,
    /// Safe package-relative location.
    pub path: String,
    /// Exact manifest media type.
    pub media_type: String,
    /// Inert bytes which are never executed.
    pub bytes: Vec<u8>,
}

/// Authored bytes and metadata for a package-contained script resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScriptResourceSpec {
    pub kind: ScriptResourceKind,
    /// Optional requested location; a collision-safe location is allocated when absent.
    pub preferred_path: Option<String>,
    pub media_type: String,
    pub bytes: Vec<u8>,
}

pub(crate) fn document_scripts(content: &str, limits: &XmlStreamLimits) -> Result<Option<Scripts>> {
    crate::document_scripts::package::read(content, limits)
}

pub(crate) fn resources(package: &OwnedPackage) -> Result<Vec<ScriptResource>> {
    let archive = package.package()?;
    let mut paths: Vec<String> = archive
        .files()?
        .into_iter()
        .filter(|path| classify_path(path).is_some())
        .collect();
    paths.sort();
    if paths.len() > MAX_RESOURCE_COUNT {
        return invalid("script resource count exceeds package limit");
    }
    let mut total = 0usize;
    let mut output = Vec::with_capacity(paths.len());
    for path in paths {
        let kind = classify_path(&path).ok_or_else(|| {
            Error::InvalidFormat(format!("invalid filtered script resource path '{path}'"))
        })?;
        let entry = archive.manifest().entries.get(&path).ok_or_else(|| {
            Error::InvalidFormat(format!("script resource '{path}' has no manifest entry"))
        })?;
        let bytes = archive.get_file(&path)?;
        validate_resource(kind, &entry.media_type, &bytes, true)?;
        total = total
            .checked_add(bytes.len())
            .ok_or_else(|| Error::InvalidFormat("script resource size overflow".to_string()))?;
        if total > MAX_TOTAL_RESOURCE_BYTES {
            return invalid("script resources exceed aggregate package limit");
        }
        output.push(ScriptResource {
            kind,
            path,
            media_type: entry.media_type.clone(),
            bytes,
        });
    }
    Ok(output)
}

pub(crate) fn find_resource(package: &OwnedPackage, path: &str) -> Result<Option<ScriptResource>> {
    let path = safe_script_path(path, None)?;
    Ok(resources(package)?
        .into_iter()
        .find(|resource| resource.path == path))
}

pub(crate) fn add_resource(
    package: &OwnedPackage,
    content: &str,
    resource: &ScriptResourceSpec,
) -> Result<(Vec<u8>, String)> {
    validate_resource(resource.kind, &resource.media_type, &resource.bytes, false)?;
    let path = match resource.preferred_path.as_deref() {
        Some(path) => {
            let path = safe_script_path(path, Some(resource.kind))?;
            ensure_available(package, &path, None)?;
            path
        },
        None => allocate_path(package, resource.kind)?,
    };
    let directories = missing_parent_directories(package, &path)?;
    let addition = Addition {
        path: path.clone(),
        bytes: resource.bytes.clone(),
        media_type: resource.media_type.clone(),
    };
    let bytes = rebuild_package(
        package,
        content,
        vec![addition],
        directories,
        Vec::new(),
        Vec::new(),
    )?;
    Ok((bytes, path))
}

pub(crate) fn replace_resource(
    package: &OwnedPackage,
    content: &str,
    path: &str,
    resource: &ScriptResourceSpec,
) -> Result<Vec<u8>> {
    let path = safe_script_path(path, Some(resource.kind))?;
    if let Some(preferred) = resource.preferred_path.as_deref()
        && safe_script_path(preferred, Some(resource.kind))? != path
    {
        return invalid("replacement script resource cannot change package location");
    }
    ensure_available(package, &path, Some(&path))?;
    validate_resource(resource.kind, &resource.media_type, &resource.bytes, false)?;
    let addition = Addition {
        path: path.clone(),
        bytes: resource.bytes.clone(),
        media_type: resource.media_type.clone(),
    };
    rebuild_package(
        package,
        content,
        vec![addition],
        Vec::new(),
        vec![path],
        Vec::new(),
    )
}

pub(crate) fn remove_resource(
    package: &OwnedPackage,
    content: &str,
    path: &str,
    limits: &XmlStreamLimits,
) -> Result<Vec<u8>> {
    let path = safe_script_path(path, None)?;
    if find_resource(package, &path)?.is_none() {
        return invalid(format!("script resource '{path}' was not found"));
    }
    if resource_is_referenced(content, &path, limits)? {
        return invalid(format!(
            "script resource '{path}' is still referenced by an event listener"
        ));
    }
    rebuild_package(
        package,
        content,
        Vec::new(),
        Vec::new(),
        vec![path],
        Vec::new(),
    )
}

fn validate_inert_href(href: &str) -> Result<()> {
    if href.is_empty() || href.len() > MAX_VALUE_BYTES || href.contains('\0') {
        return invalid("invalid script listener href");
    }
    let lower = href.to_ascii_lowercase();
    if ["javascript:", "vbscript:", "data:", "file:"]
        .iter()
        .any(|scheme| lower.starts_with(scheme))
    {
        return invalid("executable or local-file script URI is prohibited");
    }
    if href.starts_with('#') || href.starts_with("//") || uri_scheme(href).is_some() {
        return Ok(());
    }
    resolve_package_path(href).map(|_| ())?;
    Ok(())
}

fn uri_scheme(value: &str) -> Option<&str> {
    let colon = value.find(':')?;
    let boundary = value.find(['/', '?', '#']).unwrap_or(value.len());
    if colon >= boundary {
        return None;
    }
    let scheme = &value[..colon];
    (!scheme.is_empty()
        && scheme.as_bytes()[0].is_ascii_alphabetic()
        && scheme
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'-' | b'.')))
    .then_some(scheme)
}

fn resource_is_referenced(content: &str, path: &str, limits: &XmlStreamLimits) -> Result<bool> {
    let Some(scripts) = parse_scripts(content, limits)? else {
        return Ok(false);
    };
    for listener in scripts.event_listeners {
        if let EventListener::Script(listener) = listener
            && let ScriptBinding::Linked { href } = listener.binding
        {
            validate_inert_href(&href)?;
            if !href.starts_with('#')
                && !href.starts_with("//")
                && uri_scheme(&href).is_none()
                && resolve_package_path(&href)? == path
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn classify_path(path: &str) -> Option<ScriptResourceKind> {
    if path.ends_with('/') {
        return None;
    }
    if path.starts_with("Basic/") {
        if path.ends_with("/script-lc.xml")
            || path.ends_with("/script-lb.xml")
            || path == "Basic/script-lc.xml"
        {
            Some(ScriptResourceKind::BasicLibrary)
        } else if path.ends_with(".xml") {
            Some(ScriptResourceKind::BasicModule)
        } else {
            Some(ScriptResourceKind::Opaque)
        }
    } else if path.starts_with("Dialogs/") || path.starts_with("Dialog/") {
        Some(ScriptResourceKind::Dialog)
    } else if path.starts_with("Scripts/") {
        Some(ScriptResourceKind::Opaque)
    } else {
        None
    }
}

fn validate_resource(
    kind: ScriptResourceKind,
    media_type: &str,
    bytes: &[u8],
    allow_legacy_doctype: bool,
) -> Result<()> {
    if bytes.len() > MAX_RESOURCE_BYTES {
        return invalid("script resource exceeds size limit");
    }
    if media_type.len() > MAX_VALUE_BYTES || media_type.contains(['\r', '\n', '\0']) {
        return invalid("invalid script resource media type");
    }
    if matches!(
        kind,
        ScriptResourceKind::BasicLibrary
            | ScriptResourceKind::BasicModule
            | ScriptResourceKind::Dialog
    ) {
        if !matches!(media_type, "text/xml" | "application/xml" | "") {
            return invalid("XML script resource has a non-XML media type");
        }
        validate_inert_xml(bytes, allow_legacy_doctype)?;
    }
    Ok(())
}

fn validate_inert_xml(bytes: &[u8], allow_legacy_doctype: bool) -> Result<()> {
    std::str::from_utf8(bytes)
        .map_err(|_error| Error::InvalidFormat("script XML is not UTF-8".to_string()))
        .map(|_| ())?;
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut events = 0usize;
    loop {
        events += 1;
        if events > MAX_XML_EVENTS {
            return invalid("script XML event limit exceeded");
        }
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid inert script XML: {error}")))?
        {
            Event::Start(_) => {
                if depth == 0 {
                    roots += 1;
                }
                depth += 1;
                if depth > MAX_XML_DEPTH {
                    return invalid("script XML depth limit exceeded");
                }
            },
            Event::Empty(_) if depth == 0 => roots += 1,
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("script XML depth underflow".to_string())
                })?;
            },
            Event::DocType(value) => {
                let declaration = value.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid script XML DTD: {error}"))
                })?;
                let known_legacy = matches!(
                    declaration.as_ref(),
                    "script:module PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"module.dtd\""
                        | "library:library PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"library.dtd\""
                        | "library:libraries PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"libraries.dtd\""
                );
                if !allow_legacy_doctype || !known_legacy {
                    return invalid("DTD declarations are prohibited in script XML");
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(|error| {
                    Error::InvalidFormat(format!("invalid script XML entity: {error}"))
                })?;
                if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "quot" | "apos") {
                    return invalid("custom entity references are prohibited in script XML");
                }
            },
            Event::PI(_) => return invalid("processing instructions are prohibited in script XML"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || roots != 1 {
        return invalid("script XML must contain exactly one root element");
    }
    Ok(())
}

fn safe_script_path(path: &str, expected: Option<ScriptResourceKind>) -> Result<String> {
    if path.len() > MAX_VALUE_BYTES {
        return invalid("script resource path exceeds limit");
    }
    let path = resolve_package_path(path)?;
    if path.is_empty()
        || path.ends_with('/')
        || path == "mimetype"
        || path.starts_with("META-INF/")
        || matches!(
            path.as_str(),
            "content.xml" | "styles.xml" | "meta.xml" | "settings.xml"
        )
    {
        return invalid("unsafe script resource package path");
    }
    let actual = classify_path(&path).ok_or_else(|| {
        Error::InvalidFormat(
            "script resource must be stored under Basic/, Dialogs/, Dialog/, or Scripts/"
                .to_string(),
        )
    })?;
    if let Some(expected) = expected
        && actual != expected
    {
        return invalid("script resource kind does not match package location");
    }
    Ok(path)
}

fn ensure_available(package: &OwnedPackage, path: &str, replacing: Option<&str>) -> Result<()> {
    let archive = package.package()?;
    let exists = archive.has_file(path) || archive.manifest().entries.contains_key(path);
    if exists && replacing != Some(path) {
        return invalid(format!("package path '{path}' already exists"));
    }
    if replacing == Some(path) && !exists {
        return invalid(format!("script resource '{path}' was not found"));
    }
    Ok(())
}

fn allocate_path(package: &OwnedPackage, kind: ScriptResourceKind) -> Result<String> {
    let archive = package.package()?;
    let occupied: HashSet<String> = archive
        .files()?
        .into_iter()
        .chain(archive.manifest().entries.keys().cloned())
        .collect();
    for index in 1..=100_000usize {
        let candidate = match kind {
            ScriptResourceKind::BasicLibrary => format!("Basic/Library_{index}/script-lb.xml"),
            ScriptResourceKind::BasicModule => format!("Basic/Library_1/Module_{index}.xml"),
            ScriptResourceKind::Dialog => format!("Dialogs/Dialog_{index}.xml"),
            ScriptResourceKind::Opaque => format!("Scripts/Script_{index}.bin"),
        };
        if !occupied.contains(&candidate) {
            return Ok(candidate);
        }
    }
    invalid("no collision-free script resource path is available")
}

fn missing_parent_directories(package: &OwnedPackage, path: &str) -> Result<Vec<(String, String)>> {
    let archive = package.package()?;
    let mut output = Vec::new();
    for (index, byte) in path.bytes().enumerate() {
        if byte == b'/' {
            let directory = &path[..=index];
            if !archive.manifest().entries.contains_key(directory) {
                output.push((directory.to_string(), String::new()));
            }
        }
    }
    Ok(output)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

macro_rules! script_facade_methods {
    () => {
        pub fn document_scripts(&self) -> litchi_core::Result<Option<crate::Scripts>> {
            crate::package::scripts::document_scripts(
                self.state.content.xml_content(),
                &self.state.xml_limits,
            )
        }
        pub fn script_resources(&self) -> litchi_core::Result<Vec<crate::ScriptResource>> {
            crate::package::scripts::resources(&self.state.package)
        }
        pub fn find_script_resource(
            &self,
            path: &str,
        ) -> litchi_core::Result<Option<crate::ScriptResource>> {
            crate::package::scripts::find_resource(&self.state.package, path)
        }
    };
}
pub(crate) use script_facade_methods;
