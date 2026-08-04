//! Atomic authoring of inert embedded objects, OLE payloads, and images.

use crate::core::OwnedPackage;
use crate::embedded_chart::{
    Addition, EmbeddedChartHost, ObjectSpan, insert_at_host, locate_objects, rebuild_package,
    splice, unused_object_root,
};
use crate::{
    ImagePart, ImageSource, OdfEmbeddedObjectKind, OdfEmbeddedObjectPart, OdfEmbeddedObjectSource,
    OdfInlineObjectRoot, constants,
};
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use litchi_core::{Error, Result};
use litchi_odf_common::package::{is_linked_href, resolve_package_path};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_URI: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE_URI: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const XLINK_URI: &str = "http://www.w3.org/1999/xlink";
const MAX_FILES: usize = 16_384;
const MAX_BYTES: usize = 64 * 1024 * 1024;
const MAX_INLINE_BYTES: usize = 16 * 1024 * 1024;

/// Element kind for an authored embedded package resource.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OdfEmbeddedResourceKind {
    Object,
    ObjectOle,
    Image,
}

/// One file in an authored embedded OpenDocument subdocument.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfEmbeddedResourceFile {
    /// Path relative to the subdocument root, such as `content.xml`.
    pub path: String,
    pub bytes: Vec<u8>,
    /// Exact manifest media type for this file.
    pub media_type: String,
}

/// Storage for a newly authored inert resource.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OdfEmbeddedResourceSource {
    /// An external target retained as an inert link. It is never fetched or executed.
    Linked { href: String },
    /// One opaque package file.
    PackageFile {
        bytes: Vec<u8>,
        media_type: String,
        preferred_path: Option<String>,
    },
    /// An embedded OpenDocument package rooted at one package directory.
    PackageSubdocument {
        files: Vec<OdfEmbeddedResourceFile>,
        media_type: String,
        preferred_root: Option<String>,
    },
    /// A complete inline `office:document` or `math:math` element.
    InlineXml {
        root: OdfInlineObjectRoot,
        xml: String,
    },
    /// Base64-encoded into an `office:binary-data` child.
    InlineBinary {
        bytes: Vec<u8>,
        media_type: Option<String>,
    },
}

/// Typed input for adding or replacing an embedded package resource.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfEmbeddedResource {
    pub kind: OdfEmbeddedResourceKind,
    pub source: OdfEmbeddedResourceSource,
    pub frame_name: Option<String>,
    pub xml_id: Option<String>,
    /// Optional OLE class identifier. Only valid for `ObjectOle`.
    pub class_id: Option<String>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum ResourceTarget {
    Object,
    Image,
}

struct BuiltResource {
    xml: String,
    additions: Vec<Addition>,
    directories: Vec<(String, String)>,
}

#[derive(Clone)]
enum StoredLocation {
    File(String),
    Directory(String),
}

pub(crate) fn add(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    host: EmbeddedChartHost<'_>,
    resource: &OdfEmbeddedResource,
) -> Result<(Vec<u8>, usize)> {
    let target = target(resource.kind);
    let index = target_count(package, content, styles, target)?;
    let built = build_resource(package, resource, None)?;
    let updated = insert_at_host(content, host, &built.xml)?;
    let bytes = rebuild_package(
        package,
        &updated,
        built.additions,
        built.directories,
        Vec::new(),
        Vec::new(),
    )?;
    Ok((bytes, index))
}

pub(crate) fn replace(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
    target: ResourceTarget,
    resource: &OdfEmbeddedResource,
) -> Result<Vec<u8>> {
    if self::target(resource.kind) != target {
        return invalid("replacement resource kind does not match the selected API");
    }
    let (span, old) = selected(package, content, styles, index, target)?;
    let built = build_resource(package, resource, old.as_ref())?;
    let updated = splice(content, span.start, span.end, &built.xml)?;
    let (excluded_paths, excluded_prefixes) =
        cleanup(package, &updated, styles, old.as_ref(), &built)?;
    rebuild_package(
        package,
        &updated,
        built.additions,
        built.directories,
        excluded_paths,
        excluded_prefixes,
    )
}

pub(crate) fn remove(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
    target: ResourceTarget,
) -> Result<Vec<u8>> {
    let (span, old) = selected(package, content, styles, index, target)?;
    let updated = splice(content, span.start, span.end, "")?;
    let empty = BuiltResource {
        xml: String::new(),
        additions: Vec::new(),
        directories: Vec::new(),
    };
    let (excluded_paths, excluded_prefixes) =
        cleanup(package, &updated, styles, old.as_ref(), &empty)?;
    rebuild_package(
        package,
        &updated,
        Vec::new(),
        Vec::new(),
        excluded_paths,
        excluded_prefixes,
    )
}

pub(crate) fn reorder(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    from: usize,
    to: usize,
    target: ResourceTarget,
) -> Result<Vec<u8>> {
    let spans = selected_spans(package, content, styles, target)?;
    let first = spans.get(from).ok_or_else(|| bounds(from, spans.len()))?;
    let second = spans.get(to).ok_or_else(|| bounds(to, spans.len()))?;
    if from == to {
        return rebuild_package(
            package,
            content,
            Vec::new(),
            Vec::new(),
            Vec::new(),
            Vec::new(),
        );
    }
    let mut updated = String::with_capacity(content.len());
    if first.start < second.start {
        updated.push_str(&content[..first.start]);
        updated.push_str(&content[first.end..second.end]);
        updated.push_str(&content[first.start..first.end]);
        updated.push_str(&content[second.end..]);
    } else {
        updated.push_str(&content[..second.start]);
        updated.push_str(&content[first.start..first.end]);
        updated.push_str(&content[second.start..first.start]);
        updated.push_str(&content[first.end..]);
    }
    rebuild_package(
        package,
        &updated,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
    )
}

fn build_resource(
    package: &OwnedPackage,
    resource: &OdfEmbeddedResource,
    replacing: Option<&StoredLocation>,
) -> Result<BuiltResource> {
    validate_metadata(resource)?;
    let mut additions = Vec::new();
    let mut directories = Vec::new();
    let body = match &resource.source {
        OdfEmbeddedResourceSource::Linked { href } => {
            if !is_linked_href(href) {
                return invalid(
                    "linked resources must use an external, fragment, or otherwise inert target",
                );
            }
            href_element(resource, href, None)?
        },
        OdfEmbeddedResourceSource::PackageFile {
            bytes,
            media_type,
            preferred_path,
        } => {
            validate_payload(bytes.len())?;
            validate_media_type(resource.kind, media_type, false)?;
            let path = match preferred_path {
                Some(path) => validate_available_path(package, path, replacing)?,
                None => unused_file_path(package, resource.kind, media_type)?,
            };
            additions.push(Addition {
                path: path.clone(),
                bytes: bytes.clone(),
                media_type: media_type.clone(),
            });
            href_element(resource, &path, Some(media_type))?
        },
        OdfEmbeddedResourceSource::PackageSubdocument {
            files,
            media_type,
            preferred_root,
        } => {
            if resource.kind != OdfEmbeddedResourceKind::Object {
                return invalid("only draw:object supports package subdocuments");
            }
            validate_media_type(resource.kind, media_type, true)?;
            if files.is_empty() || files.len() > MAX_FILES {
                return invalid("embedded subdocument file count is outside the allowed range");
            }
            let root = match preferred_root {
                Some(root) => validate_available_root(package, root, replacing)?,
                None => unused_object_root(package)?,
            };
            let mut names = HashSet::new();
            let mut total = 0usize;
            let mut has_content = false;
            for file in files {
                let relative = validate_relative_path(&file.path)?;
                if !names.insert(relative.clone()) {
                    return invalid(format!("duplicate embedded subdocument path '{relative}'"));
                }
                if relative == constants::ODF_CONTENT {
                    has_content = true;
                    if file.media_type != "text/xml" {
                        return invalid(
                            "embedded content.xml must have manifest media type text/xml",
                        );
                    }
                    validate_xml_document(&file.bytes, "embedded content.xml")?;
                }
                validate_manifest_media_type(&file.media_type)?;
                total = total.checked_add(file.bytes.len()).ok_or_else(|| {
                    Error::InvalidFormat("embedded resource size overflow".to_string())
                })?;
                if total > MAX_BYTES {
                    return invalid("embedded subdocument exceeds size limit");
                }
                additions.push(Addition {
                    path: format!("{root}{relative}"),
                    bytes: file.bytes.clone(),
                    media_type: file.media_type.clone(),
                });
            }
            if !has_content {
                return invalid("embedded subdocument is missing content.xml");
            }
            directories.push((root.clone(), media_type.clone()));
            href_element(resource, root.trim_end_matches('/'), None)?
        },
        OdfEmbeddedResourceSource::InlineXml { root, xml } => {
            if resource.kind != OdfEmbeddedResourceKind::Object {
                return invalid("inline XML is only valid for draw:object");
            }
            validate_inline_xml(*root, xml)?;
            element(resource, xml, None)?
        },
        OdfEmbeddedResourceSource::InlineBinary { bytes, media_type } => {
            if resource.kind == OdfEmbeddedResourceKind::Object {
                return invalid("draw:object inline payloads must be XML");
            }
            validate_payload(bytes.len())?;
            if let Some(media_type) = media_type {
                validate_media_type(resource.kind, media_type, false)?;
            }
            let encoded = BASE64_STANDARD.encode(bytes);
            element(
                resource,
                &format!("<office:binary-data>{encoded}</office:binary-data>"),
                media_type.as_deref(),
            )?
        },
    };
    let mut frame = format!(
        "<draw:frame xmlns:draw=\"{DRAW_URI}\" xmlns:office=\"{OFFICE_URI}\" xmlns:xlink=\"{XLINK_URI}\""
    );
    attribute(&mut frame, "draw:name", resource.frame_name.as_deref())?;
    frame.push('>');
    frame.push_str(&body);
    frame.push_str("</draw:frame>");
    Ok(BuiltResource {
        xml: frame,
        additions,
        directories,
    })
}

fn href_element(
    resource: &OdfEmbeddedResource,
    href: &str,
    media_type: Option<&str>,
) -> Result<String> {
    let mut attrs = String::new();
    attribute(&mut attrs, "xlink:href", Some(href))?;
    attrs.push_str(" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"");
    element(resource, "", media_type).map(|element| {
        let point = element.find('>').unwrap_or(element.len());
        format!("{}{}{}", &element[..point], attrs, &element[point..])
    })
}

fn element(
    resource: &OdfEmbeddedResource,
    payload: &str,
    media_type: Option<&str>,
) -> Result<String> {
    let local = match resource.kind {
        OdfEmbeddedResourceKind::Object => "object",
        OdfEmbeddedResourceKind::ObjectOle => "object-ole",
        OdfEmbeddedResourceKind::Image => "image",
    };
    let mut out = format!("<draw:{local}");
    attribute(&mut out, "xml:id", resource.xml_id.as_deref())?;
    attribute(&mut out, "draw:class-id", resource.class_id.as_deref())?;
    if resource.kind == OdfEmbeddedResourceKind::Image {
        attribute(&mut out, "draw:mime-type", media_type)?;
    }
    out.push('>');
    out.push_str(payload);
    out.push_str("</draw:");
    out.push_str(local);
    out.push('>');
    Ok(out)
}

fn validate_metadata(resource: &OdfEmbeddedResource) -> Result<()> {
    if resource.kind != OdfEmbeddedResourceKind::ObjectOle && resource.class_id.is_some() {
        return invalid("draw:class-id is only valid for ObjectOle resources");
    }
    for value in [
        resource.frame_name.as_deref(),
        resource.xml_id.as_deref(),
        resource.class_id.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        validate_xml_chars(value)?;
    }
    Ok(())
}

fn validate_inline_xml(root: OdfInlineObjectRoot, xml: &str) -> Result<()> {
    if xml.len() > MAX_INLINE_BYTES {
        return invalid("inline XML exceeds size limit");
    }
    if xml.contains("<!DOCTYPE")
        || xml.contains("<!ENTITY")
        || xml.contains("<?")
        || xml.contains("office:scripts")
        || xml.contains("script:event-listener")
        || xml.contains("urn:oasis:names:tc:opendocument:xmlns:script")
    {
        return invalid("inline XML contains active or prohibited markup");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid inline XML: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let expected = match root {
                    OdfInlineObjectRoot::OpenDocument => b"document".as_slice(),
                    OdfInlineObjectRoot::MathMl => b"math".as_slice(),
                };
                if element.local_name().as_ref() != expected {
                    return invalid("inline XML root does not match its declared kind");
                }
                if root == OdfInlineObjectRoot::OpenDocument {
                    let valid_ns = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == OFFICE_NS);
                    if !valid_ns {
                        return invalid("inline OpenDocument root has the wrong namespace");
                    }
                    let mut mime = None;
                    for attr in element.attributes() {
                        let attr = attr.map_err(|error| {
                            Error::InvalidFormat(format!("invalid inline XML attribute: {error}"))
                        })?;
                        if attr.key.local_name().as_ref() == b"mimetype" {
                            mime = Some(
                                attr.decoded_and_normalized_value(
                                    XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )
                                .map_err(|error| {
                                    Error::InvalidFormat(format!(
                                        "invalid inline XML attribute: {error}"
                                    ))
                                })?
                                .into_owned(),
                            );
                        }
                    }
                    let mime = mime.ok_or_else(|| {
                        Error::InvalidFormat(
                            "inline OpenDocument root is missing office:mimetype".to_string(),
                        )
                    })?;
                    if !constants::is_odf_mime_type(&mime) {
                        return invalid("inline OpenDocument has an invalid media type");
                    }
                }
                return Ok(());
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("inline XML contains a prohibited declaration");
            },
            Event::Text(value) => {
                let bytes: &[u8] = value.as_ref();
                if !bytes.iter().all(u8::is_ascii_whitespace) {
                    return invalid("invalid content before inline XML root");
                }
            },
            Event::Comment(_) | Event::Decl(_) => {},
            Event::Eof => return invalid("inline XML has no root element"),
            _ => return invalid("invalid content before inline XML root"),
        }
        buffer.clear();
    }
}

fn validate_xml_document(bytes: &[u8], label: &str) -> Result<()> {
    if bytes.len() > MAX_INLINE_BYTES {
        return invalid(format!("{label} exceeds size limit"));
    }
    let text = std::str::from_utf8(bytes)
        .map_err(|_| Error::InvalidFormat(format!("{label} is not UTF-8")))?;
    if text.contains("<!DOCTYPE") || text.contains("<!ENTITY") {
        return invalid(format!("{label} contains a DTD"));
    }
    let mut reader = NsReader::from_str(text);
    let mut buffer = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid {label}: {error}")))?
        {
            Event::Eof => return Ok(()),
            _ => buffer.clear(),
        }
    }
}

fn validate_media_type(
    kind: OdfEmbeddedResourceKind,
    media_type: &str,
    subdocument: bool,
) -> Result<()> {
    validate_manifest_media_type(media_type)?;
    let lower = media_type.to_ascii_lowercase();
    if lower.contains("javascript")
        || lower.contains("ecmascript")
        || lower.contains("x-executable")
        || lower.contains("x-msdownload")
        || lower.contains("portable-executable")
        || lower.contains("x-sharedlib")
        || lower.contains("x-shellscript")
    {
        return invalid("active or executable embedded media types are prohibited");
    }
    if kind == OdfEmbeddedResourceKind::Image && !lower.starts_with("image/") {
        return invalid("draw:image package payloads require an image media type");
    }
    if subdocument && !constants::is_odf_mime_type(media_type) {
        return invalid("embedded subdocuments require an OpenDocument media type");
    }
    Ok(())
}

fn validate_manifest_media_type(value: &str) -> Result<()> {
    if value.is_empty()
        || value.trim() != value
        || value.chars().any(char::is_control)
        || !value.contains('/')
        || value.contains(';')
    {
        return invalid("invalid manifest media type");
    }
    Ok(())
}

fn validate_payload(len: usize) -> Result<()> {
    if len > MAX_BYTES {
        invalid("embedded payload exceeds size limit")
    } else {
        Ok(())
    }
}

fn validate_relative_path(path: &str) -> Result<String> {
    let path = resolve_package_path(path)?;
    if path.is_empty() || path.ends_with('/') || path.starts_with("META-INF/") || path == "mimetype"
    {
        return invalid("invalid embedded subdocument file path");
    }
    Ok(path)
}

fn validate_available_path(
    package: &OwnedPackage,
    path: &str,
    replacing: Option<&StoredLocation>,
) -> Result<String> {
    let path = resolve_package_path(path)?;
    if path.is_empty() || path.ends_with('/') || protected_path(&path) {
        return invalid("invalid embedded package file path");
    }
    let allowed = matches!(replacing, Some(StoredLocation::File(old)) if old == &path);
    if package.has_file(&path)? && !allowed {
        return invalid(format!("package path '{path}' already exists"));
    }
    Ok(path)
}

fn validate_available_root(
    package: &OwnedPackage,
    root: &str,
    replacing: Option<&StoredLocation>,
) -> Result<String> {
    let mut root = resolve_package_path(root)?;
    if root.is_empty() || protected_path(&root) {
        return invalid("invalid embedded subdocument root");
    }
    root.push('/');
    let allowed = matches!(replacing, Some(StoredLocation::Directory(old)) if old == &root);
    if !allowed && package.files()?.iter().any(|path| path.starts_with(&root)) {
        return invalid(format!("package root '{root}' already exists"));
    }
    Ok(root)
}

fn protected_path(path: &str) -> bool {
    path == "mimetype"
        || path == constants::ODF_CONTENT
        || path == constants::ODF_STYLES
        || path == constants::ODF_META
        || path.starts_with("META-INF/")
}

fn unused_file_path(
    package: &OwnedPackage,
    kind: OdfEmbeddedResourceKind,
    media_type: &str,
) -> Result<String> {
    let extension = if kind == OdfEmbeddedResourceKind::Image {
        match media_type.to_ascii_lowercase().as_str() {
            "image/png" => "png",
            "image/jpeg" => "jpg",
            "image/gif" => "gif",
            "image/svg+xml" => "svg",
            "image/webp" => "webp",
            _ => "bin",
        }
    } else {
        "bin"
    };
    for index in 1..=100_000usize {
        let path = if kind == OdfEmbeddedResourceKind::Image {
            format!("Pictures/Image_{index}.{extension}")
        } else {
            format!("Object_{index}.{extension}")
        };
        if !package.has_file(&path)? {
            return Ok(path);
        }
    }
    invalid("no collision-free embedded package path is available")
}

fn selected(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    index: usize,
    target: ResourceTarget,
) -> Result<(ObjectSpan, Option<StoredLocation>)> {
    let spans = selected_spans(package, content, styles, target)?;
    let span = spans
        .get(index)
        .cloned()
        .ok_or_else(|| bounds(index, spans.len()))?;
    let location = selected_locations(package, content, styles, target)?
        .get(index)
        .cloned()
        .flatten();
    Ok((span, location))
}

fn selected_spans(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    target: ResourceTarget,
) -> Result<Vec<ObjectSpan>> {
    match target {
        ResourceTarget::Object => {
            let all_spans = locate_objects(content)?;
            let objects = scan_objects(package, content, styles)?;
            let content_objects: Vec<_> = objects
                .iter()
                .filter(|object| object.part == OdfEmbeddedObjectPart::Content)
                .collect();
            if all_spans.len() != content_objects.len() {
                return invalid("embedded-object XML scan disagreement");
            }
            Ok(all_spans
                .into_iter()
                .zip(content_objects)
                .filter_map(|(span, object)| {
                    matches!(
                        object.kind,
                        OdfEmbeddedObjectKind::Object | OdfEmbeddedObjectKind::ObjectOle
                    )
                    .then_some(span)
                })
                .collect())
        },
        ResourceTarget::Image => locate_images(content),
    }
}

fn selected_locations(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    target: ResourceTarget,
) -> Result<Vec<Option<StoredLocation>>> {
    match target {
        ResourceTarget::Object => Ok(scan_objects(package, content, styles)?
            .into_iter()
            .filter(|object| object.part == OdfEmbeddedObjectPart::Content)
            .filter(|object| {
                matches!(
                    object.kind,
                    OdfEmbeddedObjectKind::Object | OdfEmbeddedObjectKind::ObjectOle
                )
            })
            .map(|object| match object.source {
                OdfEmbeddedObjectSource::PackageFile { path, .. } => {
                    Some(StoredLocation::File(path))
                },
                OdfEmbeddedObjectSource::PackageSubdocument { root_path, .. } => {
                    Some(StoredLocation::Directory(root_path))
                },
                _ => None,
            })
            .collect()),
        ResourceTarget::Image => Ok(scan_images(package, content, styles)?
            .into_iter()
            .filter(|image| image.part == ImagePart::Content)
            .map(|image| match image.source {
                ImageSource::PackagePart { path, .. } => Some(StoredLocation::File(path)),
                _ => None,
            })
            .collect()),
    }
}

fn target_count(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    target: ResourceTarget,
) -> Result<usize> {
    Ok(selected_spans(package, content, styles, target)?.len())
}

fn cleanup(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
    old: Option<&StoredLocation>,
    built: &BuiltResource,
) -> Result<(Vec<String>, Vec<String>)> {
    let mut excluded_paths: Vec<String> = built
        .additions
        .iter()
        .filter_map(|addition| {
            package
                .has_file(&addition.path)
                .ok()
                .filter(|exists| *exists)
                .map(|_| addition.path.clone())
        })
        .collect();
    let mut excluded_prefixes: Vec<String> = built
        .directories
        .iter()
        .filter_map(|(root, _)| {
            package
                .files()
                .ok()
                .filter(|files| files.iter().any(|path| path.starts_with(root)))
                .map(|_| root.clone())
        })
        .collect();
    let Some(old) = old else {
        return Ok((excluded_paths, excluded_prefixes));
    };
    let additions: HashSet<&str> = built
        .additions
        .iter()
        .map(|item| item.path.as_str())
        .collect();
    let has_file = |path: &str| additions.contains(path) || package.has_file(path).unwrap_or(false);
    let media_type = |path: &str| {
        built
            .additions
            .iter()
            .find(|item| item.path == path)
            .map(|item| item.media_type.clone())
            .or_else(|| {
                package
                    .package()
                    .ok()?
                    .manifest()
                    .get_media_type(path)
                    .map(str::to_string)
            })
    };
    let objects =
        crate::embedded_object::scan_packaged_objects(content, styles, has_file, media_type)?;
    let images = crate::media::scan_packaged_images(content, styles, has_file, media_type)?;
    let referenced = match old {
        StoredLocation::File(path) => objects.iter().any(|object| matches!(&object.source, OdfEmbeddedObjectSource::PackageFile { path: candidate, .. } if candidate == path))
            || images.iter().any(|image| matches!(&image.source, ImageSource::PackagePart { path: candidate, .. } if candidate == path)),
        StoredLocation::Directory(root) => objects.iter().any(|object| matches!(&object.source, OdfEmbeddedObjectSource::PackageSubdocument { root_path, .. } if root_path == root)),
    };
    if referenced {
        return Ok((excluded_paths, excluded_prefixes));
    }
    match old {
        StoredLocation::File(path) => {
            if !excluded_paths.contains(path) {
                excluded_paths.push(path.clone());
            }
        },
        StoredLocation::Directory(root) => {
            if !excluded_prefixes.contains(root) {
                excluded_prefixes.push(root.clone());
            }
        },
    }
    Ok((excluded_paths, excluded_prefixes))
}

fn scan_objects(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
) -> Result<Vec<crate::OdfEmbeddedObject>> {
    let archive = package.package()?;
    crate::embedded_object::scan_packaged_objects(
        content,
        styles,
        |path| archive.has_file(path),
        |path| archive.manifest().get_media_type(path).map(str::to_string),
    )
}

fn scan_images(
    package: &OwnedPackage,
    content: &str,
    styles: Option<&str>,
) -> Result<Vec<crate::Image>> {
    let archive = package.package()?;
    crate::media::scan_packaged_images(
        content,
        styles,
        |path| archive.has_file(path),
        |path| archive.manifest().get_media_type(path).map(str::to_string),
    )
}

fn locate_images(xml: &str) -> Result<Vec<ObjectSpan>> {
    struct Active {
        depth: usize,
        start: usize,
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<Active> = None;
    let mut spans = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("XML position overflow".to_string()))?;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid image host XML: {error}")))?;
        let is_draw =
            matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == DRAW_NS);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("XML position overflow".to_string()))?;
        match event {
            Event::Start(element) => {
                if is_draw && element.local_name().as_ref() == b"image" {
                    if active.is_some() {
                        return invalid("nested draw:image elements are not supported");
                    }
                    active = Some(Active { depth, start });
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("XML depth overflow".to_string()))?;
            },
            Event::Empty(element) if is_draw && element.local_name().as_ref() == b"image" => spans
                .push(ObjectSpan {
                    start,
                    end,
                    inline_payload: None,
                }),
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("XML depth underflow".to_string()))?;
                if is_draw
                    && element.local_name().as_ref() == b"image"
                    && active.as_ref().is_some_and(|item| item.depth == depth)
                {
                    let item = active.take().expect("active image");
                    spans.push(ObjectSpan {
                        start: item.start,
                        end,
                        inline_payload: None,
                    });
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if active.is_some() {
        return invalid("unterminated draw:image element");
    }
    Ok(spans)
}

fn target(kind: OdfEmbeddedResourceKind) -> ResourceTarget {
    match kind {
        OdfEmbeddedResourceKind::Image => ResourceTarget::Image,
        _ => ResourceTarget::Object,
    }
}

fn attribute(out: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    validate_xml_chars(value)?;
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(character),
        }
    }
    out.push('"');
    Ok(())
}

fn validate_xml_chars(value: &str) -> Result<()> {
    if value.chars().any(|character| matches!(character as u32, 0..=8 | 11 | 12 | 14..=31 | 0xD800..=0xDFFF | 0xFFFE | 0xFFFF)) {
        invalid("value contains a character forbidden by XML 1.0")
    } else { Ok(()) }
}

fn bounds(index: usize, len: usize) -> Error {
    Error::InvalidFormat(format!(
        "embedded resource index {index} is out of range for {len} resources"
    ))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
