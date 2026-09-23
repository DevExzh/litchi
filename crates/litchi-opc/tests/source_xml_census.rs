//! Change 0750: every XML member of every OPC package and every loose OOXML
//! XML part in the repository's `test-data` passes the source publication
//! audit, except a pinned list of members that are not well-formed XML.
//!
//! A member is audited when publication would audit it:
//! `xml_minifier::audit::package::is_xml_part(name, media_type)`, with the
//! media type taken from the package's own content-type map. A new refusal
//! of a fixture here means either a real producer writes a spelling the audit
//! should accept, or a fixture is not well-formed; the pinned list says which
//! members are known to be the latter.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "a census fails at the first unexpected fixture"
)]

use std::collections::{BTreeMap, HashMap};
use std::path::{Path, PathBuf};

use litchi_opc::phys_pkg::PhysPkgReader;
use quick_xml::events::Event;
use xml_minifier::audit::{self, Limits, package};

/// Members the source audit refuses, with the refusal, as `(file, member,
/// error)`. `member` is empty for a loose XML part.
const REFUSED: &[(&str, &str, &str)] = &[
    // Not UTF-8: a UTF-16 custom XML item and two embedded packages that
    // declare an XML media type.
    (
        "libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx",
        "customXml/item3.xml",
        "XML is not UTF-8 at byte 0",
    ),
    (
        "libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx",
        "word/afchunk2.docx",
        "XML is not UTF-8 at byte 14",
    ),
    (
        "libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk.docx",
        "word/afchunk2.docx",
        "XML is not UTF-8 at byte 16",
    ),
    // A deliberately mismatched end tag.
    (
        "ooxml/pptx/hyperlinks/malformed-slide.xml",
        "",
        "malformed XML at byte 268: ill-formed document: expected `</a:hlinkClick>`, but `</p:sld>` was found",
    ),
    // Change 0750: slide fragments whose prefixes the enclosing slide
    // declares, and a deliberately undeclared entity.
    (
        "ooxml/pptx/backgrounds/solid.xml",
        "",
        "malformed XML at byte 1: undeclared namespace prefix",
    ),
    (
        "ooxml/pptx/placeholders/invalid-placeholder.xml",
        "",
        "malformed XML at byte 221: reference to an undeclared entity",
    ),
    (
        "ooxml/pptx/transitions/p14_ripple.xml",
        "",
        "malformed XML at byte 196: undeclared namespace prefix",
    ),
    // Deliberate entity-expansion fixtures, refused for their DOCTYPE.
    (
        "poi/test-data/openxml4j/CorePropertiesHasEntities.ooxml",
        "docProps/core.xml",
        "DTD and DOCTYPE are not allowed at byte 57",
    ),
    (
        "poi/test-data/openxml4j/PackageRelsHasEntities.ooxml",
        "_rels/.rels",
        "DTD and DOCTYPE are not allowed at byte 57",
    ),
];

fn files(directory: &Path, found: &mut Vec<PathBuf>) {
    let mut entries: Vec<_> = std::fs::read_dir(directory)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .collect();
    entries.sort();
    for path in entries {
        if path.is_dir() {
            files(&path, found);
        } else {
            found.push(path);
        }
    }
}

/// The package's content-type map: lower-cased extension defaults and
/// lower-cased part-name overrides.
fn content_types(xml: &[u8]) -> (HashMap<String, String>, HashMap<String, String>) {
    let mut defaults = HashMap::new();
    let mut overrides = HashMap::new();
    let mut reader = quick_xml::Reader::from_reader(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Empty(tag) | Event::Start(tag)) => {
                let mut key = None;
                let mut value = None;
                for attribute in tag.attributes().flatten() {
                    let text = String::from_utf8_lossy(&attribute.value).into_owned();
                    match attribute.key.as_ref() {
                        b"Extension" | b"PartName" => key = Some(text.to_lowercase()),
                        b"ContentType" => value = Some(text),
                        _ => {},
                    }
                }
                if let (Some(key), Some(value)) = (key, value) {
                    match tag.local_name().as_ref() {
                        b"Default" => defaults.insert(key, value),
                        b"Override" => overrides.insert(key, value),
                        _ => None,
                    };
                }
            },
            Ok(Event::Eof) | Err(_) => break,
            Ok(_) => {},
        }
    }
    (defaults, overrides)
}

#[test]
fn every_well_formed_fixture_member_passes_the_source_audit() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut paths = Vec::new();
    files(&root, &mut paths);
    let mut refused = BTreeMap::new();
    let (mut packages, mut members, mut loose) = (0usize, 0usize, 0usize);
    for path in paths {
        let relative = path
            .strip_prefix(&root)
            .unwrap()
            .to_string_lossy()
            .into_owned();
        let bytes = std::fs::read(&path).unwrap();
        if bytes.starts_with(b"PK\x03\x04") {
            let Ok(archive) = PhysPkgReader::new(&bytes) else {
                continue;
            };
            let Ok(names) = archive.member_names() else {
                continue;
            };
            let Some(manifest) = names
                .iter()
                .find(|name| name.eq_ignore_ascii_case("[Content_Types].xml"))
            else {
                continue;
            };
            packages += 1;
            let (defaults, overrides) = content_types(&archive.read_member(manifest).unwrap());
            for name in &names {
                if name.ends_with('/') {
                    continue;
                }
                let part = format!("/{}", name.trim_start_matches('/')).to_lowercase();
                let extension = name.rsplit_once('.').map(|(_, extension)| extension);
                let media_type = overrides
                    .get(&part)
                    .or_else(|| {
                        extension.and_then(|extension| defaults.get(&extension.to_lowercase()))
                    })
                    .map_or("", String::as_str);
                if !package::is_xml_part(name, media_type) {
                    continue;
                }
                members += 1;
                let payload = archive.read_member(name).unwrap();
                if let Err(error) = audit::verify_source(&payload, Limits::default()) {
                    refused.insert((relative.clone(), name.clone()), error.to_string());
                }
            }
        } else if relative.starts_with("ooxml/")
            && (relative.ends_with(".xml") || relative.ends_with(".rels"))
        {
            loose += 1;
            if let Err(error) = audit::verify_source(&bytes, Limits::default()) {
                refused.insert((relative.clone(), String::new()), error.to_string());
            }
        }
    }
    eprintln!(
        "source audit census: {packages} packages, {members} XML members, {loose} loose parts, \
         {} refused",
        refused.len()
    );
    let expected: BTreeMap<_, _> = REFUSED
        .iter()
        .map(|(file, member, error)| {
            (
                ((*file).to_owned(), (*member).to_owned()),
                (*error).to_owned(),
            )
        })
        .collect();
    assert_eq!(refused, expected);
    assert!(packages >= 330, "packages {packages}");
    assert!(members >= 7_000, "members {members}");
    assert!(loose >= 80, "loose parts {loose}");
}
