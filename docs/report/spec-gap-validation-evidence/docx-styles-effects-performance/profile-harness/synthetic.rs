//! Deterministic external synthetic effects resources for the initial profile.

use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;

use serde_json::{Value, json};
use sha2::{Digest as _, Sha256};
use zip::CompressionMethod;
use zip::ZipArchive;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

use super::smoke_adapter::{BoxError, Fixture};

type Result<T> = std::result::Result<T, BoxError>;

const TRANSITIONAL_W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn target_bytes(scale: super::Scale) -> usize {
    match scale {
        super::Scale::KiB64 => 64 * 1024,
        super::Scale::MiB1 => 1024 * 1024,
        super::Scale::Native => unreachable!("native resources do not use the generator"),
    }
}

fn synthetic_xml(scale: super::Scale) -> Vec<u8> {
    let target = target_bytes(scale);
    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:styles xmlns:w=\"{TRANSITIONAL_W}\" xmlns:x=\"urn:litchi:styles-effects-profile\">"
    )
    .into_bytes();
    let mut index = 0_u64;
    while xml.len() + 64 < target {
        let style = format!(
            "<w:style w:type=\"paragraph\" w:styleId=\"p{index:08}\"><w:name w:val=\"Profile {index:08}\"/><x:opaque data=\"profile-marker-{index:08}\"/></w:style>"
        );
        xml.extend_from_slice(style.as_bytes());
        index = index.saturating_add(1);
    }
    xml.extend_from_slice(b"<x:opaque data=\"styles-effects-profile-final-marker\"/></w:styles>");
    xml
}

fn sha256(bytes: &[u8]) -> String {
    let mut value = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        use std::fmt::Write as _;
        let _ = write!(value, "{byte:02x}");
    }
    value
}

fn replace_member(package: &[u8], member: &str, replacement: &[u8]) -> Result<Vec<u8>> {
    let mut archive = ZipArchive::new(Cursor::new(package))?;
    let mut output = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(&mut output);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index)?;
        let name = entry.name().to_owned();
        if entry.is_dir() {
            writer.add_directory(name, options)?;
            continue;
        }
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes)?;
        writer.start_file(name.as_str(), options)?;
        if name == member {
            writer.write_all(replacement)?;
        } else {
            writer.write_all(&bytes)?;
        }
    }
    writer.finish()?;
    Ok(output.into_inner())
}

fn effects_member(glossary: bool) -> &'static str {
    if glossary {
        "word/glossary/stylesWithEffects.xml"
    } else {
        "word/stylesWithEffects.xml"
    }
}

fn xml_observations(xml: &[u8]) -> Result<(u64, u64, u64)> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut events = 0_u64;
    let mut depth = 0_u64;
    let mut maximum_depth = 0_u64;
    let mut style_count = 0_u64;
    loop {
        match reader.read_event_into(&mut buffer)? {
            quick_xml::events::Event::Start(event) => {
                events = events.saturating_add(1);
                depth = depth.saturating_add(1);
                maximum_depth = maximum_depth.max(depth);
                if event.name().as_ref() == b"w:style" {
                    style_count = style_count.saturating_add(1);
                }
            },
            quick_xml::events::Event::Empty(event) => {
                events = events.saturating_add(1);
                if event.name().as_ref() == b"w:style" {
                    style_count = style_count.saturating_add(1);
                }
            },
            quick_xml::events::Event::End(_) => {
                events = events.saturating_add(1);
                depth = depth
                    .checked_sub(1)
                    .ok_or("synthetic XML depth underflow")?;
            },
            quick_xml::events::Event::Eof => break,
            _ => events = events.saturating_add(1),
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err("synthetic XML depth did not close".into());
    }
    Ok((events, maximum_depth, style_count))
}

pub fn manifest_for_lane(
    smoke_lane: &str,
    glossary: bool,
    scale: super::Scale,
    output_dir: Option<&Path>,
) -> Result<Value> {
    let base = super::smoke_adapter::fixture_for_lane(smoke_lane)?;
    let xml = synthetic_xml(scale);
    let member = effects_member(glossary);
    let package = replace_member(base.package.as_ref(), member, &xml)?;
    let (xml_events, xml_depth, style_count) = xml_observations(&xml)?;
    let opaque_marker_bytes = xml
        .windows(b"profile-marker-".len())
        .filter(|window| *window == b"profile-marker-")
        .count()
        .saturating_mul(b"profile-marker-".len())
        .saturating_add(b"styles-effects-profile-final-marker".len());
    let (resource_path, package_path) = if let Some(output_dir) = output_dir {
        fs::create_dir_all(output_dir)?;
        let owner = if glossary { "glossary" } else { "main" };
        let scale_name = scale.label();
        let resource_path = output_dir.join(format!("{owner}-{scale_name}.stylesWithEffects.xml"));
        let package_path = output_dir.join(format!("{owner}-{scale_name}.docx"));
        fs::write(&resource_path, &xml)?;
        fs::write(&package_path, &package)?;
        (Some(resource_path), Some(package_path))
    } else {
        (None, None)
    };
    Ok(json!({
        "schema": "docx-styles-effects-generated-fixture-v1",
        "source_fixture": base.name,
        "owner": if glossary { "glossary" } else { "main" },
        "scale": scale.label(),
        "member": member,
        "resource_bytes": xml.len(),
        "resource_sha256": sha256(&xml),
        "package_bytes": package.len(),
        "package_sha256": sha256(&package),
        "xml_events": xml_events,
        "xml_depth": xml_depth,
        "style_count": style_count,
        "opaque_marker_bytes": opaque_marker_bytes,
        "resource_path": resource_path.map(|path| path.to_string_lossy().into_owned()),
        "package_path": package_path.map(|path| path.to_string_lossy().into_owned()),
    }))
}

pub fn fixture_from_manifest(
    manifest_path: &Path,
    smoke_lane: &str,
    glossary: bool,
    scale: super::Scale,
) -> Result<Fixture> {
    let value: Value = serde_json::from_slice(&fs::read(manifest_path)?)?;
    if value.get("schema").and_then(Value::as_str)
        != Some("docx-styles-effects-generated-fixtures-v1")
        || value.get("source_commit").and_then(Value::as_str) != Some(super::SOURCE_COMMIT)
    {
        return Err("generated fixture manifest schema/source changed".into());
    }
    let owner = if glossary { "glossary" } else { "main" };
    let row = value
        .get("fixtures")
        .and_then(Value::as_array)
        .and_then(|rows| {
            rows.iter().find(|row| {
                row.get("owner").and_then(Value::as_str) == Some(owner)
                    && row.get("scale").and_then(Value::as_str) == Some(scale.label())
            })
        })
        .ok_or_else(|| format!("generated fixture is missing for {owner}/{}", scale.label()))?;
    if row.get("schema").and_then(Value::as_str) != Some("docx-styles-effects-generated-fixture-v1")
    {
        return Err("generated fixture row schema changed".into());
    }
    let package_path = PathBuf::from(
        row.get("package_path")
            .and_then(Value::as_str)
            .ok_or("generated package_path is missing")?,
    );
    let resource_path = PathBuf::from(
        row.get("resource_path")
            .and_then(Value::as_str)
            .ok_or("generated resource_path is missing")?,
    );
    if !package_path.is_absolute() || !resource_path.is_absolute() {
        return Err("generated fixture paths must be absolute".into());
    }
    let package = fs::read(&package_path)?;
    let resource = fs::read(&resource_path)?;
    let expected_package = row
        .get("package_sha256")
        .and_then(Value::as_str)
        .ok_or("generated package hash is missing")?;
    let expected_resource = row
        .get("resource_sha256")
        .and_then(Value::as_str)
        .ok_or("generated resource hash is missing")?;
    if sha256(&package) != expected_package || sha256(&resource) != expected_resource {
        return Err(format!(
            "generated fixture hash changed for {owner}/{}",
            scale.label()
        )
        .into());
    }
    let expected_package_bytes = row
        .get("package_bytes")
        .and_then(Value::as_u64)
        .ok_or("generated package size is missing")?;
    let expected_resource_bytes = row
        .get("resource_bytes")
        .and_then(Value::as_u64)
        .ok_or("generated resource size is missing")?;
    if package.len() as u64 != expected_package_bytes
        || resource.len() as u64 != expected_resource_bytes
    {
        return Err(format!(
            "generated fixture size changed for {owner}/{}",
            scale.label()
        )
        .into());
    }
    let member = effects_member(glossary);
    let members = {
        let mut archive = ZipArchive::new(Cursor::new(package.as_slice()))?;
        let mut selected = Vec::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            if entry.name() == member {
                entry.read_to_end(&mut selected)?;
                break;
            }
        }
        selected
    };
    if members != resource {
        return Err(format!(
            "generated resource does not match package for {owner}/{}",
            scale.label()
        )
        .into());
    }
    let base = super::smoke_adapter::fixture_for_lane(smoke_lane)?;
    Ok(Fixture {
        name: "synthetic-styles-effects",
        package: Arc::from(package),
        native: false,
        signed: false,
        main_present: base.main_present,
        glossary_present: base.glossary_present,
        expected_package_sha256: None,
    })
}
