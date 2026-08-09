//! Parity and regeneration tool for checked-in OOXML producer templates.
//!
//! Run the ignored test after changing a readable source template:
//!
//! `cargo test -p xml-minifier --test ooxml_assets update -- --ignored`

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "asset declaration macros intentionally precede the registry function"
)]
#![allow(
    clippy::unwrap_used,
    reason = "development parity tests fail immediately on invalid static paths or file IO"
)]

use std::collections::BTreeSet;
use std::fs;
use std::path::{Path, PathBuf};

use xml_minifier::audit::{self, package};
use xml_minifier::minified_xml;

#[derive(Clone, Copy)]
struct Asset {
    path: &'static str,
    generated: Option<&'static str>,
    source: &'static str,
    xml: &'static str,
}

macro_rules! asset {
    ($source:tt, $path:literal) => {
        Asset {
            path: $path,
            generated: None,
            source: $source,
            xml: minified_xml!($source),
        }
    };
}

macro_rules! asset_at {
    ($source:tt, $path:literal, $generated:literal) => {
        Asset {
            path: $path,
            generated: Some($generated),
            source: $source,
            xml: minified_xml!($source),
        }
    };
}

fn assets() -> Vec<Asset> {
    vec![
        asset!(
            "../../litchi-docx/src/resources/docProps/app.xml",
            "docx/resources/docProps/app.xml"
        ),
        asset!(
            "../../litchi-docx/src/resources/docProps/core.xml",
            "docx/resources/docProps/core.xml"
        ),
        asset!(
            "../../litchi-docx/src/resources/document.xml",
            "docx/resources/document.xml"
        ),
        asset!(
            "../../litchi-docx/src/resources/fontTable.xml",
            "docx/resources/fontTable.xml"
        ),
        asset!(
            "../../litchi-docx/src/resources/numbering.xml",
            "docx/resources/numbering.xml"
        ),
        asset!(
            "../../litchi-docx/src/resources/settings.xml",
            "docx/resources/settings.xml"
        ),
        asset!(
            "../../litchi-docx/src/resources/styles.xml",
            "docx/resources/styles.xml"
        ),
        asset!(
            "../../litchi-docx/src/resources/theme/theme1.xml",
            "docx/resources/theme/theme1.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/docProps/app.xml",
            "pptx/resources/docProps/app.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/docProps/core.xml",
            "pptx/resources/docProps/core.xml"
        ),
        asset_at!(
            "../../litchi-pptx/src/notes/resources/notesMaster.xml",
            "pptx/resources/notesMaster.xml",
            "../litchi-pptx/src/notes/resources/generated/notesMaster.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/presProps.xml",
            "pptx/resources/presProps.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/presentation.xml",
            "pptx/resources/presentation.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout1.xml",
            "pptx/resources/slideLayouts/slideLayout1.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout10.xml",
            "pptx/resources/slideLayouts/slideLayout10.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout11.xml",
            "pptx/resources/slideLayouts/slideLayout11.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout2.xml",
            "pptx/resources/slideLayouts/slideLayout2.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout3.xml",
            "pptx/resources/slideLayouts/slideLayout3.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout4.xml",
            "pptx/resources/slideLayouts/slideLayout4.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout5.xml",
            "pptx/resources/slideLayouts/slideLayout5.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout6.xml",
            "pptx/resources/slideLayouts/slideLayout6.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout7.xml",
            "pptx/resources/slideLayouts/slideLayout7.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout8.xml",
            "pptx/resources/slideLayouts/slideLayout8.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideLayouts/slideLayout9.xml",
            "pptx/resources/slideLayouts/slideLayout9.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/slideMasters/slideMaster1.xml",
            "pptx/resources/slideMasters/slideMaster1.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/theme/theme1.xml",
            "pptx/resources/theme/theme1.xml"
        ),
        asset!(
            "../../litchi-pptx/src/resources/viewProps.xml",
            "pptx/resources/viewProps.xml"
        ),
        asset!(
            "../../litchi-xlsb/src/host/resources/docProps/core.xml",
            "xlsb/resources/docProps/core.xml"
        ),
        asset!(
            "../../litchi-xlsb/src/host/resources/theme/theme1.xml",
            "xlsb/resources/theme/theme1.xml"
        ),
        asset!(
            "../../litchi-xlsx/src/package/resources/docProps/app.xml",
            "xlsx/resources/docProps/app.xml"
        ),
        asset!(
            "../../litchi-xlsx/src/package/resources/docProps/core.xml",
            "xlsx/resources/docProps/core.xml"
        ),
        asset!(
            "../../litchi-xlsx/src/package/resources/metadata.xml",
            "xlsx/resources/metadata.xml"
        ),
        asset!(
            "../../litchi-xlsx/src/package/resources/sharedStrings.xml",
            "xlsx/resources/sharedStrings.xml"
        ),
        asset!(
            "../../litchi-xlsx/src/package/resources/styles.xml",
            "xlsx/resources/styles.xml"
        ),
        asset!(
            "../../litchi-xlsx/src/package/resources/theme/theme1.xml",
            "xlsx/resources/theme/theme1.xml"
        ),
        asset!(
            "../../litchi-xlsx/src/package/resources/workbook.xml",
            "xlsx/resources/workbook.xml"
        ),
        asset!(
            "../../litchi-xlsx/src/package/resources/worksheets/sheet1.xml",
            "xlsx/resources/worksheets/sheet1.xml"
        ),
    ]
}

fn generated_path(asset: Asset) -> PathBuf {
    if let Some(path) = asset.generated {
        return PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(path);
    }
    let (owner, relative) = asset.path.split_once("/resources/").unwrap();
    let (crate_path, resource_root) = match owner {
        "docx" => ("../litchi-docx/src", "resources"),
        "pptx" => ("../litchi-pptx/src", "resources"),
        "xlsb" => ("../litchi-xlsb/src/host", "resources"),
        "xlsx" => ("../litchi-xlsx/src/package", "resources"),
        _ => panic!("unknown OOXML asset owner: {owner}"),
    };
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join(crate_path)
        .join(resource_root)
        .join("generated")
        .join(relative)
}

fn source_path(asset: Asset) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join(asset.source)
}

fn workspace_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .unwrap()
        .to_path_buf()
}

fn generated_asset_paths() -> BTreeSet<PathBuf> {
    let mut paths = Vec::new();
    collect_generated_xml(&workspace_root().join("crates"), &mut paths);
    paths.sort();
    paths.into_iter().collect()
}

fn registered_compact_crate_assets() -> BTreeSet<PathBuf> {
    // Add non-generated, crate-owned producer assets here only when they are
    // intentionally embedded by production code. Corpus and provenance XML
    // never belongs in this set.
    [
        "crates/litchi-docx/src/resources/customXml/item1.xml",
        "crates/litchi-docx/src/resources/customXml/itemProps1.xml",
        "crates/litchi-docx/src/resources/stylesWithEffects.xml",
    ]
    .into_iter()
    .map(|relative| workspace_root().join(relative).canonicalize().unwrap())
    .collect()
}

fn collect_generated_xml(directory: &Path, paths: &mut Vec<PathBuf>) {
    for entry_result in fs::read_dir(directory).unwrap() {
        let entry = entry_result.unwrap();
        let path = entry.path();
        let kind = entry.file_type().unwrap();
        if kind.is_dir() {
            collect_generated_xml(&path, paths);
        } else if kind.is_file() && is_generated_xml(&path) {
            paths.push(path);
        }
    }
}

fn is_generated_xml(path: &Path) -> bool {
    path.extension().is_some_and(|extension| extension == "xml")
        && path
            .components()
            .any(|component| component.as_os_str() == "generated")
        && path
            .components()
            .any(|component| component.as_os_str() == "resources")
}

fn production_rust_sources() -> Vec<PathBuf> {
    let mut sources = Vec::new();
    collect_production_rust_sources(&workspace_root().join("crates"), &mut sources);
    sources.sort();
    sources
}

fn collect_production_rust_sources(directory: &Path, sources: &mut Vec<PathBuf>) {
    for entry_result in fs::read_dir(directory).unwrap() {
        let entry = entry_result.unwrap();
        let path = entry.path();
        let kind = entry.file_type().unwrap();
        if kind.is_dir() {
            if !is_iwork_owner(&path) {
                collect_production_rust_sources(&path, sources);
            }
        } else if kind.is_file()
            && path.extension().is_some_and(|extension| extension == "rs")
            && path
                .components()
                .any(|component| component.as_os_str() == "src")
        {
            sources.push(path);
        }
    }
}

fn is_iwork_owner(path: &Path) -> bool {
    path.components().any(|component| {
        component.as_os_str().to_str().is_some_and(|name| {
            name.starts_with("litchi-iwa")
                || matches!(name, "litchi-keynote" | "litchi-numbers" | "litchi-pages")
        })
    })
}

fn is_provenance_fixture(path: &Path) -> bool {
    let root = workspace_root();
    path.starts_with(root.join("test-data")) || path.starts_with(root.join("3rdparty"))
}

fn xml_include_paths(source: &str) -> Vec<&str> {
    const PREFIX: &str = "include_str!(";

    let mut paths = Vec::new();
    let mut remaining = source;
    while let Some(index) = remaining.find(PREFIX) {
        let after_prefix = &remaining[index + PREFIX.len()..];
        let literal = after_prefix.trim_start();
        let Some(literal_body) = literal.strip_prefix('"') else {
            remaining = after_prefix;
            continue;
        };
        let Some(end) = literal_body.find('"') else {
            break;
        };
        let path = &literal_body[..end];
        if Path::new(path)
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        {
            paths.push(path);
        }
        remaining = &literal_body[end + 1..];
    }
    paths
}

#[test]
fn checked_in_assets_match_minifier_output() {
    let limits = audit::Limits::default();
    for asset in assets() {
        let source_path = source_path(asset);
        let source = fs::read_to_string(&source_path).unwrap();
        let path = generated_path(asset);
        let actual = fs::read_to_string(&path).unwrap();
        assert_eq!(
            source,
            asset.xml,
            "authoring source is not already compact: {}",
            source_path.display()
        );
        assert_eq!(
            actual,
            asset.xml,
            "generated asset is stale: {}",
            path.display()
        );
        let _source_report = audit::verify(source.as_bytes(), limits).unwrap_or_else(|error| {
            panic!("authoring source is not compact ({}): {error}", asset.path)
        });
        let _stored_report = package::verify(
            [package::Part::new(asset.path, actual.as_bytes())],
            package::Limits::default(),
        )
        .unwrap_or_else(|error| panic!("checked-in asset is not compact: {error}"));
    }
}

#[test]
fn registry_covers_every_checked_in_generated_xml_asset() {
    let registered = assets()
        .into_iter()
        .filter_map(|asset| {
            let path = generated_path(asset).canonicalize().unwrap();
            is_generated_xml(&path).then_some(path)
        })
        .collect::<BTreeSet<_>>();
    let discovered = generated_asset_paths();
    assert_eq!(
        discovered, registered,
        "every checked-in resources/generated XML output must have an ADR-0017 parity entry"
    );
}

#[test]
fn explicitly_registered_compact_assets_are_compact() {
    let limits = audit::Limits::default();
    for asset_path in registered_compact_crate_assets() {
        let xml = fs::read(&asset_path).unwrap();
        let _report = audit::verify(&xml, limits).unwrap_or_else(|error| {
            panic!(
                "registered crate XML asset is not compact: {}: {error}",
                asset_path.display()
            )
        });
    }
}

#[test]
fn production_xml_includes_are_registered_compact_assets() {
    let mut registered = generated_asset_paths();
    registered.extend(registered_compact_crate_assets());
    let limits = audit::Limits::default();
    let mut includes = BTreeSet::new();

    for source_path in production_rust_sources() {
        let source = fs::read_to_string(&source_path).unwrap();
        for included in xml_include_paths(&source) {
            let asset_path = source_path
                .parent()
                .unwrap()
                .join(included)
                .canonicalize()
                .unwrap();
            if is_provenance_fixture(&asset_path) {
                continue;
            }
            assert!(
                asset_path.starts_with(workspace_root().join("crates")),
                "production XML include must be crate-owned or a provenance fixture: {} (from {})",
                asset_path.display(),
                source_path.display(),
            );
            assert!(
                registered.contains(&asset_path),
                "production XML include is not a registered compact producer asset: {} (from {})",
                asset_path.display(),
                source_path.display(),
            );
            includes.insert(asset_path);
        }
    }

    for asset_path in includes {
        let xml = fs::read(&asset_path).unwrap();
        let _report = audit::verify(&xml, limits).unwrap_or_else(|error| {
            panic!(
                "production XML include is not compact: {}: {error}",
                asset_path.display()
            )
        });
    }
}

#[test]
#[ignore = "explicit checked-in asset regeneration"]
fn update_checked_in_assets() {
    for asset in assets() {
        fs::write(source_path(asset), asset.xml).unwrap();
        let path = generated_path(asset);
        fs::create_dir_all(path.parent().unwrap()).unwrap();
        fs::write(path, asset.xml).unwrap();
    }
}
