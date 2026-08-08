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

use std::fs;
use std::path::PathBuf;

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
#[ignore = "explicit checked-in asset regeneration"]
fn update_checked_in_assets() {
    for asset in assets() {
        fs::write(source_path(asset), asset.xml).unwrap();
        let path = generated_path(asset);
        fs::create_dir_all(path.parent().unwrap()).unwrap();
        fs::write(path, asset.xml).unwrap();
    }
}
