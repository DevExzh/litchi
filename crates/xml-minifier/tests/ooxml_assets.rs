//! Parity and regeneration tool for checked-in OOXML producer templates.
//!
//! Run the ignored test after changing a readable source template:
//!
//! `cargo test -p xml-minifier --test ooxml_assets update -- --ignored`

use std::fs;
use std::path::PathBuf;

use xml_minifier::minified_xml;

#[derive(Clone, Copy)]
struct Asset {
    path: &'static str,
    xml: &'static str,
}

macro_rules! asset {
    ($source:tt, $path:literal) => {
        Asset {
            path: $path,
            xml: minified_xml!($source),
        }
    };
}

fn assets() -> Vec<Asset> {
    vec![
        asset!(
            "../../litchi-ooxml/src/docx/resources/docProps/app.xml",
            "docx/resources/docProps/app.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/docx/resources/docProps/core.xml",
            "docx/resources/docProps/core.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/docx/resources/document.xml",
            "docx/resources/document.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/docx/resources/fontTable.xml",
            "docx/resources/fontTable.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/docx/resources/numbering.xml",
            "docx/resources/numbering.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/docx/resources/settings.xml",
            "docx/resources/settings.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/docx/resources/styles.xml",
            "docx/resources/styles.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/docx/resources/theme/theme1.xml",
            "docx/resources/theme/theme1.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/docx/resources/webSettings.xml",
            "docx/resources/webSettings.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/docProps/app.xml",
            "pptx/resources/docProps/app.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/docProps/core.xml",
            "pptx/resources/docProps/core.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/notesMaster.xml",
            "pptx/resources/notesMaster.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/presProps.xml",
            "pptx/resources/presProps.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/presentation.xml",
            "pptx/resources/presentation.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout1.xml",
            "pptx/resources/slideLayouts/slideLayout1.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout10.xml",
            "pptx/resources/slideLayouts/slideLayout10.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout11.xml",
            "pptx/resources/slideLayouts/slideLayout11.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout2.xml",
            "pptx/resources/slideLayouts/slideLayout2.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout3.xml",
            "pptx/resources/slideLayouts/slideLayout3.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout4.xml",
            "pptx/resources/slideLayouts/slideLayout4.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout5.xml",
            "pptx/resources/slideLayouts/slideLayout5.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout6.xml",
            "pptx/resources/slideLayouts/slideLayout6.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout7.xml",
            "pptx/resources/slideLayouts/slideLayout7.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout8.xml",
            "pptx/resources/slideLayouts/slideLayout8.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideLayouts/slideLayout9.xml",
            "pptx/resources/slideLayouts/slideLayout9.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/slideMasters/slideMaster1.xml",
            "pptx/resources/slideMasters/slideMaster1.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/tableStyles.xml",
            "pptx/resources/tableStyles.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/theme/theme1.xml",
            "pptx/resources/theme/theme1.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/pptx/resources/viewProps.xml",
            "pptx/resources/viewProps.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsb/resources/docProps/core.xml",
            "xlsb/resources/docProps/core.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsb/resources/theme/theme1.xml",
            "xlsb/resources/theme/theme1.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsx/resources/docProps/app.xml",
            "xlsx/resources/docProps/app.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsx/resources/docProps/core.xml",
            "xlsx/resources/docProps/core.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsx/resources/metadata.xml",
            "xlsx/resources/metadata.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsx/resources/sharedStrings.xml",
            "xlsx/resources/sharedStrings.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsx/resources/styles.xml",
            "xlsx/resources/styles.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsx/resources/theme/theme1.xml",
            "xlsx/resources/theme/theme1.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsx/resources/workbook.xml",
            "xlsx/resources/workbook.xml"
        ),
        asset!(
            "../../litchi-ooxml/src/xlsx/resources/worksheets/sheet1.xml",
            "xlsx/resources/worksheets/sheet1.xml"
        ),
    ]
}

fn generated_path(asset: Asset) -> PathBuf {
    let (owner, relative) = asset.path.split_once("/resources/").unwrap();
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../litchi-ooxml/src")
        .join(owner)
        .join("resources/generated")
        .join(relative)
}

#[test]
fn checked_in_assets_match_minifier_output() {
    for asset in assets() {
        let path = generated_path(asset);
        let actual = fs::read_to_string(&path).unwrap();
        assert_eq!(
            actual,
            asset.xml,
            "generated asset is stale: {}",
            path.display()
        );
    }
}

#[test]
#[ignore = "explicit checked-in asset regeneration"]
fn update_checked_in_assets() {
    for asset in assets() {
        let path = generated_path(asset);
        fs::create_dir_all(path.parent().unwrap()).unwrap();
        fs::write(path, asset.xml).unwrap();
    }
}
