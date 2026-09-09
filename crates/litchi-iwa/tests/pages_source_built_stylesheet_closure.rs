//! Regression coverage for the stylesheet imported by source-built and native
//! Pages tables.
//!
//! Pages reuses the Numbers table graph when it creates an initial body table.
//! The imported stylesheet archive must therefore retain the complete typed
//! stylesheet closure, including the Numbers paragraph and list presets.  A
//! physical object identifier alone is not sufficient here: Pages already
//! uses some of those identifiers for unrelated objects.

use std::collections::{BTreeMap, BTreeSet};
use std::io;

use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa_archive::iwa::{Archive, ArchiveObject, SnappyStream};
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_protos::{tss, tst};
use litchi_pages::Package as FocusedPagesPackage;
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
const STYLESHEET_COMPONENT: &str = "Index/DocumentStylesheet.iwa";
const STYLESHEET_OBJECT_ID: u64 = 40;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const NATIVE_SOURCE_BUILT_TABLE_STYLESHEET_BEFORE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/source-built-table-stylesheet-before.pages"
));
const NATIVE_SOURCE_BUILT_TABLE_STYLESHEET_SAVED: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/source-built-table-stylesheet-native-saved.pages"
));

fn component_archives(source: &[u8]) -> TestResult<BTreeMap<String, Archive>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut archives = BTreeMap::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let decompressed = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&decompressed)?;
        if archives.insert(entry.name().to_owned(), archive).is_some() {
            return Err(io::Error::other(format!("duplicate IWA member {}", entry.name())).into());
        }
    }
    Ok(archives)
}

fn object_references(object: &ArchiveObject) -> BTreeSet<u64> {
    object
        .archive_info
        .message_infos
        .iter()
        .flat_map(|message| {
            message.object_references.iter().copied().chain(
                message
                    .field_infos
                    .iter()
                    .flat_map(|field| field.object_references.iter().copied()),
            )
        })
        .collect()
}

fn package_object_ids(archives: &BTreeMap<String, Archive>) -> BTreeSet<u64> {
    archives
        .values()
        .flat_map(|archive| {
            archive
                .objects
                .iter()
                .filter_map(|object| object.archive_info.identifier)
        })
        .collect()
}

fn stylesheet_root(
    archives: &BTreeMap<String, Archive>,
) -> TestResult<(&Archive, tss::StylesheetArchive)> {
    let archive = archives
        .get(STYLESHEET_COMPONENT)
        .ok_or_else(|| io::Error::other("missing DocumentStylesheet component"))?;
    let object = archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(STYLESHEET_OBJECT_ID))
        .ok_or_else(|| io::Error::other("missing imported stylesheet root object"))?;
    let payload = object
        .messages
        .iter()
        .find(|message| message.type_ == STYLESHEET_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing imported stylesheet root payload"))?;
    Ok((
        archive,
        tss::StylesheetArchive::decode(payload.data.as_slice())?,
    ))
}

fn table_model(
    archives: &BTreeMap<String, Archive>,
    model_object_id: u64,
) -> TestResult<tst::TableModelArchive> {
    let archive = archives
        .get(DOCUMENT_COMPONENT)
        .ok_or_else(|| io::Error::other("missing Document component"))?;
    let object = archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(model_object_id))
        .ok_or_else(|| io::Error::other("missing source-built table model object"))?;
    let payload = object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing source-built table model payload"))?;
    Ok(tst::TableModelArchive::decode(payload.data.as_slice())?)
}

fn assert_stylesheet_closure(source: &[u8], model_object_id: u64) -> TestResult<()> {
    let archives = component_archives(source)?;
    let package_ids = package_object_ids(&archives);
    let (stylesheet_archive, stylesheet) = stylesheet_root(&archives)?;
    let retained_ids = stylesheet_archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .collect::<BTreeSet<_>>();
    let stylesheet_ids = stylesheet
        .styles
        .iter()
        .map(|reference| reference.identifier)
        .collect::<BTreeSet<_>>();

    // The source Numbers package contains additional paragraph and list
    // presets, but Pages does not copy those objects into its table graph.
    // Keep their references out of the imported root instead of allowing an
    // unrelated Pages object with the same identifier to satisfy the check.
    for identifier in 41..=47 {
        assert!(
            !stylesheet_ids.contains(&identifier),
            "unused Numbers stylesheet preset {identifier} remains in the imported Pages root"
        );
    }

    // Verify every typed stylesheet edge resolves to a physical object and
    // that the target still has the style message type expected by Numbers.
    let expected_style_types = [
        (11, 2_023),  // list style
        (12, 2_022),  // body paragraph style
        (13, 2_021),  // character style
        (14, 2_025),  // shape style
        (15, 3_016),  // media style
        (16, 10_024), // drop-cap style
        (17, 12_050), // sheet style
        (18, 6_003),  // table style
        (19, 6_004),  // cell style
    ];
    for &(identifier, _) in &expected_style_types {
        assert!(
            stylesheet_ids.contains(&identifier),
            "retained stylesheet object {identifier} is absent from the imported root"
        );
        assert!(
            retained_ids.contains(&identifier),
            "retained stylesheet object {identifier} is absent from the physical archive"
        );
    }
    let expected_style_ids = expected_style_types
        .iter()
        .map(|(identifier, _)| *identifier)
        .collect::<BTreeSet<_>>();
    assert_eq!(
        stylesheet_ids, expected_style_ids,
        "imported stylesheet root retained an unexpected typed style edge"
    );
    for reference in &stylesheet.styles {
        let identifier = reference.identifier;
        assert!(
            retained_ids.contains(&identifier),
            "stylesheet style reference {identifier} has no retained object"
        );
        let expected_type =
            expected_style_types
                .iter()
                .find_map(|(known_identifier, message_type)| {
                    (*known_identifier == identifier).then_some(*message_type)
                });
        let object = stylesheet_archive
            .objects
            .iter()
            .find(|object| object.archive_info.identifier == Some(identifier))
            .ok_or_else(|| io::Error::other(format!("missing stylesheet object {identifier}")))?;
        if let Some(expected_type) = expected_type {
            assert!(
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == expected_type),
                "stylesheet object {identifier} has no expected message type {expected_type}"
            );
        }
    }
    for entry in &stylesheet.identifier_to_style_map {
        let identifier = entry.style.identifier;
        assert!(
            retained_ids.contains(&identifier),
            "stylesheet identifier map target {identifier} has no retained object"
        );
        assert!(
            stylesheet_ids.contains(&identifier),
            "stylesheet identifier map target {identifier} is absent from styles"
        );
    }
    let body_identifier = "litchi-paragraph-body";
    let body_entries = stylesheet
        .identifier_to_style_map
        .iter()
        .filter(|entry| entry.identifier == body_identifier)
        .collect::<Vec<_>>();
    assert_eq!(
        body_entries.len(),
        1,
        "the imported Numbers Body preset mapping must be retained exactly once"
    );
    assert_eq!(body_entries[0].style.identifier, 12);
    let map_targets = stylesheet
        .identifier_to_style_map
        .iter()
        .map(|entry| entry.style.identifier)
        .collect::<BTreeSet<_>>();
    assert_eq!(
        map_targets,
        BTreeSet::from([12]),
        "unused Numbers preset targets remain in the imported identifier map"
    );

    // Header style references are part of the same imported graph.  They must
    // point at the exact retained stylesheet entries and preserve the Numbers
    // builder's row/column pairing.
    let model = table_model(&archives, model_object_id)?;
    let header_refs = [
        ("header row text", &model.header_row_text_style),
        ("header column text", &model.header_column_text_style),
        ("header row", &model.header_row_style),
        ("header column", &model.header_column_style),
    ];
    let mut header_ids = Vec::with_capacity(header_refs.len());
    for (name, reference) in header_refs {
        let identifier = reference.identifier;
        assert!(
            retained_ids.contains(&identifier),
            "{name} style reference {identifier} is not physically retained"
        );
        assert!(
            stylesheet_ids.contains(&identifier),
            "{name} style reference {identifier} is absent from the stylesheet root"
        );
        header_ids.push(identifier);
    }
    assert_eq!(
        header_ids[0], header_ids[1],
        "table header row and column text styles diverged"
    );
    assert_eq!(
        header_ids[2], header_ids[3],
        "table header row and column cell styles diverged"
    );
    assert_eq!(
        header_ids,
        [12, 12, 19, 19],
        "source-built table header styles no longer use the retained Numbers defaults"
    );

    // ArchiveInfo carries additional strong edges that must remain materialized
    // after the imported graph is copied.  Check the complete generated
    // package, not only the stylesheet component, so a missing table child or
    // cross-component edge cannot be hidden by the typed stylesheet checks.
    for (component, archive) in &archives {
        for object in &archive.objects {
            for identifier in object_references(object) {
                assert!(
                    package_ids.contains(&identifier),
                    "{component} object {:?} has dangling strong reference {identifier}",
                    object.archive_info.identifier
                );
            }
        }
    }

    Ok(())
}

fn assert_native_stylesheet_fixture(source: &[u8], label: &str) -> TestResult<()> {
    let focused = FocusedPagesPackage::from_bytes(source).map_err(|error| {
        io::Error::other(format!("{label} focused package parse failed: {error:?}"))
    })?;
    let focused_text = focused
        .text()
        .map_err(|error| io::Error::other(format!("{label} focused text failed: {error:?}")))?;
    assert!(
        focused_text.contains("Cities"),
        "{label} focused body text lost the Cities marker: {focused_text:?}"
    );
    assert!(
        focused_text.contains('\u{fffc}'),
        "{label} focused body text lost the table object replacement character: {focused_text:?}"
    );
    let mut focused_reencoded = Vec::new();
    focused.write_to(&mut focused_reencoded)?;
    assert_eq!(
        focused_reencoded, source,
        "{label} focused read/reencode changed the native package"
    );

    let focused_tables = focused.body_tables()?;
    assert_eq!(
        focused_tables.len(),
        1,
        "{label} should expose one body table"
    );
    let focused_table = focused_tables
        .get(0)
        .ok_or_else(|| io::Error::other(format!("{label} body table is missing")))?;
    assert_eq!(focused_table.name(), "Cities", "{label} table name changed");
    assert_eq!(
        (focused_table.rows(), focused_table.columns()),
        (5, 4),
        "{label} dimensions changed"
    );

    let editor = PagesEditor::from_bytes(source).map_err(|error| {
        io::Error::other(format!("{label} legacy package parse failed: {error:?}"))
    })?;
    let body_text = editor
        .body_text()
        .map_err(|error| io::Error::other(format!("{label} legacy body text failed: {error:?}")))?;
    assert!(
        body_text.contains("Cities"),
        "{label} legacy body text lost the Cities marker: {body_text:?}"
    );
    assert!(
        body_text.contains('\u{fffc}'),
        "{label} legacy body text lost the table object replacement character: {body_text:?}"
    );
    assert_eq!(
        editor.to_bytes()?,
        source,
        "{label} legacy no-op changed bytes"
    );

    let reopened = PagesEditor::from_bytes(source)?;
    assert_eq!(
        reopened.body_text()?,
        body_text,
        "{label} body changed on reopen"
    );
    let focused_reopened = FocusedPagesPackage::from_bytes(source).map_err(|error| {
        io::Error::other(format!("{label} focused package reopen failed: {error:?}"))
    })?;
    assert_eq!(
        focused_reopened.body_tables()?,
        focused_tables,
        "{label} table changed on reopen"
    );
    assert_eq!(reopened.to_bytes()?, source, "{label} reopen changed bytes");
    Ok(())
}

#[test]
fn source_built_pages_table_retains_numbers_stylesheet_closure() -> TestResult {
    let pages = PagesDocumentBuilder::new()
        .body_text("Body text")
        .body_table("Imported Numbers table", 5, 4)
        .build()?;
    let source = pages.to_bytes()?;
    let model_object_id = PagesEditor::from_bytes(&source)?
        .tables()?
        .first()
        .ok_or_else(|| io::Error::other("source-built table is missing after initial build"))?
        .model_object_id;

    assert_stylesheet_closure(&source, model_object_id)?;

    let reopened = PagesEditor::from_bytes(&source)?;
    let reopened_model_object_id = reopened
        .tables()?
        .first()
        .ok_or_else(|| io::Error::other("source-built table is missing after reopen"))?
        .model_object_id;
    assert_eq!(model_object_id, reopened_model_object_id);
    assert_stylesheet_closure(&source, reopened_model_object_id)?;
    Ok(())
}

#[test]
fn native_source_built_pages_table_stylesheet_fixtures_roundtrip_semantically() -> TestResult {
    assert_native_stylesheet_fixture(
        NATIVE_SOURCE_BUILT_TABLE_STYLESHEET_BEFORE,
        "source-built-table-stylesheet-before",
    )?;
    assert_native_stylesheet_fixture(
        NATIVE_SOURCE_BUILT_TABLE_STYLESHEET_SAVED,
        "source-built-table-stylesheet-native-saved",
    )?;
    Ok(())
}
