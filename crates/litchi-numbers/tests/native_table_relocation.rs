//! Native Numbers coverage for moving a table between rooted sheets.
//!
//! This fixture was authored and reopened by Numbers 14.4. The source sheet
//! owns `Table 1` and an ordinary image; the destination sheet already owns a
//! separate table. Relocation must preserve both table projections and the
//! image graph while changing only the owning document/table components.

use std::io;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_core::{Archive, ArchiveInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tsd, tsp};
use litchi_numbers::cell::Value;
use litchi_numbers::{
    ImageSelector, Package, SheetSelector, Table, table::relocation::transaction::Path,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const SOURCE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/numbers/table-relocation-native.numbers"
);
const RESAVED_SOURCE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/numbers/table-relocation-native-resaved.numbers"
);
const CALCULATION_ENGINE_MEMBER: &str = "Index/CalculationEngine.iwa";
const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const SOURCE_SHEET: &str = "Sheet 1";
const SOURCE_TABLE: &str = "Table 1";
const DESTINATION_SHEET: &str = "Sheet 2";
const DESTINATION_TABLE: &str = "Destination table";

#[derive(Debug, Clone, PartialEq)]
struct ImageSemanticSnapshot {
    data_identifier: Option<u64>,
    original_size: Option<(f32, f32)>,
    natural_size: Option<(f32, f32)>,
    adjustments: Option<Vec<u8>>,
}

#[derive(Debug, Clone, PartialEq)]
struct ImageArchiveSnapshot {
    member: String,
    object_id: u64,
    archive_info: ArchiveInfo,
    messages: Vec<RawMessage>,
    semantic: ImageSemanticSnapshot,
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn table(package: &Package, sheet: &str, name: &str) -> TestResult<Table> {
    package
        .table(sheet, name)?
        .cloned()
        .ok_or_else(|| io::Error::other(format!("table {sheet}/{name} is missing")).into())
}

fn table_names(package: &Package, sheet: usize) -> Vec<String> {
    package.sheets()[sheet]
        .tables()
        .map(|table| table.name().to_owned())
        .collect()
}

fn text_at(table: &Table, address: &str) -> TestResult<String> {
    match table.get_a1(address)? {
        Some(Value::Text(value)) => Ok(value.clone()),
        Some(value) => {
            Err(io::Error::other(format!("{address} contains {value:?}, expected text")).into())
        },
        None => Err(io::Error::other(format!("{address} is missing")).into()),
    }
}

fn image_archive(source: &[u8]) -> TestResult<ImageArchiveSnapshot> {
    let catalog = Catalog::from_bytes(source)?;
    let mut found = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = match SnappyStream::decompress(entry.data()) {
            Ok(stream) => stream,
            Err(_) => continue,
        };
        let archive = match Archive::parse(stream.as_bytes()) {
            Ok(archive) => archive,
            Err(_) => continue,
        };
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ != IMAGE_MESSAGE_TYPE {
                    continue;
                }
                let image = match tsd::ImageArchive::decode(message.data.as_slice()) {
                    Ok(image) => image,
                    Err(_) => continue,
                };
                if image.data.is_none() {
                    continue;
                }
                if found.is_some() {
                    return Err(io::Error::other(
                        "native relocation fixture contains multiple ordinary images",
                    )
                    .into());
                }
                let size = |size: Option<&tsp::Size>| size.map(|size| (size.width, size.height));
                found = Some(ImageArchiveSnapshot {
                    member: entry.name().to_owned(),
                    object_id: object
                        .archive_info
                        .identifier
                        .ok_or_else(|| io::Error::other("native image has no object ID"))?,
                    archive_info: object.archive_info.clone(),
                    messages: object.messages.clone(),
                    semantic: ImageSemanticSnapshot {
                        data_identifier: image.data.as_ref().map(|data| data.identifier),
                        original_size: size(image.original_size.as_ref()),
                        natural_size: size(image.natural_size.as_ref()),
                        adjustments: image
                            .image_adjustments
                            .as_ref()
                            .map(|adjustments| adjustments.encode_to_vec()),
                    },
                });
            }
        }
    }
    found.ok_or_else(|| io::Error::other("native relocation image is missing").into())
}

fn assert_data_assets_unchanged(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut count = 0;
    for entry in before
        .iter()
        .filter(|entry| entry.name().starts_with("Data/"))
    {
        count += 1;
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or_else(|| io::Error::other("relocation removed an image asset"))?;
        assert_eq!(entry.data(), candidate.data(), "image asset bytes changed");
        assert_eq!(
            entry.raw_record().local_record(),
            candidate.raw_record().local_record(),
            "image asset ZIP record changed"
        );
    }
    assert!(count > 0, "native relocation fixture has no Data assets");
    Ok(())
}

fn is_preview(name: &str) -> bool {
    matches!(
        name,
        "preview.jpg" | "preview-micro.jpg" | "preview-web.jpg"
    )
}

fn assert_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for entry in before.iter() {
        let Some(candidate) = after.iter().find(|other| other.name() == entry.name()) else {
            assert!(
                is_preview(entry.name()),
                "unexpected member removed: {}",
                entry.name()
            );
            continue;
        };
        if is_preview(entry.name()) {
            continue;
        }
        if entry.data() != candidate.data() {
            changed.push(entry.name().to_owned());
        } else {
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unchanged member {} lost its exact local record",
                entry.name()
            );
        }
    }
    for entry in after.iter() {
        if before.iter().all(|other| other.name() != entry.name()) {
            assert!(
                is_preview(entry.name()),
                "unexpected member added: {}",
                entry.name()
            );
        }
    }
    changed.sort_unstable();
    assert_eq!(
        changed.len(),
        2,
        "relocation changed unexpected package members: {changed:?}"
    );
    assert_eq!(
        changed,
        vec![
            CALCULATION_ENGINE_MEMBER.to_owned(),
            DOCUMENT_MEMBER.to_owned()
        ],
        "relocation changed unexpected package members"
    );
    Ok(())
}

fn assert_relocated(
    package: &Package,
    source_table: &Table,
    destination: &Table,
    marker: &str,
) -> TestResult {
    assert_eq!(package.sheets().len(), 2);
    assert_eq!(table_names(package, 0), Vec::<String>::new());
    assert_eq!(
        table_names(package, 1),
        vec![DESTINATION_TABLE.to_owned(), SOURCE_TABLE.to_owned()]
    );
    assert_eq!(
        table(package, DESTINATION_SHEET, DESTINATION_TABLE)?,
        *destination
    );
    let moved = table(package, DESTINATION_SHEET, SOURCE_TABLE)?;
    assert_eq!(moved, *source_table);
    assert_eq!(text_at(&moved, "B2")?, marker);
    Ok(())
}

#[test]
fn native_table_relocation_preserves_tables_image_and_exact_inverse() -> TestResult {
    let source = std::fs::read(SOURCE)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package
            .sheets()
            .iter()
            .map(|sheet| sheet.name())
            .collect::<Vec<_>>(),
        [SOURCE_SHEET, DESTINATION_SHEET]
    );
    let source_table = table(&package, SOURCE_SHEET, SOURCE_TABLE)?;
    let destination = table(&package, DESTINATION_SHEET, DESTINATION_TABLE)?;
    assert_eq!(source_table.dimensions().rows(), 22);
    assert_eq!(source_table.dimensions().columns(), 7);
    assert_eq!(destination.dimensions().rows(), 10);
    assert_eq!(destination.dimensions().columns(), 5);
    assert_eq!(
        text_at(&source_table, "B2")?,
        "Native image adjustment marker"
    );
    assert_eq!(text_at(&destination, "B2")?, "Destination stays");
    let image_before = image_archive(&source)?;

    let noop = package.move_table(SOURCE_SHEET, SOURCE_TABLE, SOURCE_SHEET)?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, source);

    let edit = package.edit_table_relocation(SOURCE_SHEET, SOURCE_TABLE, DESTINATION_SHEET)?;
    assert_eq!(
        edit.path(),
        Path::Table {
            source_sheet: 0,
            table: 0,
            destination_sheet: 1,
        }
    );
    assert_eq!(edit.source_sheet_position(), 0);
    assert_eq!(edit.table_position(), 0);
    assert_eq!(edit.destination_sheet_position(), 1);
    assert_eq!(edit.destination_table_position(), 1);
    let commit = edit.commit()?;
    assert!(!commit.patch().is_noop());
    let patch = commit.patch().clone();
    let target = exact_bytes(commit.package())?;
    assert_relocated(
        commit.package(),
        &source_table,
        &destination,
        "Native image adjustment marker",
    )?;
    assert_eq!(image_archive(&target)?, image_before);
    assert_data_assets_unchanged(&source, &target)?;
    assert_locality(&source, &target)?;

    let reopened = Package::from_bytes(&target)?;
    assert_relocated(
        &reopened,
        &source_table,
        &destination,
        "Native image adjustment marker",
    )?;
    let reopened_bytes = exact_bytes(&reopened)?;
    assert_eq!(image_archive(&reopened_bytes)?, image_before);
    assert_data_assets_unchanged(&source, &reopened_bytes)?;

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    let restored = reopened.apply_table_relocation(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn native_resaved_relocation_fixture_preserves_moved_state_image_and_noop() -> TestResult {
    const MOVED_MARKER: &str = "Native relocation saved marker";

    let resaved = std::fs::read(RESAVED_SOURCE)?;
    let package = Package::from_bytes(&resaved)?;
    assert_eq!(
        package
            .sheets()
            .iter()
            .map(|sheet| sheet.name())
            .collect::<Vec<_>>(),
        [SOURCE_SHEET, DESTINATION_SHEET]
    );
    let moved = table(&package, DESTINATION_SHEET, SOURCE_TABLE)?;
    let destination = table(&package, DESTINATION_SHEET, DESTINATION_TABLE)?;
    assert_eq!(moved.dimensions().rows(), 22);
    assert_eq!(moved.dimensions().columns(), 7);
    assert_eq!(destination.dimensions().rows(), 10);
    assert_eq!(destination.dimensions().columns(), 5);
    assert_relocated(&package, &moved, &destination, MOVED_MARKER)?;

    let original = std::fs::read(SOURCE)?;
    let resaved_image = image_archive(&resaved)?;
    let original_image = image_archive(&original)?;
    assert_eq!(resaved_image.semantic, original_image.semantic);
    let original_package = Package::from_bytes(&original)?;
    assert_eq!(
        package
            .sheet_image_adjustments(SheetSelector::name(SOURCE_SHEET), ImageSelector::index(0),)?,
        original_package
            .sheet_image_adjustments(SheetSelector::name(SOURCE_SHEET), ImageSelector::index(0),)?
    );
    assert_data_assets_unchanged(&original, &resaved)?;

    let noop = package.move_table(DESTINATION_SHEET, SOURCE_TABLE, DESTINATION_SHEET)?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, resaved);
    let reopened = Package::from_bytes(&resaved)?;
    assert_relocated(&reopened, &moved, &destination, MOVED_MARKER)?;
    Ok(())
}
