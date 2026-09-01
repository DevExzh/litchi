//! Contract tests for the deterministic Numbers data-format fixture.

#[path = "support/table_cell_data_format_fixture.rs"]
mod fixture;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_protos::tst;
use litchi_numbers_wire::BncCell;
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

#[test]
fn shared_and_unshared_modes_have_stable_format_ownership() -> TestResult {
    let shared = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    let shared_again = fixture::synthetic_package(fixture::FormatSharing::Shared)?;
    assert_eq!(shared, shared_again);

    let shared_list = format_list(&shared)?;
    assert_eq!(shared_list.entries.len(), 1);
    assert_eq!(shared_list.entries[0].key, fixture::FIRST_FORMAT_KEY);
    assert_eq!(shared_list.entries[0].refcount, 2);

    let unshared = fixture::synthetic_package(fixture::FormatSharing::Unshared)?;
    assert_ne!(shared, unshared);
    let unshared_list = format_list(&unshared)?;
    assert_eq!(
        unshared_list
            .entries
            .iter()
            .map(|entry| (entry.key, entry.refcount))
            .collect::<Vec<_>>(),
        vec![
            (fixture::FIRST_FORMAT_KEY, 1),
            (fixture::SECOND_FORMAT_KEY, 1)
        ]
    );

    let cells = tile_cells(&shared)?;
    assert_eq!(cells.len(), 2);
    let first = BncCell::parse(&cells[0])?;
    let second = BncCell::parse(&cells[1])?;
    assert_eq!(first.format_identifier(), Some(fixture::FIRST_FORMAT_KEY));
    assert_eq!(second.format_identifier(), Some(fixture::FIRST_FORMAT_KEY));
    Ok(())
}

#[test]
fn malformed_modes_are_deterministic_and_keep_unrelated_members() -> TestResult {
    for corruption in [
        fixture::Corruption::DuplicateFormatKey,
        fixture::Corruption::MissingFormatEntry,
        fixture::Corruption::RefcountMismatch,
        fixture::Corruption::DuplicateFormatList,
        fixture::Corruption::AliasedFormatList,
        fixture::Corruption::UnsupportedFormatType,
        fixture::Corruption::MalformedFormatPayload,
        fixture::Corruption::WrongCellFormatKey,
        fixture::Corruption::UnterminatedUnknownGroup,
    ] {
        let first = fixture::corrupted_package(corruption)?;
        let second = fixture::corrupted_package(corruption)?;
        assert_eq!(
            first, second,
            "corruption is not deterministic: {corruption:?}"
        );

        let catalog = Catalog::from_bytes(&first)?;
        let sentinel = catalog
            .iter()
            .find(|entry| entry.name() == fixture::SENTINEL_MEMBER)
            .ok_or("sentinel member is missing")?;
        assert_eq!(sentinel.data(), b"unrelated data-format sentinel");
        for member in fixture::PREVIEW_MEMBERS {
            assert!(catalog.iter().any(|entry| entry.name() == member));
        }
    }
    Ok(())
}

fn format_list(source: &[u8]) -> TestResult<tst::TableDataList> {
    // The first sidecar message is String, so find the Format message in the
    // archive rather than depending on a generated-message index.
    let archive = fixture::member_archive(source, fixture::TABLES_MEMBER)?;
    let sidecar = archive
        .object(fixture::SIDECAR_ID)
        .ok_or("format sidecar is missing")?;
    for message in &sidecar.messages {
        let list = tst::TableDataList::decode(message.data.as_slice())?;
        if list.list_type == tst::table_data_list::ListType::Format as i32 {
            return Ok(list);
        }
    }
    Err("format list is missing".into())
}

fn tile_cells(source: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let payload = fixture::object_message(
        source,
        fixture::TABLES_MEMBER,
        fixture::TILE_ID,
        fixture::TILE_TYPE,
    )?;
    let tile = tst::Tile::decode(payload.as_slice())?;
    let row = tile.row_infos.first().ok_or("format row is missing")?;
    let storage = row
        .cell_storage_buffer
        .as_deref()
        .ok_or("format storage is missing")?;
    let offsets = row
        .cell_offsets
        .as_deref()
        .ok_or("format offsets are missing")?;
    let mut cells = Vec::new();
    for index in 0..row.cell_count as usize {
        let start = usize::from(u16::from_le_bytes([
            offsets[index * 2],
            offsets[index * 2 + 1],
        ]));
        let end = if index + 1 == row.cell_count as usize {
            storage.len()
        } else {
            usize::from(u16::from_le_bytes([
                offsets[index * 2 + 2],
                offsets[index * 2 + 3],
            ]))
        };
        cells.push(storage[start..end].to_vec());
    }
    Ok(cells)
}
