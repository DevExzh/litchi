//! Self-tests: the harness must fail, readably, on every difference class
//! it claims to detect. The "routes" here are synthetic payload rewrites of
//! a native fixture, not host or leaf operations.

use litchi_iwa_archive::Limits as ArchiveLimits;
use litchi_iwa_archive::iwa::{Archive, RawMessage, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_protos::tst;
use prost::Message as _;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::top_level_fields;

const INPUT: &str = "iwork/numbers/table-merges-native.numbers";
const TABLE_MODEL: u32 = 6_001;
const TILE: u32 = 6_002;

/// Rewrite the payload of the first message of `message_type`.
fn rewrite(
    source: &[u8],
    message_type: u32,
    change: impl Fn(&[u8]) -> TestResult<Vec<u8>>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source, Default::default())?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data(), Default::default())?;
        let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
        let found = archive.objects.iter_mut().find_map(|object| {
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == message_type)?;
            Some((object, index))
        });
        let Some((object, index)) = found else {
            continue;
        };
        let data = change(&object.messages[index].data)?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            Default::default(),
        )?;
        let component =
            SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?
                .to_vec();
        return Ok(catalog.reassemble_to_bytes(
            &[EntryEdit::new(entry.name(), component.as_slice())],
            ArchiveLimits::default(),
        )?);
    }
    Err(format!("no message of type {message_type}").into())
}

/// Append an unknown varint field to a payload.
fn with_unknown_field(field: u32) -> impl Fn(&[u8]) -> TestResult<Vec<u8>> {
    move |data| {
        let mut data = data.to_vec();
        let mut tag = u64::from(field) << 3;
        while tag >= 0x80 {
            data.push((tag as u8) | 0x80);
            tag >>= 7;
        }
        data.push(tag as u8);
        data.push(1);
        Ok(data)
    }
}

fn renamed_table(data: &[u8]) -> TestResult<Vec<u8>> {
    let mut model = tst::TableModelArchive::decode(data)?;
    model.table_name = "renamed by the leaf".to_owned();
    Ok(model.encode_to_vec())
}

fn error_text(result: TestResult<crate::harness::Report>) -> String {
    match result {
        Ok(report) => panic!("{} unexpectedly passed", report.header),
        Err(error) => error.to_string(),
    }
}

#[test]
fn unknown_field_probe_is_unknown_to_the_schema() -> TestResult {
    let input = fixture(INPUT)?;
    let changed = rewrite(&input.bytes, TABLE_MODEL, with_unknown_field(9_001))?;
    assert_ne!(changed, input.bytes);
    assert!(top_level_fields(&[0xc8, 0xb2, 0x04, 0x01]).contains(&9_001));
    Ok(())
}

#[test]
fn tier_a_fails_and_tier_b_reports_when_only_unknown_fields_differ() -> TestResult {
    let input = fixture(INPUT)?;
    let host = || rewrite(&input.bytes, TABLE_MODEL, with_unknown_field(9_001));
    let leaf = || rewrite(&input.bytes, TABLE_MODEL, with_unknown_field(9_002));

    let failure = error_text(
        Case::new(
            "self_tier_a",
            Format::Numbers,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .host(|_| host())
        .leaf(|_| Ok(LeafRun::output(leaf()?)))
        .run(),
    );
    assert!(failure.contains("tier A: outputs differ"), "{failure}");
    assert!(failure.contains("equal on known fields"), "{failure}");

    let report = Case::new(
        "self_tier_b",
        Format::Numbers,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|_| host())
    .leaf(|_| Ok(LeafRun::output(leaf()?)))
    .run()?;
    assert!(report.diff.violations.is_empty());
    let notes = report.diff.notes.join("\n");
    assert!(notes.contains("TST.TableModelArchive"), "{notes}");
    assert!(notes.contains("9001") && notes.contains("9002"), "{notes}");
    Ok(())
}

#[test]
fn tier_b_fails_when_the_leaf_changes_an_untouched_object() -> TestResult {
    let input = fixture(INPUT)?;
    let failure = error_text(
        Case::new(
            "self_untouched",
            Format::Numbers,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(|source| rewrite(source, TABLE_MODEL, with_unknown_field(9_001)))
        .leaf(|source| {
            let model = rewrite(source, TABLE_MODEL, with_unknown_field(9_001))?;
            Ok(LeafRun::output(rewrite(
                &model,
                TILE,
                with_unknown_field(9_003),
            )?))
        })
        .run(),
    );
    assert!(
        failure.contains("the leaf changed an object the host left untouched"),
        "{failure}"
    );
    Ok(())
}

#[test]
fn tier_b_fails_on_a_known_field_difference_with_a_readable_diff() -> TestResult {
    let input = fixture(INPUT)?;
    let failure = error_text(
        Case::new(
            "self_known",
            Format::Numbers,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .host(|source| rewrite(source, TABLE_MODEL, with_unknown_field(9_001)))
        .leaf(|source| {
            Ok(LeafRun::output(rewrite(
                source,
                TABLE_MODEL,
                renamed_table,
            )?))
        })
        .read("table names", |bytes| {
            let package = litchi_numbers::Package::from_bytes(bytes)?;
            Ok(package.sheets()[0]
                .tables()
                .map(|table| table.name().to_owned())
                .collect::<Vec<_>>())
        })
        .run(),
    );
    assert!(failure.contains("known fields differ"), "{failure}");
    assert!(failure.contains("renamed by the leaf"), "{failure}");
    assert!(failure.contains("read `table names` differs"), "{failure}");
    Ok(())
}

#[test]
fn tier_b_fails_without_a_schema_for_a_changed_type() -> TestResult {
    let input = fixture(INPUT)?;
    let failure = error_text(
        Case::new(
            "self_schema",
            Format::Numbers,
            Tier::B,
            Evidence::Queued,
            &input,
        )
        .known_fields(|known| known.without(TABLE_MODEL))
        .host(|source| rewrite(source, TABLE_MODEL, with_unknown_field(9_001)))
        .leaf(|source| {
            Ok(LeafRun::output(rewrite(
                source,
                TABLE_MODEL,
                with_unknown_field(9_002),
            )?))
        })
        .run(),
    );
    assert!(
        failure.contains("no known-field schema is registered for message type 6001"),
        "{failure}"
    );
    Ok(())
}

#[test]
fn missing_native_evidence_inverse_and_noop_fail() -> TestResult {
    let input = fixture(INPUT)?;
    let failure = error_text(
        Case::new(
            "self_evidence",
            Format::Numbers,
            Tier::A,
            Evidence::Native("test-data/iwork/numbers/no-such-receipt.json"),
            &input,
        )
        .host(|source| Ok(source.to_vec()))
        .leaf(|source| Ok(LeafRun::with_inverse(source.to_vec(), Vec::new())))
        .run(),
    );
    assert!(failure.contains("native evidence"), "{failure}");
    assert!(
        failure.contains("leaf inverse does not restore"),
        "{failure}"
    );
    assert!(failure.contains("no-op"), "{failure}");
    Ok(())
}
