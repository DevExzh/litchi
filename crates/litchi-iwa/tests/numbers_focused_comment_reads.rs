//! Differential read coverage for the focused Numbers comment owner.
//!
//! The semantic assertions use selectors, checked positions, and A1 addresses
//! only.  The second half of each positive assertion is an independent wire
//! oracle: it walks the test archive, decodes the native comment-storage and
//! author records, and compares the focused projection with the source-order
//! facts.  It deliberately does not call the deprecated raw-ID editor reads.

use std::io;

use litchi_iwa_archive::{
    Limits,
    iwa::{Archive, SnappyStream},
    package::{Catalog, EntryEdit},
};
use litchi_iwa_protos::{tsd, tsk, tsp, tst};
use litchi_numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableCellCommentError, TableCellCommentLimitKind, TableSelector,
};
use litchi_numbers_wire::BncCell;
use prost::Message as _;

#[path = "../../litchi-numbers/tests/support/numbers_comment_fixture.rs"]
mod fixture;
use fixture::{Corruption, FixtureMode};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const COMMENT_STORAGE_TYPE: u32 = fixture::COMMENT_STORAGE_TYPE;
const ANNOTATION_AUTHOR_TYPE: u32 = fixture::ANNOTATION_AUTHOR_TYPE;

#[derive(Debug, Clone, PartialEq)]
struct WireAuthor {
    display_name: Option<String>,
    public_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
struct WireComment {
    text: String,
    timestamp: Option<f64>,
    author: Option<WireAuthor>,
    replies: Vec<WireComment>,
}

fn member_archive(source: &[u8], name: &str) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other(format!("missing archive member {name}")))?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn rewrite_storage_payload<F>(source: &[u8], identifier: u64, rewrite: F) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut tsd::CommentStorageArchive),
{
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == fixture::DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("comment fixture document member is missing"))?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let object = archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or_else(|| io::Error::other(format!("comment storage {identifier} is missing")))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == COMMENT_STORAGE_TYPE)
        .ok_or_else(|| {
            io::Error::other(format!("comment storage {identifier} payload is missing"))
        })?;
    let mut payload = tsd::CommentStorageArchive::decode(message.data.as_slice())?;
    rewrite(&mut payload);
    message.data = payload.encode_to_vec();
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(
            fixture::DOCUMENT_MEMBER,
            compressed.as_slice(),
        )],
        Limits::default(),
    )?)
}

fn rewrite_author_payload<F>(source: &[u8], identifier: u64, rewrite: F) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut tsk::AnnotationAuthorArchive),
{
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == fixture::DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("comment fixture document member is missing"))?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let object = archive
        .objects
        .iter_mut()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or_else(|| io::Error::other(format!("author {identifier} is missing")))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == ANNOTATION_AUTHOR_TYPE)
        .ok_or_else(|| io::Error::other(format!("author {identifier} payload is missing")))?;
    let mut payload = tsk::AnnotationAuthorArchive::decode(message.data.as_slice())?;
    rewrite(&mut payload);
    message.data = payload.encode_to_vec();
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(
            fixture::DOCUMENT_MEMBER,
            compressed.as_slice(),
        )],
        Limits::default(),
    )?)
}

fn all_archives(source: &[u8]) -> TestResult<Vec<Archive>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut archives = Vec::new();
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        archives.push(Archive::parse(
            SnappyStream::decompress(entry.data())?.as_bytes(),
        )?);
    }
    Ok(archives)
}

fn storage_payload(source: &[u8], identifier: u64) -> TestResult<tsd::CommentStorageArchive> {
    for archive in all_archives(source)? {
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == COMMENT_STORAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.len() != 1 {
            return Err(io::Error::other(format!(
                "comment storage {identifier} has {} payloads",
                messages.len()
            ))
            .into());
        }
        return Ok(tsd::CommentStorageArchive::decode(
            messages[0].data.as_slice(),
        )?);
    }
    Err(io::Error::other(format!("comment storage {identifier} is missing")).into())
}

fn author_payload(source: &[u8], identifier: u64) -> TestResult<WireAuthor> {
    for archive in all_archives(source)? {
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == ANNOTATION_AUTHOR_TYPE)
            .collect::<Vec<_>>();
        if messages.len() != 1 {
            return Err(io::Error::other(format!(
                "author {identifier} has {} payloads",
                messages.len()
            ))
            .into());
        }
        let author = tsk::AnnotationAuthorArchive::decode(messages[0].data.as_slice())?;
        return Ok(WireAuthor {
            display_name: author.name,
            public_id: author.public_id,
        });
    }
    Err(io::Error::other(format!("author {identifier} is missing")).into())
}

fn wire_comment(source: &[u8], identifier: u64) -> TestResult<WireComment> {
    let comment = storage_payload(source, identifier)?;
    let author = comment
        .author
        .as_ref()
        .map(|reference| author_payload(source, reference.identifier))
        .transpose()?;
    let mut replies = Vec::new();
    replies.try_reserve_exact(comment.replies.len())?;
    for reply in &comment.replies {
        replies.push(wire_comment(source, reply.identifier)?);
    }
    Ok(WireComment {
        text: comment.text.unwrap_or_default(),
        timestamp: comment.creation_date.map(|date| date.seconds),
        author,
        replies,
    })
}

fn unique_comment_storage_identifier(source: &[u8]) -> TestResult<u64> {
    let mut identifiers = Vec::new();
    for archive in all_archives(source)? {
        for object in archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            for message in object.messages {
                if message.type_ == COMMENT_STORAGE_TYPE {
                    tsd::CommentStorageArchive::decode(message.data.as_slice())?;
                    identifiers.push(identifier);
                }
            }
        }
    }
    match identifiers.as_slice() {
        [identifier] => Ok(*identifier),
        [] => Err(io::Error::other("permanent comment storage is missing").into()),
        _ => Err(io::Error::other("permanent comment storage is ambiguous").into()),
    }
}

fn wire_comment_key(source: &[u8], row: usize) -> TestResult<Option<u32>> {
    let archive = member_archive(source, fixture::DOCUMENT_MEMBER)?;
    let mut keys = Vec::new();
    for object in &archive.objects {
        for payload in &object.messages {
            let Ok(tile) = tst::Tile::decode(payload.data.as_slice()) else {
                continue;
            };
            let Some(row_info) = tile.row_infos.get(row) else {
                continue;
            };
            let Some(bytes) = row_info.cell_storage_buffer.as_deref() else {
                continue;
            };
            let Ok(cell) = BncCell::parse(bytes) else {
                continue;
            };
            if let Some(key) = cell.comment_identifier() {
                keys.push(key);
            }
        }
    }
    match keys.as_slice() {
        [key] => Ok(Some(*key)),
        [] => Err(io::Error::other("comment fixture cell key is missing").into()),
        _ => Err(io::Error::other("comment fixture cell key is ambiguous").into()),
    }
}

fn wire_root_identifier(source: &[u8], row: usize) -> TestResult<u64> {
    let key = wire_comment_key(source, row)?
        .ok_or_else(|| io::Error::other("comment fixture cell has no comment key"))?;
    let archive = member_archive(source, fixture::DOCUMENT_MEMBER)?;
    let mut roots = Vec::new();
    for object in &archive.objects {
        for payload in &object.messages {
            if payload.type_ == fixture::TABLE_DATA_LIST_TYPE || payload.type_ == 6_201 {
                let Ok(list) = tst::TableDataList::decode(payload.data.as_slice()) else {
                    continue;
                };
                if list.list_type != tst::table_data_list::ListType::CommentStorage as i32 {
                    continue;
                }
                if let Some(entry) = list.entries.iter().find(|entry| entry.key == key) {
                    if let Some(reference) = entry.comment_storage.as_ref() {
                        roots.push(reference.identifier);
                    }
                }
                continue;
            }
            if payload.type_ != 6_011 {
                continue;
            }
            let Ok(segment) = tst::TableDataListSegment::decode(payload.data.as_slice()) else {
                continue;
            };
            if segment.list_type != tst::table_data_list::ListType::CommentStorage as i32 {
                continue;
            }
            if let Some(entry) = segment.entries.iter().find(|entry| entry.key == key) {
                if let Some(reference) = entry.comment_storage.as_ref() {
                    roots.push(reference.identifier);
                }
            }
        }
    }
    match roots.as_slice() {
        [identifier] => Ok(*identifier),
        [] => Err(io::Error::other("comment fixture list key is missing").into()),
        _ => Err(io::Error::other("comment fixture list key is ambiguous").into()),
    }
}

fn assert_focused_replies_match_wire(
    package: &Package,
    source: &[u8],
    row: usize,
    address: &str,
) -> TestResult {
    let root_identifier = wire_root_identifier(source, row)?;
    assert_focused_replies_match_wire_for_table(
        package,
        source,
        fixture::TABLE_NAME,
        address,
        root_identifier,
    )
}

fn assert_focused_replies_match_wire_for_table(
    package: &Package,
    source: &[u8],
    table: &str,
    address: &str,
    root_identifier: u64,
) -> TestResult {
    let expected = wire_comment(source, root_identifier)?;
    let root = package
        .table_cell_comment_a1(fixture::SHEET_NAME, table, address)?
        .ok_or_else(|| io::Error::other(format!("focused root comment is missing at {address}")))?;
    assert_metadata(&root, &expected);
    let replies = package.table_cell_comment_replies_a1(fixture::SHEET_NAME, table, address)?;
    assert_eq!(
        replies.len(),
        expected.replies.len(),
        "reply count at {address}"
    );
    for (reply, expected) in replies.iter().zip(expected.replies.iter()) {
        assert_metadata(reply, expected);
    }
    Ok(())
}

fn package_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn assert_metadata(comment: &impl CommentMetadata, expected: &WireComment) {
    assert_eq!(
        comment.text(),
        expected.text,
        "focused text differs from the independent storage payload"
    );
    assert_eq!(
        comment.timestamp().map(|timestamp| timestamp.as_f64()),
        expected.timestamp,
        "focused timestamp differs from the independent storage payload"
    );
    let focused_author = comment.author();
    match (&focused_author, &expected.author) {
        (None, None) => {},
        (Some(author), Some(expected)) => {
            assert_eq!(author.display_name(), expected.display_name.as_deref());
            assert_eq!(author.public_id(), expected.public_id.as_deref());
        },
        (focused, expected) => {
            panic!("focused author presence differs from wire oracle: {focused:?} vs {expected:?}")
        },
    }
}

trait CommentMetadata {
    fn text(&self) -> &str;
    fn timestamp(&self) -> Option<litchi_numbers::package::comments::CommentTimestamp>;
    fn author(&self) -> Option<&litchi_numbers::package::comments::CommentAuthor>;
}

impl CommentMetadata for litchi_numbers::TableCellComment {
    fn text(&self) -> &str {
        self.text()
    }

    fn timestamp(&self) -> Option<litchi_numbers::package::comments::CommentTimestamp> {
        self.timestamp()
    }

    fn author(&self) -> Option<&litchi_numbers::package::comments::CommentAuthor> {
        self.author()
    }
}

impl CommentMetadata for litchi_numbers::TableCellCommentReply {
    fn text(&self) -> &str {
        self.text()
    }

    fn timestamp(&self) -> Option<litchi_numbers::package::comments::CommentTimestamp> {
        self.timestamp()
    }

    fn author(&self) -> Option<&litchi_numbers::package::comments::CommentAuthor> {
        self.author()
    }
}

fn assert_source_unchanged(package: &Package, source: &[u8]) -> TestResult {
    assert_eq!(package_bytes(package)?, source);
    Ok(())
}

#[test]
fn focused_numbers_comment_reads_match_wire_oracle_by_selector_and_a1() -> TestResult {
    let source = fixture::fixture(FixtureMode::DuplicateText, None)?;
    let package = Package::from_bytes(&source)?;
    let root_identifier = wire_root_identifier(&source, 0)?;
    assert_eq!(wire_comment_key(&source, 0)?, Some(1));
    let expected = wire_comment(&source, root_identifier)?;
    assert_eq!(expected.replies.len(), 3);

    let root = package
        .table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .ok_or_else(|| io::Error::other("focused root comment is missing"))?;
    assert_metadata(&root, &expected);

    let named_root = package
        .table_cell_comment(
            fixture::SHEET_NAME,
            fixture::TABLE_NAME,
            CellPosition::from_a1("A1")?,
        )?
        .ok_or_else(|| io::Error::other("named A1 root comment is missing"))?;
    assert_eq!(named_root, root);
    let addressed_root = package
        .table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?
        .ok_or_else(|| io::Error::other("A1 root comment is missing"))?;
    assert_eq!(addressed_root, root);

    let replies = package.table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    )?;
    let reply_texts = replies.iter().map(|reply| reply.text()).collect::<Vec<_>>();
    assert_eq!(
        reply_texts,
        expected
            .replies
            .iter()
            .map(|reply| reply.text.as_str())
            .collect::<Vec<_>>()
    );
    for (reply, expected) in replies.iter().zip(expected.replies.iter()) {
        assert_metadata(reply, expected);
    }
    let addressed_replies =
        package.table_cell_comment_replies_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?;
    assert_eq!(addressed_replies, replies);
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_numbers_comment_reads_handle_absent_and_shared_roots() -> TestResult {
    let absent = fixture::fixture(FixtureMode::Rootless, None)?;
    let absent_package = Package::from_bytes(&absent)?;
    assert_eq!(
        absent_package.table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::from_a1("A1")?,
        )?,
        None
    );
    assert!(matches!(
        absent_package.table_cell_comment_replies_a1(
            SheetSelector::index(0),
            TableSelector::index(0),
            "A1",
        ),
        Err(TableCellCommentError::CommentNotFound { .. })
    ));
    assert_source_unchanged(&absent_package, &absent)?;

    let shared = fixture::fixture(FixtureMode::SharedRoot, None)?;
    let shared_package = Package::from_bytes(&shared)?;
    for address in ["A1", "A2"] {
        let comment = shared_package
            .table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, address)?
            .ok_or_else(|| io::Error::other(format!("shared root missing at {address}")))?;
        assert_eq!(comment.text(), "root comment");
    }
    assert_source_unchanged(&shared_package, &shared)?;

    Ok(())
}

#[test]
fn focused_numbers_shared_reply_reads_match_wire_oracle_for_each_root() -> TestResult {
    let source = fixture::fixture(FixtureMode::SharedReply, None)?;
    let package = Package::from_bytes(&source)?;
    assert_focused_replies_match_wire(&package, &source, 0, "A1")?;
    assert_focused_replies_match_wire(&package, &source, 1, "A2")?;
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_numbers_cross_component_reply_reads_match_wire_oracle() -> TestResult {
    let source = fixture::fixture(FixtureMode::CrossComponent, None)?;
    let package = Package::from_bytes(&source)?;
    assert_focused_replies_match_wire(&package, &source, 0, "A1")?;
    assert_focused_replies_match_wire(&package, &source, 1, "A2")?;
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_numbers_multitable_same_key_reads_use_selected_list_local_refcount() -> TestResult {
    let source = fixture::multitable_same_key_fixture()?;
    let package = Package::from_bytes(&source)?;
    assert_focused_replies_match_wire_for_table(
        &package,
        &source,
        fixture::TABLE_NAME,
        "A1",
        fixture::ROOT_COMMENT_ID,
    )?;
    assert_focused_replies_match_wire_for_table(
        &package,
        &source,
        fixture::SECOND_TABLE_NAME,
        "A1",
        fixture::SECOND_ROOT_COMMENT_ID,
    )?;
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_numbers_segmented_comment_reads_match_wire_oracle() -> TestResult {
    let source = fixture::fixture(FixtureMode::SegmentedRoot, None)?;
    let package = Package::from_bytes(&source)?;
    assert_focused_replies_match_wire(&package, &source, 0, "A1")?;
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_numbers_segmented_reads_pin_the_selected_parent_edge() -> TestResult {
    let source = fixture::segmented_with_unrelated_segment_fixture()?;
    let package = Package::from_bytes(&source)?;
    assert_focused_replies_match_wire_for_table(
        &package,
        &source,
        fixture::TABLE_NAME,
        "A1",
        fixture::ROOT_COMMENT_ID,
    )?;
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_numbers_read_accepts_native_style_missing_field_info_headers() -> TestResult {
    let source = fixture::fixture(FixtureMode::SingleRoot, Some(Corruption::FieldInfoMissing))?;
    let package = Package::from_bytes(&source)?;
    package
        .table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?
        .ok_or_else(|| io::Error::other("missing-field-info root comment is missing"))?;
    let replies = package
        .table_cell_comment_replies_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")
        .map_err(|error| io::Error::other(format!("missing-field-info replies: {error:?}")))?;
    assert_eq!(replies.len(), 1);
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_numbers_comment_reads_reject_malformed_graphs_atomically_and_keep_unknown_wire()
-> TestResult {
    for corruption in [
        Corruption::DuplicateReplyReference,
        Corruption::SelfReplyReference,
        Corruption::MissingReplyReference,
        Corruption::NestedReplyReference,
        Corruption::ExternalReplyReference,
        Corruption::TypedReplyReference,
        Corruption::MissingReply,
        Corruption::WrongReplyType,
        Corruption::SegmentedList,
    ] {
        let source = fixture::fixture(FixtureMode::SingleRoot, Some(corruption))?;
        let before = source.clone();
        match Package::from_bytes(&source) {
            Err(_) => {},
            Ok(package) => {
                assert!(
                    package
                        .table_cell_comment_replies_a1(
                            SheetSelector::index(0),
                            TableSelector::index(0),
                            "A1",
                        )
                        .is_err(),
                    "malformed {corruption:?} graph was published"
                );
                assert_source_unchanged(&package, &source)?;
            },
        }
        assert_eq!(
            source, before,
            "malformed source was mutated during admission"
        );
    }

    let unknown = fixture::fixture(FixtureMode::SingleRoot, Some(Corruption::UnknownWire))?;
    let unknown_package = Package::from_bytes(&unknown)?;
    match unknown_package.table_cell_comment_replies_a1(
        SheetSelector::index(0),
        TableSelector::index(0),
        "A1",
    ) {
        Ok(replies) => assert_eq!(
            replies.iter().map(|reply| reply.text()).collect::<Vec<_>>(),
            vec!["first reply"]
        ),
        Err(TableCellCommentError::InvalidSource { .. }) => {},
        Err(error) => return Err(io::Error::other(error.to_string()).into()),
    }
    assert_source_unchanged(&unknown_package, &unknown)?;
    Ok(())
}

#[test]
fn permanent_numbers_comment_fixture_preserves_metadata_and_source_on_read() -> TestResult {
    let source = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../litchi-numbers/tests/fixtures/comment-edit-root.numbers"
    ))
    .as_slice();
    let package = Package::from_bytes(source)?;
    let root_identifier = unique_comment_storage_identifier(source)?;
    let expected = wire_comment(source, root_identifier)?;
    let sheet = package
        .sheets()
        .first()
        .ok_or_else(|| io::Error::other("permanent fixture sheet is missing"))?;
    let table = sheet
        .tables()
        .next()
        .ok_or_else(|| io::Error::other("permanent fixture table is missing"))?;
    let root = package
        .table_cell_comment_a1(sheet.name(), table.name(), "B2")?
        .ok_or_else(|| io::Error::other("permanent root comment is missing"))?;
    assert_metadata(&root, &expected);
    assert!(
        root.author()
            .and_then(|author| author.display_name())
            .is_some()
    );
    assert!(root.timestamp().is_some());
    assert!(
        package
            .table_cell_comment_replies_a1(sheet.name(), table.name(), "B2")?
            .is_empty()
    );
    assert_source_unchanged(&package, source)?;
    Ok(())
}

#[test]
fn focused_comment_metadata_rejects_nonfinite_timestamps_atomically() -> TestResult {
    for seconds in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let source = rewrite_storage_payload(
            &fixture::fixture(FixtureMode::SingleRoot, None)?,
            fixture::ROOT_COMMENT_ID,
            |comment| comment.creation_date = Some(tsp::Date { seconds }),
        )?;
        let before = source.clone();
        match Package::from_bytes(&source) {
            Err(_) => {},
            Ok(package) => {
                assert!(matches!(
                    package.table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1",),
                    Err(TableCellCommentError::InvalidSource { .. })
                ));
                assert_source_unchanged(&package, &source)?;
            },
        }
        assert_eq!(source, before, "failed read must not mutate its source");
    }
    Ok(())
}

#[test]
fn focused_comment_metadata_rejects_nonfinite_reply_timestamps_before_projection() -> TestResult {
    for seconds in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let source = rewrite_storage_payload(
            &fixture::fixture(FixtureMode::SingleRoot, None)?,
            fixture::FIRST_REPLY_ID,
            |comment| comment.creation_date = Some(tsp::Date { seconds }),
        )?;
        let package = match Package::from_bytes(&source) {
            Err(_) => continue,
            Ok(package) => package,
        };
        let root = package
            .table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?
            .ok_or_else(|| io::Error::other("finite root comment is missing"))?;
        assert!(root.timestamp().is_some());
        assert!(matches!(
            package.table_cell_comment_replies_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1",),
            Err(TableCellCommentError::InvalidSource { .. })
        ));
        assert_source_unchanged(&package, &source)?;
    }
    Ok(())
}

#[test]
fn focused_comment_metadata_canonicalizes_signed_zero() -> TestResult {
    let source = rewrite_storage_payload(
        &fixture::fixture(FixtureMode::SingleRoot, None)?,
        fixture::ROOT_COMMENT_ID,
        |comment| comment.creation_date = Some(tsp::Date { seconds: -0.0 }),
    )?;
    let package = Package::from_bytes(&source)?;
    let comment = package
        .table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?
        .ok_or_else(|| io::Error::other("signed-zero root comment is missing"))?;
    let timestamp = comment
        .timestamp()
        .ok_or_else(|| io::Error::other("signed-zero timestamp is missing"))?;
    assert_eq!(timestamp.as_f64().to_bits(), 0.0_f64.to_bits());
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_comment_metadata_preserves_optional_absence_and_redacts_debug() -> TestResult {
    let source = rewrite_storage_payload(
        &fixture::fixture(FixtureMode::SingleRoot, None)?,
        fixture::ROOT_COMMENT_ID,
        |comment| {
            comment.author = None;
            comment.creation_date = None;
        },
    )?;
    let package = Package::from_bytes(&source)?;
    let comment = package
        .table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?
        .ok_or_else(|| io::Error::other("authorless root comment is missing"))?;
    assert!(comment.author().is_none());
    assert!(comment.timestamp().is_none());
    assert_source_unchanged(&package, &source)?;

    let source = fixture::fixture(FixtureMode::DuplicateText, None)?;
    let package = Package::from_bytes(&source)?;
    let root = package
        .table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?
        .ok_or_else(|| io::Error::other("metadata root comment is missing"))?;
    let replies =
        package.table_cell_comment_replies_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?;
    assert!(root.author().is_some());
    assert!(root.timestamp().is_some());
    assert!(
        replies
            .iter()
            .all(|reply| { reply.timestamp().is_some() && reply.author().is_some() })
    );
    for rendered in [format!("{root:?}"), format!("{:?}", replies[0])] {
        for secret in [
            "root comment",
            "first reply",
            "duplicate",
            "Reply fixture author",
            "reply-fixture-author",
        ] {
            assert!(
                !rendered.contains(secret),
                "debug output leaked comment metadata {secret:?}: {rendered}"
            );
        }
    }
    assert_source_unchanged(&package, &source)?;

    let missing_author =
        fixture::fixture(FixtureMode::SingleRoot, Some(Corruption::MissingAuthor))?;
    let before = missing_author.clone();
    match Package::from_bytes(&missing_author) {
        Err(_) => {},
        Ok(package) => {
            assert!(matches!(
                package.table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1",),
                Err(TableCellCommentError::InvalidSource { .. })
            ));
            assert_source_unchanged(&package, &missing_author)?;
        },
    }
    assert_eq!(missing_author, before);
    Ok(())
}

#[test]
fn focused_comment_metadata_keeps_present_authors_with_empty_optional_fields() -> TestResult {
    let source = rewrite_author_payload(
        &fixture::fixture(FixtureMode::SingleRoot, None)?,
        fixture::AUTHOR_ID,
        |author| {
            author.name = None;
            author.public_id = None;
            author.public_ids.clear();
        },
    )?;
    let package = Package::from_bytes(&source)?;
    let root = package
        .table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?
        .ok_or_else(|| io::Error::other("empty-field author root is missing"))?;
    let author = root
        .author()
        .ok_or_else(|| io::Error::other("present native author was dropped"))?;
    assert_eq!(author.display_name(), None);
    assert_eq!(author.public_id(), None);
    let replies =
        package.table_cell_comment_replies_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?;
    assert_eq!(replies.len(), 1);
    let reply_author = replies[0]
        .author()
        .ok_or_else(|| io::Error::other("reply author presence was dropped"))?;
    assert_eq!(reply_author.display_name(), None);
    assert_eq!(reply_author.public_id(), None);
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_comment_metadata_budget_counts_author_and_text_together() -> TestResult {
    let source = fixture::fixture(FixtureMode::DuplicateText, None)?;
    let expected = wire_comment(&source, wire_root_identifier(&source, 0)?)?;
    let one_reply_bytes = expected.replies[0].text.len()
        + expected.replies[0]
            .author
            .as_ref()
            .and_then(|author| author.display_name.as_ref())
            .map_or(0, String::len)
        + expected.replies[0]
            .author
            .as_ref()
            .and_then(|author| author.public_id.as_ref())
            .map_or(0, String::len);
    let all_reply_bytes = expected
        .replies
        .iter()
        .map(|reply| {
            reply.text.len()
                + reply
                    .author
                    .as_ref()
                    .and_then(|author| author.display_name.as_ref())
                    .map_or(0, String::len)
                + reply
                    .author
                    .as_ref()
                    .and_then(|author| author.public_id.as_ref())
                    .map_or(0, String::len)
        })
        .sum::<usize>();
    assert!(
        one_reply_bytes < 64,
        "fixture no longer has a bounded reply"
    );
    assert!(
        all_reply_bytes > 64,
        "fixture no longer exercises aggregate output"
    );

    let semantic = PackageSemanticLimits::default()
        .with_projection_limits(PackageSemanticLimits::MAX_MATERIALIZED_CELLS, 64)?;
    let package = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(PackageLimits::default(), semantic),
    )?;
    assert!(matches!(
        package.table_cell_comment_replies_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1"),
        Err(TableCellCommentError::LimitExceeded {
            kind: TableCellCommentLimitKind::TextBytes,
            ..
        })
    ));
    assert_source_unchanged(&package, &source)?;
    Ok(())
}

#[test]
fn focused_comment_metadata_ignores_repeated_public_ids_when_bounded() -> TestResult {
    let base = fixture::fixture(FixtureMode::SingleRoot, None)?;
    let source = rewrite_author_payload(&base, fixture::AUTHOR_ID, |author| {
        author.public_ids = (0..8)
            .map(|index| format!("ignored-public-id-{index}-{}", "x".repeat(128)))
            .collect();
    })?;
    let semantic = PackageSemanticLimits::default()
        .with_projection_limits(PackageSemanticLimits::MAX_MATERIALIZED_CELLS, 64)?;
    let package = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(PackageLimits::default(), semantic),
    )?;
    let comment = package
        .table_cell_comment_a1(fixture::SHEET_NAME, fixture::TABLE_NAME, "A1")?
        .ok_or_else(|| io::Error::other("bounded author root comment is missing"))?;
    let author = comment
        .author()
        .ok_or_else(|| io::Error::other("bounded author metadata is missing"))?;
    assert_eq!(author.display_name(), Some("Reply fixture author"));
    assert_eq!(author.public_id(), Some("reply-fixture-author"));
    assert_source_unchanged(&package, &source)?;
    Ok(())
}
