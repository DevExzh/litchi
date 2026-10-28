//! Keynote slide-table cell parity cases (wave W1-C, record 0372): cell
//! writes through `slide::drawable::Edit::set_table_cell`.

use std::hint::black_box;

use litchi_iwa::keynote;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::comment::AuthorId;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_protos::{tsd, tsk, tsp};
use litchi_keynote::slide::drawable::{Kind, TableCellValue};
use litchi_keynote::{DrawableSelector, Package, SlideSelector, TableSelector};
use prost::Message as _;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_drawable_cases::{leaf_position, leaf_run};
use crate::known_fields::{Decoded, top_level_fields};
use crate::leaf_bytes;
use crate::text_glue::Workspace;

const TABLE: &str = "iwork/keynote/table-discovery.key";
const SLIDE: usize = 0;
const COMMENT_ROW: usize = 0;
const COMMENT_COLUMN: usize = 0;
const AUTHOR_MESSAGE_TYPE: u32 = 212;

/// Normalize identities only for the object created by the root-create case.
///
/// Replacement and read cases register [`decode_comment_storage`] instead;
/// their source-owned date and UUID are part of the parity contract and must
/// remain observable.  Keeping the creation decoder out of those cases is
/// important: a blanket decoder would hide a leaf that rewrote an untouched
/// comment's identity.
fn masked_new_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tsd::CommentStorageArchive::decode(data)?;
    if message.creation_date.is_some() {
        message.creation_date = Some(tsp::Date { seconds: 0.0 });
    }
    if message.storage_uuid.is_some() {
        message.storage_uuid = Some(tsp::Uuid { lower: 1, upper: 1 });
    }
    let mut decoded =
        crate::known_fields::decode_as::<tsd::CommentStorageArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn decode_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let decoded = crate::known_fields::decode_as::<tsd::CommentStorageArchive>(data)?;
    Ok(decoded)
}

fn mutate_unrelated_properties(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == "Metadata/Properties.plist")
        .ok_or("Keynote source has no properties metadata member")?;
    let mut replacement = entry.data().to_vec();
    let start = replacement
        .windows(b"documentUUID".len())
        .position(|window| window == b"documentUUID")
        .ok_or("properties metadata has no document UUID")?;
    let index = replacement
        .iter()
        .enumerate()
        .skip(start + b"documentUUID".len())
        .find_map(|(index, byte)| byte.is_ascii_hexdigit().then_some(index))
        .ok_or("document UUID has no hexadecimal payload")?;
    replacement[index] = if replacement[index] == b'0' {
        b'1'
    } else {
        b'0'
    };
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(entry.name(), &replacement)],
        Default::default(),
    )?)
}

fn comment_position_at(row: usize, column: usize) -> TestResult<At> {
    Ok(At::new(u32::try_from(row)?, u32::try_from(column)?))
}

fn host_comment_thread_at(bytes: &[u8], row: usize, column: usize) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let comment = editor.slide_table_cell_comment(SLIDE, model, row, column)?;
    let replies = if comment.is_some() {
        editor.slide_table_cell_comment_replies(SLIDE, model, row, column)?
    } else {
        Vec::new()
    };
    Ok(format!(
        "root={:?};replies={:?}",
        comment.as_ref().map(|value| value.comment.text.as_str()),
        replies
            .iter()
            .map(|reply| reply.comment.text.as_str())
            .collect::<Vec<_>>()
    ))
}

fn leaf_comment_thread_at(bytes: &[u8], row: usize, column: usize) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let position = comment_position_at(row, column)?;
    let comment = package.slide_table_cell_comment(
        SlideSelector::index(SLIDE),
        TableSelector::index(0),
        position,
    )?;
    let replies = package.slide_table_cell_comment_replies(
        SlideSelector::index(SLIDE),
        TableSelector::index(0),
        position,
    )?;
    Ok(format!(
        "root={:?};replies={:?}",
        comment.as_ref().map(|value| value.text()),
        replies.iter().map(|reply| reply.text()).collect::<Vec<_>>()
    ))
}

fn populated_comment_fixture(with_reply: bool) -> TestResult<Fixture> {
    // Build the source through the host first, outside the differential
    // Case.  This is deliberately not the leaf's own writer: clear and read
    // must consume the same host-authored populated bytes.
    let mut editor = keynote::Builder::new().title("W1-G comments").build()?;
    editor.add_slide_table(
        SLIDE,
        "Review",
        2,
        2,
        DrawablePoint { x: 120.0, y: 180.0 },
        DrawableSize {
            width: 840.0,
            height: 360.0,
        },
    )?;
    let model = host_model(&editor)?;
    editor.set_slide_table_cell_comment(
        SLIDE,
        model,
        COMMENT_ROW,
        COMMENT_COLUMN,
        "W1-G read root",
    )?;
    if with_reply {
        editor.add_slide_table_cell_comment_reply(
            SLIDE,
            model,
            COMMENT_ROW,
            COMMENT_COLUMN,
            "W1-G read reply",
        )?;
    }
    Ok(Fixture {
        path: if with_reply {
            "generated:keynote-table-comment-root-and-reply"
        } else {
            "generated:keynote-table-comment-root"
        },
        bytes: editor.to_bytes()?,
    })
}

fn comment_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    mask_new_metadata: bool,
    host: impl FnOnce(&mut keynote::Draft, u64) -> TestResult + 'static,
    leaf: impl FnOnce(Package, At) -> TestResult<LeafRun> + 'static,
) -> TestResult {
    comment_case_at(
        name,
        tier,
        evidence,
        input,
        COMMENT_ROW,
        COMMENT_COLUMN,
        mask_new_metadata,
        host,
        leaf,
    )
}

fn comment_case_at(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    row: usize,
    column: usize,
    mask_new_metadata: bool,
    host: impl FnOnce(&mut keynote::Draft, u64) -> TestResult + 'static,
    leaf: impl FnOnce(Package, At) -> TestResult<LeafRun> + 'static,
) -> TestResult {
    let mut case = Case::new(name, Format::Keynote, tier, evidence, input)
        .host(move |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            host(&mut editor, model)?;
            let bytes = editor.to_bytes()?;
            Ok(bytes)
        })
        .leaf(move |source| {
            let package = Package::from_bytes(source)?;
            let result = leaf(package, comment_position_at(row, column)?)?;
            Ok(result)
        })
        .leaf_structural_metadata()
        .read("host slide-table comment thread", move |bytes| {
            host_comment_thread_at(bytes, row, column)
        })
        .read("leaf slide-table comment thread", move |bytes| {
            leaf_comment_thread_at(bytes, row, column)
        });
    if mask_new_metadata {
        case = case.known_fields(|known| {
            known.with(
                3_056,
                "TSD.CommentStorageArchive (new-object date and storage UUID masked)",
                masked_new_comment_storage,
            )
        });
    } else {
        case = case.known_fields(|known| {
            known.with(
                3_056,
                "TSD.CommentStorageArchive (source date and storage UUID preserved)",
                decode_comment_storage,
            )
        });
    }
    case.run()?;
    Ok(())
}

fn comment_read_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    row: usize,
    column: usize,
    require_reply: bool,
) -> TestResult {
    Case::new(name, Format::Keynote, tier, evidence, input)
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host slide-table comment thread", move |bytes| {
            let editor = keynote::Draft::from_bytes(bytes)?;
            let model = host_model(&editor)?;
            let comment = editor
                .slide_table_cell_comment(SLIDE, model, row, column)?
                .ok_or("host read did not find the populated root comment")?;
            if comment.comment.text.is_empty()
                || comment.comment.creation_date_seconds.is_none()
                || comment.comment.author_id.is_none()
            {
                return Err("host root comment is missing text/date/author metadata".into());
            }
            let replies = editor.slide_table_cell_comment_replies(SLIDE, model, row, column)?;
            if require_reply && replies.is_empty() {
                return Err("host read did not find a populated reply".into());
            }
            if !require_reply && !replies.is_empty() {
                return Err("host root read unexpectedly found replies".into());
            }
            let reply_metadata = replies
                .iter()
                .map(|reply| {
                    if reply.comment.text.is_empty()
                        || reply.comment.creation_date_seconds.is_none()
                        || reply.comment.author_id.is_none()
                    {
                        return Err::<String, Box<dyn std::error::Error>>(
                            "host reply is missing text/date/author metadata".into(),
                        );
                    }
                    Ok(format!(
                        "{}|{:?}|{}",
                        reply.comment.text,
                        reply.comment.creation_date_seconds,
                        reply.comment.author_id.is_some()
                    ))
                })
                .collect::<TestResult<Vec<_>>>()?;
            Ok(format!(
                "root={}|{:?}|{};replies={reply_metadata:?}",
                comment.comment.text,
                comment.comment.creation_date_seconds,
                comment.comment.author_id.is_some()
            ))
        })
        .read("leaf slide-table comment thread", move |bytes| {
            let package = Package::from_bytes(bytes)?;
            let comment = package
                .slide_table_cell_comment(
                    SlideSelector::index(SLIDE),
                    TableSelector::index(0),
                    comment_position_at(row, column)?,
                )?
                .ok_or("leaf read did not find the populated root comment")?;
            if comment.text().is_empty()
                || comment.timestamp().is_none()
                || comment.author().is_none()
            {
                return Err("leaf root comment is missing text/date/author metadata".into());
            }
            let replies = package.slide_table_cell_comment_replies(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                comment_position_at(row, column)?,
            )?;
            if require_reply && replies.is_empty() {
                return Err("leaf read did not find a populated reply".into());
            }
            if !require_reply && !replies.is_empty() {
                return Err("leaf root read unexpectedly found replies".into());
            }
            let reply_metadata = replies
                .iter()
                .map(|reply| {
                    if reply.text().is_empty()
                        || reply.timestamp().is_none()
                        || reply.author().is_none()
                    {
                        return Err::<String, Box<dyn std::error::Error>>(
                            "leaf reply is missing text/date/author metadata".into(),
                        );
                    }
                    Ok(format!(
                        "{}|{:?}|{}",
                        reply.text(),
                        reply.timestamp().map(|value| value.as_f64()),
                        reply.author().is_some()
                    ))
                })
                .collect::<TestResult<Vec<_>>>()?;
            Ok(format!(
                "root={}|{:?}|{};replies={reply_metadata:?}",
                comment.text(),
                comment.timestamp().map(|value| value.as_f64()),
                comment.author().is_some()
            ))
        })
        .run()?;
    Ok(())
}

fn author_metadata(
    bytes: &[u8],
    author_id: Option<AuthorId>,
) -> TestResult<(Option<String>, Option<String>)> {
    let Some(author_id) = author_id else {
        return Ok((None, None));
    };
    let located = Workspace::open(bytes)?.message(author_id.get(), AUTHOR_MESSAGE_TYPE)?;
    let author = tsk::AnnotationAuthorArchive::decode(located.payload.as_slice())?;
    Ok((author.name, author.public_id))
}

/// Compare host and leaf readers against the same source bytes.  The
/// differential harness also reads each route's output independently, which
/// cannot prove that the two semantic readers agree on an untouched source.
fn assert_comment_semantics(
    bytes: &[u8],
    row: usize,
    column: usize,
    require_reply: bool,
) -> TestResult {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let host = editor
        .slide_table_cell_comment(SLIDE, model, row, column)?
        .ok_or("host read did not find the populated root comment")?;
    let host_replies = editor.slide_table_cell_comment_replies(SLIDE, model, row, column)?;
    if require_reply && host_replies.is_empty() {
        return Err("host source must contain a non-empty reply list".into());
    }
    if !require_reply && !host_replies.is_empty() {
        return Err("host source unexpectedly contains replies".into());
    }

    let package = Package::from_bytes(bytes)?;
    let position = comment_position_at(row, column)?;
    let leaf = package
        .slide_table_cell_comment(
            SlideSelector::index(SLIDE),
            TableSelector::index(0),
            position,
        )?
        .ok_or("leaf read did not find the populated root comment")?;
    let leaf_replies = package.slide_table_cell_comment_replies(
        SlideSelector::index(SLIDE),
        TableSelector::index(0),
        comment_position_at(row, column)?,
    )?;
    if require_reply && leaf_replies.is_empty() {
        return Err("leaf source must contain a non-empty reply list".into());
    }
    if !require_reply && !leaf_replies.is_empty() {
        return Err("leaf source unexpectedly contains replies".into());
    }

    assert_eq!(host.comment.text, leaf.text(), "root text");
    assert_eq!(
        host.comment.creation_date_seconds,
        leaf.timestamp().map(|value| value.as_f64()),
        "root timestamp"
    );
    assert_eq!(
        author_metadata(bytes, host.comment.author_id)?,
        {
            let author = leaf.author().ok_or("leaf root author is missing")?;
            (
                author.display_name().map(str::to_owned),
                author.public_id().map(str::to_owned),
            )
        },
        "root author metadata"
    );
    assert_eq!(host_replies.len(), leaf_replies.len(), "reply count");
    for (index, (host_reply, leaf_reply)) in
        host_replies.iter().zip(leaf_replies.iter()).enumerate()
    {
        assert_eq!(
            host_reply.comment.text,
            leaf_reply.text(),
            "reply {index} text"
        );
        assert_eq!(
            host_reply.comment.creation_date_seconds,
            leaf_reply.timestamp().map(|value| value.as_f64()),
            "reply {index} timestamp"
        );
        assert_eq!(
            author_metadata(bytes, host_reply.comment.author_id)?,
            {
                let author = leaf_reply
                    .author()
                    .ok_or_else(|| format!("leaf reply {index} author is missing"))?;
                (
                    author.display_name().map(str::to_owned),
                    author.public_id().map(str::to_owned),
                )
            },
            "reply {index} author metadata"
        );
    }
    Ok(())
}

fn host_model(editor: &keynote::Draft) -> TestResult<u64> {
    Ok(editor
        .slide_tables(SLIDE)?
        .first()
        .ok_or("no slide table")?
        .model_object_id)
}

/// Host read: every cell of the first slide table.
fn host_cells(bytes: &[u8]) -> TestResult<Vec<((usize, usize), String)>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let table = editor.slide_table(SLIDE, host_model(&editor)?)?;
    Ok(table
        .iter_cells()
        .map(|(at, value)| (at, format!("{value:?}")))
        .collect())
}

/// Leaf read: every cell of the first slide table.
fn leaf_cells(bytes: &[u8]) -> TestResult<Vec<((u32, u32), String)>> {
    let package = Package::from_bytes(bytes)?;
    let table = package.slide_table_cells(SlideSelector::index(SLIDE), TableSelector::index(0))?;
    Ok(table
        .iter_cells()
        .map(|cell| {
            (
                (cell.position().row(), cell.position().column()),
                format!("{:?}", cell.value()),
            )
        })
        .collect())
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft) -> TestResult,
    leaf: impl FnOnce(&[u8]) -> TestResult<LeafRun>,
) -> TestResult {
    let input = fixture(TABLE)?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(leaf)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("host cells", host_cells)
        .read("leaf cells", leaf_cells)
        .run()?;
    Ok(())
}

/// The first stored cell of the first slide table (a cleared absent cell
/// is not a clear).
fn first_stored(editor: &keynote::Draft) -> TestResult<(usize, usize)> {
    let table = editor.slide_table(SLIDE, host_model(editor)?)?;
    Ok(table
        .iter_cells()
        .map(|(position, _)| position)
        .next()
        .ok_or("table has no stored cell")?)
}

fn leaf_set(source: &[u8], cells: &[(u32, u32, TableCellValue)]) -> TestResult<LeafRun> {
    let table = leaf_position(source, Kind::Table)?;
    leaf_run(source, |edit| {
        for (row, column, value) in cells {
            edit.set_table_cell(
                SlideSelector::index(SLIDE),
                DrawableSelector::index(table),
                *row,
                *column,
                value.clone(),
            )?;
        }
        Ok(())
    })
}

#[test]
fn keynote_slide_table_comment_set_parity() -> TestResult {
    let input = fixture(TABLE)?;
    comment_case(
        "keynote_slide_table_comment_set_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        true,
        |editor, model| {
            let table = editor.slide_table(SLIDE, model)?;
            if table.get_cell(COMMENT_ROW, COMMENT_COLUMN).is_none() {
                return Err("comment-create case must target an existing stored cell".into());
            }
            editor.set_slide_table_cell_comment(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                "W1-G root",
            )?;
            Ok(())
        },
        |package, position| {
            let commit = package.set_slide_table_cell_comment(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
                "W1-G root",
            )?;
            let restored = commit
                .package()
                .apply_slide_table_cell_comment(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    )
}

#[test]
fn keynote_slide_table_comment_clear_parity() -> TestResult {
    let input = fixture("iwork/keynote/slide-table-read-native.key")?;
    comment_case_at(
        "keynote_slide_table_comment_clear_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        1,
        3,
        false,
        |editor, model| {
            editor.clear_slide_table_cell_comment(SLIDE, model, 1, 3)?;
            Ok(())
        },
        |package, position| {
            let clear = package.clear_slide_table_cell_comment(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
            )?;
            let undo_clear = clear
                .package()
                .apply_slide_table_cell_comment(&clear.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(clear.package()),
                leaf_bytes!(undo_clear.package()),
            ))
        },
    )
}

#[test]
fn keynote_slide_table_comment_replace_preserves_source_metadata() -> TestResult {
    let input = fixture("iwork/keynote/slide-table-read-native.key")?;
    comment_case_at(
        "keynote_slide_table_comment_replace_preserves_source_metadata",
        Tier::B,
        Evidence::Queued,
        &input,
        1,
        3,
        false,
        |editor, model| {
            editor.set_slide_table_cell_comment(SLIDE, model, 1, 3, "W1-G replacement")?;
            Ok(())
        },
        |package, position| {
            let commit = package.set_slide_table_cell_comment(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
                "W1-G replacement",
            )?;
            let restored = commit
                .package()
                .apply_slide_table_cell_comment(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    )
}

#[test]
fn keynote_slide_table_comment_patch_rejects_unrelated_metadata_change() -> TestResult {
    let input = fixture("iwork/keynote/slide-table-read-native.key")?;
    let package = Package::from_bytes(&input.bytes)?;
    let commit = package.set_slide_table_cell_comment(
        SlideSelector::index(SLIDE),
        TableSelector::index(0),
        comment_position_at(1, 3)?,
        "W1-G metadata conflict",
    )?;
    let stale_bytes = mutate_unrelated_properties(&input.bytes)?;
    let stale = Package::from_bytes(&stale_bytes)?;
    let error = stale.apply_slide_table_cell_comment(commit.patch());
    assert!(matches!(
        error,
        Err(litchi_keynote::slide::table::comments::Error::PatchConflict)
    ));
    assert_eq!(leaf_bytes!(stale), stale_bytes);
    Ok(())
}

/// Full Keynote facade cost probe for replacing a source-owned root comment.
/// The input is a native-authored root with stable metadata, so the probe
/// measures the same existing-comment rewrite on both routes.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_comment_cost_abba() -> TestResult {
    let input = fixture("iwork/keynote/slide-table-read-native.key")?;
    let position = comment_position_at(1, 3)?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.set_slide_table_cell_comment(SLIDE, model, 1, 3, "W1-G measured")?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package.set_slide_table_cell_comment(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
                "W1-G measured",
            )?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Full Keynote facade cost probe for reading a host-authored root/reply
/// thread.  The semantic projection is consumed by `black_box` so the probe
/// includes the reader work.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_comment_read_cost_abba() -> TestResult {
    let input = populated_comment_fixture(true)?;
    let position = comment_position_at(COMMENT_ROW, COMMENT_COLUMN)?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let comment = editor
                .slide_table_cell_comment(SLIDE, model, COMMENT_ROW, COMMENT_COLUMN)?
                .ok_or("host read did not find the populated root comment")?;
            let replies = editor.slide_table_cell_comment_replies(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
            )?;
            black_box((comment.comment.text.len(), replies.len()));
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let comment = package
                .slide_table_cell_comment(
                    SlideSelector::index(SLIDE),
                    TableSelector::index(0),
                    position,
                )?
                .ok_or("leaf read did not find the populated root comment")?;
            let replies = package.slide_table_cell_comment_replies(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
            )?;
            black_box((comment.text().len(), replies.len()));
            Ok(leaf_bytes!(package))
        },
    )
}

/// Full Keynote facade cost probe for creating a root comment on the stored
/// target cell used by the existing root-create parity case.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_comment_create_cost_abba() -> TestResult {
    let input = fixture(TABLE)?;
    let position = comment_position_at(COMMENT_ROW, COMMENT_COLUMN)?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.set_slide_table_cell_comment(
                SLIDE,
                model,
                COMMENT_ROW,
                COMMENT_COLUMN,
                "W1-G create",
            )?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package.set_slide_table_cell_comment(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
                "W1-G create",
            )?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Full Keynote facade cost probe for clearing a host-authored root comment.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn keynote_slide_table_comment_clear_cost_abba() -> TestResult {
    let input = populated_comment_fixture(false)?;
    let position = comment_position_at(COMMENT_ROW, COMMENT_COLUMN)?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.clear_slide_table_cell_comment(SLIDE, model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let clear = package.clear_slide_table_cell_comment(
                SlideSelector::index(SLIDE),
                TableSelector::index(0),
                position,
            )?;
            Ok(leaf_bytes!(clear.package()))
        },
    )
}

#[test]
fn keynote_slide_table_comment_read_parity() -> TestResult {
    let input = fixture("iwork/keynote/slide-table-read-native.key")?;
    assert_comment_semantics(&input.bytes, 1, 3, false)?;
    comment_read_case(
        "keynote_slide_table_comment_read_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        1,
        3,
        false,
    )
}

#[test]
fn keynote_slide_table_comment_replies_read_parity() -> TestResult {
    let input = populated_comment_fixture(true)?;
    assert_comment_semantics(&input.bytes, COMMENT_ROW, COMMENT_COLUMN, true)?;
    comment_read_case(
        "keynote_slide_table_comment_replies_read_parity",
        Tier::A,
        Evidence::Queued,
        &input,
        COMMENT_ROW,
        COMMENT_COLUMN,
        true,
    )
}

#[test]
fn keynote_slide_table_cell_set_parity() -> TestResult {
    run(
        "keynote_slide_table_cell_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let model = host_model(editor)?;
            editor.set_slide_table_cell(
                SLIDE,
                model,
                1,
                1,
                keynote::TableCellValue::Text("Parity".into()),
            )?;
            Ok(())
        },
        |source| leaf_set(source, &[(1, 1, TableCellValue::Text("Parity".into()))]),
    )
}

#[test]
fn keynote_slide_table_cells_set_parity() -> TestResult {
    run(
        "keynote_slide_table_cells_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let model = host_model(editor)?;
            editor.set_slide_table_cells(
                SLIDE,
                model,
                [
                    keynote::TableCellUpdate::new(
                        1,
                        1,
                        keynote::TableCellValue::Number(42.5.try_into()?),
                    ),
                    keynote::TableCellUpdate::new(2, 1, keynote::TableCellValue::Boolean(true)),
                ],
            )?;
            Ok(())
        },
        |source| {
            leaf_set(
                source,
                &[
                    (1, 1, TableCellValue::Number(42.5)),
                    (2, 1, TableCellValue::Boolean(true)),
                ],
            )
        },
    )
}

#[test]
fn keynote_slide_table_cell_clear_parity() -> TestResult {
    run(
        "keynote_slide_table_cell_clear_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let model = host_model(editor)?;
            let (row, column) = first_stored(editor)?;
            editor.clear_slide_table_cell(SLIDE, model, row, column)?;
            Ok(())
        },
        |source| {
            let (row, column) = first_stored(&keynote::Draft::from_bytes(source)?)?;
            leaf_set(
                source,
                &[(
                    u32::try_from(row)?,
                    u32::try_from(column)?,
                    TableCellValue::Empty,
                )],
            )
        },
    )
}
