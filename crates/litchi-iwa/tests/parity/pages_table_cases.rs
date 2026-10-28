//! Pages body-table parity cases of wave W1-B (record 0371): cell values
//! through the leaf `table::values` family and the cell reads.

use std::hint::black_box;

use litchi_iwa::pages::{self, CellValue, TableCellUpdate};
use litchi_iwa_common::comment::AuthorId;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_protos::tsk;
use litchi_pages::Package;
use litchi_pages::table::values::Value;
use prost::Message as _;

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;
use crate::text_glue::Workspace;

const CELLS: &str = "iwork/pages/body-table-cells-native.pages";
/// Native two-table source without formulas (the values family refuses
/// cells that take part in a formula graph, as `body-table-cells-native`
/// has).
const CATALOG: &str = "iwork/pages/body-table-catalog-native.pages";
const TABLE: usize = 0;
const COMMENT_ROW: usize = 1;
const COMMENT_COLUMN: usize = 1;
const AUTHOR_MESSAGE_TYPE: u32 = 212;

fn comment_position() -> TestResult<At> {
    Ok(At::new(
        u32::try_from(COMMENT_ROW)?,
        u32::try_from(COMMENT_COLUMN)?,
    ))
}

fn host_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let comment = editor.table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?;
    let replies = if comment.is_some() {
        editor.table_cell_comment_replies(model, COMMENT_ROW, COMMENT_COLUMN)?
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

fn leaf_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let position = comment_position()?;
    let comment = package.body_table_cell_comment(TABLE, position)?;
    let replies = if comment.is_some() {
        package.body_table_cell_comment_replies(TABLE, position)?
    } else {
        Box::new([])
    };
    Ok(format!(
        "root={:?};replies={:?}",
        comment.as_ref().map(|value| value.text()),
        replies.iter().map(|reply| reply.text()).collect::<Vec<_>>()
    ))
}

fn host_populated_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let comment = editor
        .table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?
        .ok_or("host read did not find the populated root comment")?;
    if comment.comment.text.is_empty()
        || comment.comment.creation_date_seconds.is_none()
        || comment.comment.author_id.is_none()
        || comment.comment.storage_uuid.is_none()
    {
        return Err("host root comment is missing text/date/author/storage metadata".into());
    }
    let replies = editor.table_cell_comment_replies(model, COMMENT_ROW, COMMENT_COLUMN)?;
    if replies.is_empty() {
        return Err("host read did not find a populated reply".into());
    }
    let reply_metadata = replies
        .iter()
        .map(|reply| {
            if reply.comment.text.is_empty()
                || reply.comment.creation_date_seconds.is_none()
                || reply.comment.author_id.is_none()
                || reply.comment.storage_uuid.is_none()
            {
                return Err::<String, Box<dyn std::error::Error>>(
                    "host reply is missing text/date/author/storage metadata".into(),
                );
            }
            Ok(format!(
                "{}|{:?}|{}|{}",
                reply.comment.text,
                reply.comment.creation_date_seconds,
                reply.comment.author_id.is_some(),
                reply.comment.storage_uuid.is_some()
            ))
        })
        .collect::<TestResult<Vec<_>>>()?;
    Ok(format!(
        "root={}|{:?}|{}|{};replies={reply_metadata:?}",
        comment.comment.text,
        comment.comment.creation_date_seconds,
        comment.comment.author_id.is_some(),
        comment.comment.storage_uuid.is_some()
    ))
}

fn leaf_populated_comment_thread(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let position = comment_position()?;
    let comment = package
        .body_table_cell_comment(TABLE, position)?
        .ok_or("leaf read did not find the populated root comment")?;
    if comment.text().is_empty() || comment.timestamp().is_none() || comment.author().is_none() {
        return Err("leaf root comment is missing text/date/author metadata".into());
    }
    let replies = package.body_table_cell_comment_replies(TABLE, position)?;
    if replies.is_empty() {
        return Err("leaf read did not find a populated reply".into());
    }
    let reply_metadata = replies
        .iter()
        .map(|reply| {
            if reply.text().is_empty() || reply.timestamp().is_none() || reply.author().is_none() {
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

/// Compare both readers on the same host-authored source.  `Case::read`
/// intentionally invokes each reader on its own host/leaf output, so that
/// harness result is not enough to establish semantic reader parity.
fn assert_populated_comment_semantics(bytes: &[u8]) -> TestResult {
    let editor = pages::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let host = editor
        .table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?
        .ok_or("host read did not find the populated root comment")?;
    let host_replies = editor.table_cell_comment_replies(model, COMMENT_ROW, COMMENT_COLUMN)?;
    if host_replies.is_empty() {
        return Err("host source must contain a non-empty reply list".into());
    }

    let package = Package::from_bytes(bytes)?;
    let position = comment_position()?;
    let leaf = package
        .body_table_cell_comment(TABLE, position)?
        .ok_or("leaf read did not find the populated root comment")?;
    let leaf_replies = package.body_table_cell_comment_replies(TABLE, position)?;
    if leaf_replies.is_empty() {
        return Err("leaf source must contain a non-empty reply list".into());
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

fn comment_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &crate::harness::Fixture,
    input_path: &'static str,
    host: impl FnOnce(&mut pages::Draft, u64) -> TestResult + 'static,
    leaf: impl FnOnce(Package, At) -> TestResult<LeafRun> + 'static,
) -> TestResult {
    debug_assert_eq!(input.path, input_path);
    Case::new(name, Format::Pages, tier, evidence, input)
        .host(move |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            host(&mut editor, model)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(move |source| {
            let package = Package::from_bytes(source)?;
            leaf(package, comment_position()?)
        })
        .leaf_structural_metadata()
        .read("host table comment thread", host_comment_thread)
        .read("leaf table comment thread", leaf_comment_thread)
        .run()?;
    Ok(())
}

fn host_model(editor: &pages::Draft) -> TestResult<u64> {
    Ok(editor
        .tables()?
        .get(TABLE)
        .ok_or("body table missing")?
        .model_object_id)
}

fn host_cells(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = pages::Draft::from_bytes(bytes)?;
    let table = editor.table(host_model(&editor)?)?;
    let mut cells = table
        .iter_cells()
        .map(|((row, column), value)| format!("{row},{column}: {value:?}"))
        .collect::<Vec<_>>();
    cells.push(format!("count {}", table.cell_count()));
    Ok(cells)
}

fn leaf_cells(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:#?}",
        Package::from_bytes(bytes)?.body_table_cells(TABLE)?
    ))
}

fn host_value(value: &Value) -> TestResult<CellValue> {
    Ok(match value {
        Value::Text(text) => CellValue::Text(text.clone()),
        Value::Number(number) => CellValue::number(*number)?,
        Value::Boolean(flag) => CellValue::Boolean(*flag),
        Value::Empty => CellValue::Empty,
        _ => return Err("unmapped value".into()),
    })
}

fn values_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    updates: Vec<(u32, u32, Value)>,
) -> TestResult {
    let input = fixture(CATALOG)?;
    let updates = &updates;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            if let [(row, column, value)] = updates.as_slice() {
                editor.set_table_cell(
                    model,
                    *row as usize,
                    *column as usize,
                    host_value(value)?,
                )?;
            } else {
                let batch = updates
                    .iter()
                    .map(|(row, column, value)| {
                        Ok(TableCellUpdate::new(
                            *row as usize,
                            *column as usize,
                            host_value(value)?,
                        ))
                    })
                    .collect::<TestResult<Vec<_>>>()?;
                editor.set_table_cells(model, batch)?;
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_values(TABLE)?;
            for (row, column, value) in updates {
                edit.set(*row, *column, value.clone())?;
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_table_values(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host table cells", host_cells)
        .read("leaf body_table_cells", leaf_cells)
        .run()?;
    Ok(())
}

#[test]
fn pages_body_table_cell_parity() -> TestResult {
    values_case(
        "pages_body_table_cell_parity",
        Tier::A,
        Evidence::Queued,
        vec![(1, 3, Value::Text("Zürich".to_owned()))],
    )
}

#[test]
fn pages_body_table_cells_batch_parity() -> TestResult {
    values_case(
        "pages_body_table_cells_batch_parity",
        Tier::A,
        Evidence::Queued,
        vec![
            (2, 3, Value::Number(42.25)),
            (1, 2, Value::Boolean(false)),
            (2, 0, Value::Empty),
        ],
    )
}

#[test]
fn pages_body_table_cells_read_parity() -> TestResult {
    let input = fixture(CELLS)?;
    Case::new(
        "pages_body_table_cells_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Native("test-data/iwork/pages/body-table-cells-native-receipt.json"),
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host table cells", host_cells)
    .read("leaf body_table_cells", leaf_cells)
    .run()?;
    // The two reads agree value for value.
    let host = pages::Draft::from_bytes(&input.bytes)?;
    let host_table = host.table(host_model(&host)?)?;
    let leaf = Package::from_bytes(&input.bytes)?
        .body_table_cells(TABLE)?
        .into_table();
    for ((row, column), value) in host_table.iter_cells() {
        let at = litchi_iwa_common::table::coordinate::At::new(
            u32::try_from(row)?,
            u32::try_from(column)?,
        );
        assert_eq!(leaf.get(at), Some(value), "{row},{column}");
    }
    Ok(())
}

const COMMENTS: &str = "iwork/pages/body-table-comment-hidden-native-saved.pages";

/// The table list, merges and cell comments read the same through the host
/// (`tables`, `Table::{merges, get_comment, iter_comments, comment_count}`,
/// `table_cell_comment[_replies]`) and the leaf (`body_tables`,
/// `body_table_merges`, `body_table_cells`) on native fixtures.
#[test]
fn pages_body_table_reads_parity() -> TestResult {
    let input = fixture(CELLS)?;
    Case::new(
        "pages_body_table_reads_parity",
        Format::Pages,
        Tier::A,
        Evidence::Native("test-data/iwork/pages/body-table-cells-native-receipt.json"),
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("leaf body_table_cells", leaf_cells)
    .run()?;
    for path in [
        COMMENTS,
        CELLS,
        CATALOG,
        "iwork/pages/body-table-merges-native.pages",
    ] {
        let input = fixture(path)?;
        let host = pages::Draft::from_bytes(&input.bytes)?;
        let package = Package::from_bytes(&input.bytes)?;
        let catalog = package.body_tables()?;
        let tables = host.tables()?;
        assert_eq!(tables.len(), catalog.len(), "{path}");
        for (index, info) in tables.iter().enumerate() {
            let entry = catalog.get(index).ok_or("catalog entry")?;
            assert_eq!(entry.name(), info.name, "{path}");
            let host_table = host.table(info.model_object_id)?;
            let leaf_read = package.body_table_cells(index)?;
            assert_eq!(
                host_table.comment_count(),
                leaf_read.comment_count(),
                "{path}"
            );
            for ((row, column), comment) in host_table.iter_comments() {
                let at = litchi_iwa_common::table::coordinate::At::new(
                    u32::try_from(row)?,
                    u32::try_from(column)?,
                );
                assert_eq!(leaf_read.get_comment(at), Some(comment), "{path}");
                assert_eq!(host_table.get_comment(row, column), Some(comment));
                let native = host
                    .table_cell_comment(info.model_object_id, row, column)?
                    .ok_or("host cell comment")?;
                assert_eq!(native.comment.text, comment.text(), "{path}");
                assert_eq!(
                    host.table_cell_comment_replies(info.model_object_id, row, column)?
                        .len(),
                    comment.reply_count(),
                    "{path}"
                );
            }
            assert_eq!(
                host_table.merges(),
                package.body_table_merges(index)?.as_slice(),
                "{path}"
            );
        }
    }
    Ok(())
}

#[test]
fn pages_body_table_comment_replies_read_parity() -> TestResult {
    let input = fixture(COMMENTS)?;
    assert_populated_comment_semantics(&input.bytes)?;
    Case::new(
        "pages_body_table_comment_replies_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read(
        "host populated table comment thread",
        host_populated_comment_thread,
    )
    .read(
        "leaf populated table comment thread",
        leaf_populated_comment_thread,
    )
    .run()?;
    Ok(())
}

#[test]
fn pages_body_table_comment_set_parity() -> TestResult {
    let base = fixture(CELLS)?;
    let mut authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    authored.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G original")?;
    let input = crate::harness::Fixture {
        path: "generated:pages-table-comment-root-for-set",
        bytes: authored.to_bytes()?,
    };
    comment_case(
        "pages_body_table_comment_set_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        "generated:pages-table-comment-root-for-set",
        |editor, model| {
            editor.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G root")?;
            Ok(())
        },
        |package, position| {
            let commit = package.set_body_table_cell_comment(TABLE, position, "W1-G root")?;
            let restored = commit
                .package()
                .apply_body_table_cell_comment(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    )
}

#[test]
fn pages_body_table_comment_clear_parity() -> TestResult {
    let base = fixture(CELLS)?;
    let mut authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    authored.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G clear root")?;
    let input = crate::harness::Fixture {
        path: "generated:pages-table-comment-root-for-clear",
        bytes: authored.to_bytes()?,
    };
    comment_case(
        "pages_body_table_comment_clear_parity",
        Tier::B,
        Evidence::Queued,
        &input,
        "generated:pages-table-comment-root-for-clear",
        |editor, model| {
            editor.clear_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(())
        },
        |package, position| {
            let clear = package.clear_body_table_cell_comment(TABLE, position)?;
            let restored = clear
                .package()
                .apply_body_table_cell_comment(&clear.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(clear.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    )
}

/// Full Pages facade cost probe for the root-comment rewrite.  The source is
/// host-authored before ABBA, so the host and leaf both measure replacement of
/// the same existing root rather than construction of different inputs.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_cost_abba() -> TestResult {
    let base = fixture(CELLS)?;
    let mut authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    authored.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G original")?;
    let input = authored.to_bytes()?;
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &crate::harness::Fixture {
            path: "generated:pages-table-comment-root-cost",
            bytes: input,
        },
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G measured")?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package.set_body_table_cell_comment(TABLE, position, "W1-G measured")?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Full Pages facade cost probe for reading a populated root/reply thread.
/// The source is host-authored before ABBA and the semantic values are fed to
/// `black_box`, so this measures the read route rather than ZIP serialization
/// alone.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_read_cost_abba() -> TestResult {
    let input = fixture(COMMENTS)?;
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let comment = editor
                .table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?
                .ok_or("host read did not find the populated root comment")?;
            let replies = editor.table_cell_comment_replies(model, COMMENT_ROW, COMMENT_COLUMN)?;
            black_box((comment.comment.text.len(), replies.len()));
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let comment = package
                .body_table_cell_comment(TABLE, position)?
                .ok_or("leaf read did not find the populated root comment")?;
            let replies = package.body_table_cell_comment_replies(TABLE, position)?;
            black_box((comment.text().len(), replies.len()));
            Ok(leaf_bytes!(package))
        },
    )
}

/// Full Pages facade cost probe for creating a root comment on an existing
/// stored cell without a prior root comment.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_create_cost_abba() -> TestResult {
    let base = fixture(CELLS)?;
    let authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    if authored
        .table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?
        .is_some()
    {
        return Err("Pages create probe target already has a root comment".into());
    }
    let input = crate::harness::Fixture {
        path: "generated:pages-table-comment-root-create-cost",
        bytes: base.bytes,
    };
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G create")?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package.set_body_table_cell_comment(TABLE, position, "W1-G create")?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Full Pages facade cost probe for clearing a host-authored populated root.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn pages_body_table_comment_clear_cost_abba() -> TestResult {
    let base = fixture(CELLS)?;
    let mut authored = pages::Draft::from_bytes(&base.bytes)?;
    let model = host_model(&authored)?;
    authored.set_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN, "W1-G clear")?;
    let input = crate::harness::Fixture {
        path: "generated:pages-table-comment-root-clear-cost",
        bytes: authored.to_bytes()?,
    };
    let position = comment_position()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.clear_table_cell_comment(model, COMMENT_ROW, COMMENT_COLUMN)?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let clear = package.clear_body_table_cell_comment(TABLE, position)?;
            Ok(leaf_bytes!(clear.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_cell_value_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(CATALOG)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            editor.set_table_cell(model, 1, 3, CellValue::Text("Zürich".to_owned()))?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_values(TABLE)?;
            edit.set(1, 3, Value::Text("Zürich".to_owned()))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn table_cell_value_batch_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(CATALOG)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let updates = (0..1_000)
                .map(|index| {
                    TableCellUpdate::new(
                        1 + index % 4,
                        index / 4 % 4,
                        CellValue::Text(format!("v{index}")),
                    )
                })
                .collect::<Vec<_>>();
            // Later updates of one coordinate replace earlier ones; keep one per cell.
            let mut seen = std::collections::BTreeMap::new();
            for update in updates {
                seen.insert((update.row, update.column), update);
            }
            editor.set_table_cells(model, seen.into_values())?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_table_values(TABLE)?;
            let mut seen = std::collections::BTreeMap::new();
            for index in 0..1_000u32 {
                seen.insert((1 + index % 4, index / 4 % 4), format!("v{index}"));
            }
            for ((row, column), text) in seen {
                edit.set(row, column, Value::Text(text))?;
            }
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

/// `clear_table_cell` against the values family's `Value::Empty` on a cell
/// the host filled first (both drop the emptied row and its bare header).
#[test]
fn pages_body_table_cell_clear_parity() -> TestResult {
    clear_case(Tier::A, Evidence::Queued)
}

fn clear_case(tier: Tier, evidence: Evidence) -> TestResult {
    let source = fixture(CATALOG)?;
    let mut editor = pages::Draft::from_bytes(&source.bytes)?;
    let model = host_model(&editor)?;
    editor.set_table_cell(model, 1, 3, CellValue::Text("Zürich".to_owned()))?;
    let input = crate::harness::Fixture {
        path: "iwork/pages/body-table-catalog-native.pages + host D2 \"Zürich\"",
        bytes: editor.to_bytes()?,
    };
    Case::new(
        "pages_body_table_cell_clear_parity",
        Format::Pages,
        tier,
        evidence,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        let model = host_model(&editor)?;
        editor.clear_table_cell(model, 1, 3)?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_body_table_values(TABLE)?;
        edit.set(1, 3, Value::Empty)?;
        let commit = edit.commit()?;
        let restored = commit
            .package()
            .apply_body_table_values(&commit.patch().inverse())?;
        Ok(LeafRun::with_inverse(
            leaf_bytes!(commit.package()),
            leaf_bytes!(restored.package()),
        ))
    })
    .read("host table cells", host_cells)
    .read("leaf body_table_cells", leaf_cells)
    .run()?;
    Ok(())
}
