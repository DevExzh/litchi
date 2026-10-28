//! Numbers table-cell comment mutation parity (wave W1-G, record 0379).
//!
//! The first shared slice covers only root replacement.  Direct reply
//! mutations remain an explicit host-only follow-up until their graph
//! transitions are lifted into the shared native engine.

#[cfg(feature = "parity-host-oracle")]
use std::collections::BTreeSet;
#[cfg(feature = "parity-host-oracle")]
use std::hint::black_box;

use litchi_iwa::numbers;
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_common::comment::AuthorId;
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_protos::tsd;
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_protos::tsk;
#[cfg(feature = "parity-host-oracle")]
use litchi_iwa_protos::tst;
use litchi_numbers::table::At;
use litchi_numbers::{Package, SheetSelector, TableSelector};
#[cfg(feature = "parity-host-oracle")]
use prost::Message as _;

#[cfg(feature = "parity-host-oracle")]
use crate::harness::fixture;
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
#[cfg(feature = "parity-host-oracle")]
use crate::known_fields::{Decoded, decode_as, top_level_fields};
use crate::leaf_bytes;
#[cfg(feature = "parity-host-oracle")]
use crate::text_glue::Workspace;

#[cfg(feature = "parity-host-oracle")]
const FIXTURE: &str = "iwork/numbers/comment-reader-native-created.numbers";
const ROW: usize = 1;
const COLUMN: usize = 1;
#[cfg(feature = "parity-host-oracle")]
const AUTHOR_MESSAGE_TYPE: u32 = 212;
#[cfg(feature = "parity-host-oracle")]
const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

#[cfg(feature = "parity-host-oracle")]
#[derive(Clone, Copy, Debug)]
enum Operation {
    Set,
    Create,
    Clear,
}

fn position() -> At {
    At::new(ROW as u32, COLUMN as u32)
}

fn host_model(editor: &numbers::Draft) -> TestResult<u64> {
    crate::numbers_cases::host_first_table_model_id(editor)
}

fn populated_comment_fixture() -> TestResult<Fixture> {
    // Keep the source-built marker alive while the host writes both records:
    // `Draft::set_cell_comment` therefore takes its legacy fallback, rather
    // than delegating this fixture's authoring to the focused leaf owner.
    let mut editor = numbers::Builder::new().table_dimensions(4, 4).build()?;
    let model = host_model(&editor)?;
    editor.set_cell_comment(model, ROW, COLUMN, "W1-G read root")?;
    editor.add_cell_comment_reply(model, ROW, COLUMN, "W1-G read reply")?;
    Ok(Fixture {
        path: "generated:numbers-table-comment-root-and-reply",
        bytes: editor.to_bytes()?,
    })
}

#[cfg(feature = "parity-host-oracle")]
fn native_fixture() -> TestResult<Fixture> {
    fixture(FIXTURE)
}

#[cfg(feature = "parity-host-oracle")]
fn rootless_fixture() -> TestResult<Fixture> {
    let native = native_fixture()?;
    let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
    editor.__parity_clear_table_cell_comment_legacy(TableSelector::index(0), ROW, COLUMN)?;
    let input = Fixture {
        path: "generated:numbers-table-comment-root-create-parity",
        bytes: editor.to_bytes()?,
    };
    if numbers::Draft::from_bytes(&input.bytes)?
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .is_some()
    {
        return Err("Numbers create target still has a root comment".into());
    }
    Ok(input)
}

fn host_populated_replies(bytes: &[u8]) -> TestResult<String> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let replies = editor.cell_comment_replies(model, ROW, COLUMN)?;
    if replies.is_empty() {
        return Err("host read did not find a populated comment reply".into());
    }
    let metadata = replies
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
    Ok(format!("replies={metadata:?}"))
}

fn leaf_populated_thread(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let root = package
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
        .ok_or("leaf read did not find a populated root comment")?;
    if root.text().is_empty() || root.timestamp().is_none() || root.author().is_none() {
        return Err("leaf root comment is missing text/date/author metadata".into());
    }
    let replies = package.table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        position(),
    )?;
    if replies.is_empty() {
        return Err("leaf read did not find a populated comment reply".into());
    }
    let metadata = replies
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
    // The host's focused Numbers facade exposes reply metadata here; the
    // root assertion above still proves that the same host-authored source
    // contains a populated root, but it is not projected into this legacy
    // host read string.
    Ok(format!("replies={metadata:?}"))
}

#[cfg(feature = "parity-host-oracle")]
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

#[cfg(feature = "parity-host-oracle")]
fn decode_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let decoded = decode_as::<tsd::CommentStorageArchive>(data)?;
    Ok(decoded)
}

#[cfg(feature = "parity-host-oracle")]
fn masked_new_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tsd::CommentStorageArchive::decode(data)?;
    if message.creation_date.is_some() {
        message.creation_date = Some(litchi_iwa_protos::tsp::Date { seconds: 0.0 });
    }
    if message.storage_uuid.is_some() {
        message.storage_uuid = Some(litchi_iwa_protos::tsp::Uuid { lower: 1, upper: 1 });
    }
    let mut decoded = decode_as::<tsd::CommentStorageArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

#[cfg(feature = "parity-host-oracle")]
fn masked_clear_table_data_list(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tst::TableDataList::decode(data)?;
    message.next_list_id = 0;
    let mut decoded = decode_as::<tst::TableDataList>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

#[cfg(feature = "parity-host-oracle")]
fn existing_comment_metadata(
    bytes: &[u8],
) -> TestResult<(Vec<u8>, Option<f64>, Option<(u64, u64)>)> {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let comment = editor
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .ok_or("selected native source has no existing root comment")?;
    let located = Workspace::open(bytes)?.message(comment.storage_id.get(), 3056)?;
    let decoded = tsd::CommentStorageArchive::decode(located.payload.as_slice())?;
    let view = litchi_iwa_common::wire::WireView::parse(
        located.payload.as_slice(),
        litchi_iwa_common::WireLimits::DEFAULT,
    )?;
    let mut non_text_fields = Vec::new();
    for field in view.fields().filter(|field| field.number() != 1) {
        non_text_fields.extend_from_slice(field.raw());
    }
    Ok((
        non_text_fields,
        decoded.creation_date.map(|date| date.seconds),
        decoded.storage_uuid.map(|uuid| (uuid.lower, uuid.upper)),
    ))
}

/// Compare host and leaf semantic values on the same host-authored bytes.
/// The `Case::read` callbacks otherwise consume separate route outputs and
/// cannot establish that the readers project an untouched source identically.
#[cfg(feature = "parity-host-oracle")]
fn assert_populated_comment_semantics(bytes: &[u8]) -> TestResult {
    let editor = numbers::Draft::from_bytes(bytes)?;
    let model = host_model(&editor)?;
    let host = editor
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .ok_or("host read did not find the populated root comment")?;
    let host_replies = editor.cell_comment_replies(model, ROW, COLUMN)?;
    if host_replies.is_empty() {
        return Err("host source must contain a non-empty reply list".into());
    }

    let package = Package::from_bytes(bytes)?;
    let leaf = package
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
        .ok_or("leaf read did not find the populated root comment")?;
    let leaf_replies = package.table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        position(),
    )?;
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

#[cfg(feature = "parity-host-oracle")]
fn preview_names(bytes: &[u8]) -> TestResult<BTreeSet<String>> {
    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(bytes, Default::default())?;
    Ok(catalog
        .iter()
        .filter(|entry| ROOT_PREVIEWS.contains(&entry.name()))
        .map(|entry| entry.name().to_owned())
        .collect())
}

#[cfg(feature = "parity-host-oracle")]
fn host_route(source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    match operation {
        Operation::Set | Operation::Create => {
            // This is the deliberately independent legacy oracle.  The
            // helper resolves the selector and calls the old package writer
            // directly; it never constructs a focused `litchi_numbers`
            // package.
            editor.__parity_set_table_cell_comment_legacy(
                TableSelector::index(0),
                ROW,
                COLUMN,
                "W1-G revised root",
            )?;
        },
        Operation::Clear => {
            editor.__parity_clear_table_cell_comment_legacy(
                TableSelector::index(0),
                ROW,
                COLUMN,
            )?;
        },
    }
    Ok(editor.to_bytes()?)
}

#[cfg(feature = "parity-host-oracle")]
fn leaf_forward(source: &[u8], operation: Operation) -> TestResult<Vec<u8>> {
    let package = Package::from_bytes(source)?;
    let commit = match operation {
        Operation::Set | Operation::Create => package.set_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
            "W1-G revised root",
        )?,
        Operation::Clear => package.clear_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
        )?,
    };
    Ok(leaf_bytes!(commit.package()))
}

#[cfg(feature = "parity-host-oracle")]
fn leaf_route(source: &[u8], operation: Operation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let commit = match operation {
        Operation::Set | Operation::Create => package.set_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
            "W1-G revised root",
        )?,
        Operation::Clear => package.clear_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            position(),
        )?,
    };
    let restored = commit
        .package()
        .apply_table_cell_comment(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

#[cfg(feature = "parity-host-oracle")]
fn run(
    name: &'static str,
    operation: Operation,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
) -> TestResult {
    let mut case = Case::new(name, Format::Numbers, tier, evidence, input)
        .host(move |source| host_route(source, operation))
        .leaf(move |source| leaf_route(source, operation))
        .known_fields(|known| {
            known
                .with(3_056, "TSD.CommentStorageArchive", decode_comment_storage)
                .with(
                    212,
                    "TSK.AnnotationAuthorArchive",
                    decode_as::<litchi_iwa_protos::tsk::AnnotationAuthorArchive>,
                )
                .with(
                    213,
                    "TSK.AnnotationAuthorStorageArchive",
                    decode_as::<litchi_iwa_protos::tsk::AnnotationAuthorStorageArchive>,
                )
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata();
    if matches!(operation, Operation::Create) {
        case = case.known_fields(|known| {
            known.with(
                3_056,
                "TSD.CommentStorageArchive (new-object date and storage UUID masked)",
                masked_new_comment_storage,
            )
        });
    } else if matches!(operation, Operation::Clear) {
        case = case.known_fields(|known| {
            known.with(
                6_005,
                "TST.TableDataList (clear watermark masked)",
                masked_clear_table_data_list,
            )
        });
    }
    case.run()
        .map_err(|error| format!("{operation:?}: {error}"))?;
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_replies_read_parity() -> TestResult {
    let input = populated_comment_fixture()?;
    assert_populated_comment_semantics(&input.bytes)?;
    Case::new(
        "numbers_table_cell_comment_replies_read_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(numbers::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host table-cell replies", host_populated_replies)
    .read("leaf table-cell comment thread", leaf_populated_thread)
    .run()?;
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_set_parity() -> TestResult {
    let input = native_fixture()?;
    let before = existing_comment_metadata(&input.bytes)?;
    if before.1.is_none() || before.2.is_none() {
        return Err("native root comment is missing source-owned date/UUID metadata".into());
    }
    run(
        "numbers_table_cell_comment_set_parity",
        Operation::Set,
        Tier::B,
        Evidence::Queued,
        &input,
    )?;
    let host = host_route(&input.bytes, Operation::Set)?;
    let leaf = leaf_route(&input.bytes, Operation::Set)?.output;
    for (route, output) in [("host", host), ("leaf", leaf)] {
        let after = existing_comment_metadata(&output)?;
        assert_eq!(
            after.0, before.0,
            "{route} rewrote existing comment metadata"
        );
        assert_eq!(after.1, before.1, "{route} rewrote existing comment date");
        assert_eq!(after.2, before.2, "{route} rewrote existing comment UUID");
    }
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_create_parity() -> TestResult {
    let input = rootless_fixture()?;
    run(
        "numbers_table_cell_comment_create_parity",
        Operation::Create,
        Tier::B,
        Evidence::Queued,
        &input,
    )?;
    let host = host_route(&input.bytes, Operation::Create)?;
    let leaf = leaf_route(&input.bytes, Operation::Create)?.output;
    let host_root = numbers::Draft::from_bytes(&host)?
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .ok_or("legacy host create did not persist the root comment")?;
    assert_eq!(host_root.comment.text, "W1-G revised root");
    let leaf_root = Package::from_bytes(&leaf)?
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
        .ok_or("leaf create did not persist the root comment")?;
    assert_eq!(leaf_root.text(), "W1-G revised root");
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_clear_parity() -> TestResult {
    let input = native_fixture()?;
    let editor = numbers::Draft::from_bytes(&input.bytes)?;
    let model = host_model(&editor)?;
    if editor
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .is_none()
    {
        return Err("Numbers clear target has no root comment".into());
    }
    if !editor.cell_comment_replies(model, ROW, COLUMN)?.is_empty() {
        return Err("Numbers clear target unexpectedly has replies".into());
    }
    run(
        "numbers_table_cell_comment_clear_parity",
        Operation::Clear,
        Tier::B,
        Evidence::Queued,
        &input,
    )?;
    let host = host_route(&input.bytes, Operation::Clear)?;
    let leaf = leaf_route(&input.bytes, Operation::Clear)?.output;
    assert!(
        numbers::Draft::from_bytes(&host)?
            .__parity_read_table_cell_comment(TableSelector::index(0), position())?
            .is_none(),
        "legacy host clear left a root comment"
    );
    assert!(
        Package::from_bytes(&leaf)?
            .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
            .is_none(),
        "leaf clear left a root comment"
    );
    Ok(())
}

#[cfg(feature = "parity-host-oracle")]
#[test]
fn numbers_table_cell_comment_host_oracle_control() -> TestResult {
    let input = native_fixture()?;
    let output = host_route(&input.bytes, Operation::Set)?;
    assert_ne!(output, input.bytes, "legacy oracle must publish a change");
    let comment = Package::from_bytes(&output)?
        .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
        .ok_or("legacy host oracle did not persist the comment")?;
    assert_eq!(comment.text(), "W1-G revised root");

    // The legacy helper above calls `set_cell_comment_in_package` directly;
    // it does not construct or invoke the focused package owner.  Preserve
    // the three preview members as the discriminating legacy behavior: the
    // focused leaf route intentionally invalidates them after its write.
    let input_previews = preview_names(&input.bytes)?;
    assert_eq!(input_previews.len(), ROOT_PREVIEWS.len());
    assert_eq!(preview_names(&output)?, input_previews);
    let leaf_output = leaf_route(&input.bytes, Operation::Set)?.output;
    assert!(
        preview_names(&leaf_output)?.len() < input_previews.len(),
        "focused route should invalidate at least one root preview"
    );
    Ok(())
}

/// Full Numbers facade cost probe for the focused root replacement.  This
/// uses the same native source as the parity case and the independent legacy
/// host oracle; run it with `--features parity-host-oracle` under callgrind.
#[cfg(feature = "parity-host-oracle")]
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_cost_abba() -> TestResult {
    let input = native_fixture()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, Operation::Set),
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package.set_table_cell_comment(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
                "W1-G revised root",
            )?;
            Ok(leaf_bytes!(commit.package()))
        },
    )
}

/// Full Numbers facade cost probe for reading a populated root/reply thread.
/// This uses the feature-gated legacy host reader as the differential oracle;
/// it is not a supported host API and is not native application evidence.
#[cfg(feature = "parity-host-oracle")]
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_read_cost_abba() -> TestResult {
    let input = populated_comment_fixture()?;
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let editor = numbers::Draft::from_bytes(source)?;
            let model = host_model(&editor)?;
            let root = editor
                .__parity_read_table_cell_comment(TableSelector::index(0), position())?
                .ok_or("host read did not find the populated root comment")?;
            let replies = editor.cell_comment_replies(model, ROW, COLUMN)?;
            black_box((root.comment.text.len(), replies.len()));
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let root = package
                .table_cell_comment(SheetSelector::index(0), TableSelector::index(0), position())?
                .ok_or("leaf read did not find the populated root comment")?;
            let replies = package.table_cell_comment_replies(
                SheetSelector::index(0),
                TableSelector::index(0),
                position(),
            )?;
            black_box((root.text().len(), replies.len()));
            Ok(leaf_bytes!(package))
        },
    )
}

/// Full Numbers facade cost probe for creating a root comment.  The input is
/// host-authored and cleared before ABBA, and both mutation routes use the
/// feature-gated legacy writer/reader oracle boundary described above.
#[cfg(feature = "parity-host-oracle")]
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_create_cost_abba() -> TestResult {
    let native = native_fixture()?;
    let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
    editor.__parity_clear_table_cell_comment_legacy(TableSelector::index(0), ROW, COLUMN)?;
    let input = Fixture {
        path: "generated:numbers-table-comment-root-create-cost",
        bytes: editor.to_bytes()?,
    };
    if numbers::Draft::from_bytes(&input.bytes)?
        .__parity_read_table_cell_comment(TableSelector::index(0), position())?
        .is_some()
    {
        return Err("Numbers create probe target still has a root comment".into());
    }
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, Operation::Create),
        |source| leaf_forward(source, Operation::Create),
    )
}

/// Full Numbers facade cost probe for clearing a host-authored root comment.
/// The root has no replies, matching the supported root-only clear profile.
#[cfg(feature = "parity-host-oracle")]
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_table_cell_comment_clear_cost_abba() -> TestResult {
    let native = native_fixture()?;
    let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
    editor.__parity_set_table_cell_comment_legacy(
        TableSelector::index(0),
        ROW,
        COLUMN,
        "W1-G clear",
    )?;
    let input = Fixture {
        path: "generated:numbers-table-comment-root-clear-cost",
        bytes: editor.to_bytes()?,
    };
    crate::pages_cost_probes::abba(
        &input,
        |source| host_route(source, Operation::Clear),
        |source| leaf_forward(source, Operation::Clear),
    )
}
