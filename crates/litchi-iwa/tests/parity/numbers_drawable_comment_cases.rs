//! Numbers sheet drawable comment parity (wave W2-A, record 0378): the host
//! `set_/clear_sheet_drawable_comment`, `*_sheet_drawable_comment_reply` and
//! `sheet_drawable_comment*` families against
//! `litchi_numbers::Package::edit_sheet_drawable_comment`.

use litchi_iwa::numbers;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::{DrawableSelector, Kind};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::numbers_drawable_cases::{authored_fixture, chart_fixture, host_drawable};

/// The commentable kinds of the authored and chart fixtures.
const KINDS: [Kind; 4] = [Kind::TextBox, Kind::Shape, Kind::Movie, Kind::Chart];

#[derive(Clone, Debug)]
enum Change {
    Set(&'static str),
    Clear,
    AddReply(&'static str),
    SetReply(usize, &'static str),
    RemoveReply(usize),
}

fn selector(kind: Kind) -> DrawableSelector {
    DrawableSelector::of(kind, 0)
}

fn host_change(source: &[u8], kind: Kind, change: &Change) -> TestResult<Vec<u8>> {
    let mut e = numbers::Draft::from_bytes(source)?;
    let (s, d) = host_drawable(&e, kind, 0)?;
    let reply = |e: &numbers::Draft, position: usize| -> TestResult<u64> {
        Ok(e.sheet_drawable_comment_replies(s, d)?
            .get(position)
            .ok_or("no reply")?
            .storage_id
            .get())
    };
    match change {
        Change::Set(body) => e.set_sheet_drawable_comment(s, d, *body)?,
        Change::Clear => e.clear_sheet_drawable_comment(s, d)?,
        Change::AddReply(body) => {
            e.add_sheet_drawable_comment_reply(s, d, *body)?;
        },
        Change::SetReply(position, body) => {
            let id = reply(&e, *position)?;
            e.set_sheet_drawable_comment_reply(s, d, id, *body)?;
        },
        Change::RemoveReply(position) => {
            let id = reply(&e, *position)?;
            e.remove_sheet_drawable_comment_reply(s, d, id)?;
        },
    }
    Ok(e.to_bytes()?)
}

fn leaf_change(source: &[u8], kind: Kind, change: &Change) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_sheet_drawable_comment(0, selector(kind))?;
    match change {
        Change::Set(body) => edit.set(body)?,
        Change::Clear => edit.clear()?,
        Change::AddReply(body) => edit.add_reply(body)?,
        Change::SetReply(position, body) => edit.set_reply(*position, body)?,
        Change::RemoveReply(position) => edit.remove_reply(*position)?,
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_sheet_drawable_comment(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_read(bytes: &[u8], kind: Kind) -> TestResult<String> {
    let e = numbers::Draft::from_bytes(bytes)?;
    let (s, d) = host_drawable(&e, kind, 0)?;
    let comment = e.sheet_drawable_comment(s, d)?.map(|comment| {
        (
            comment.comment.text,
            e.sheet_drawable_comment_replies(s, d)
                .map(|replies| {
                    replies
                        .into_iter()
                        .map(|reply| reply.comment.text)
                        .collect::<Vec<_>>()
                })
                .ok(),
        )
    });
    Ok(format!("{comment:?}"))
}

fn leaf_read(bytes: &[u8], kind: Kind) -> TestResult<String> {
    let comment = Package::from_bytes(bytes)?
        .sheet_drawable_comment(0, selector(kind))?
        .map(|comment| (comment.body().to_owned(), Some(comment.replies().to_vec())));
    Ok(format!("{comment:?}"))
}

fn run(name: &'static str, input: &Fixture, kind: Kind, change: &Change) -> TestResult {
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, input)
        .host(|source| host_change(source, kind, change))
        .leaf(|source| leaf_change(source, kind, change))
        .read("host comment", move |bytes| host_read(bytes, kind))
        .read("leaf comment", move |bytes| leaf_read(bytes, kind))
        .known_fields(|known| {
            crate::text_comment_cases::masked(known)
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
        .leaf_structural_metadata()
        .run()
        .map_err(|error| format!("{kind:?} {change:?}: {error}"))?;
    Ok(())
}

/// The authored fixture (the chart fixture for a chart) after host
/// `changes` on `kind`.
fn with(kind: Kind, changes: &[Change]) -> TestResult<Fixture> {
    let mut bytes = if kind == Kind::Chart {
        chart_fixture()?.bytes
    } else {
        authored_fixture()?.bytes
    };
    for change in changes {
        bytes = host_change(&bytes, kind, change)?;
    }
    Ok(Fixture {
        path: "authored + host drawable comment",
        bytes,
    })
}

const THREAD: [Change; 3] = [
    Change::Set("Root"),
    Change::AddReply("First reply"),
    Change::AddReply("Second reply"),
];

#[test]
fn numbers_sheet_drawable_comment_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        let name = "numbers_sheet_drawable_comment_set_parity";
        run(
            name,
            &with(kind, &[])?,
            kind,
            &Change::Set("A drawable comment"),
        )?;
        run(name, &with(kind, &THREAD)?, kind, &Change::Set("Rewritten"))?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_drawable_comment_clear_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        let name = "numbers_sheet_drawable_comment_clear_parity";
        run(name, &with(kind, &THREAD[..1])?, kind, &Change::Clear)?;
        run(name, &with(kind, &THREAD)?, kind, &Change::Clear)?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_drawable_comment_reply_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        let name = "numbers_sheet_drawable_comment_reply_add_parity";
        run(
            name,
            &with(kind, &THREAD[..1])?,
            kind,
            &Change::AddReply("Reply"),
        )?;
        run(
            name,
            &with(kind, &THREAD)?,
            kind,
            &Change::AddReply("Third"),
        )?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_drawable_comment_reply_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        for position in [0, 1] {
            run(
                "numbers_sheet_drawable_comment_reply_set_parity",
                &with(kind, &THREAD)?,
                kind,
                &Change::SetReply(position, "Edited reply"),
            )?;
        }
    }
    Ok(())
}

#[test]
fn numbers_sheet_drawable_comment_reply_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        for position in [0, 1] {
            run(
                "numbers_sheet_drawable_comment_reply_remove_parity",
                &with(kind, &THREAD)?,
                kind,
                &Change::RemoveReply(position),
            )?;
        }
    }
    Ok(())
}

#[test]
fn numbers_sheet_drawable_comment_read_parity() -> TestResult {
    for kind in KINDS {
        for fixture in [
            with(kind, &[])?,
            with(kind, &THREAD[..1])?,
            with(kind, &THREAD)?,
        ] {
            assert_eq!(
                host_read(&fixture.bytes, kind)?,
                leaf_read(&fixture.bytes, kind)?
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

/// The commentable drawables of the first sheet as (message type, has a
/// comment), sorted.
fn host_drawables(bytes: &[u8]) -> TestResult<Vec<(u32, bool)>> {
    let e = numbers::Draft::from_bytes(bytes)?;
    let sheet = crate::numbers_drawable_cases::first_sheet(&e)?;
    let mut found: Vec<(u32, bool)> = e
        .sheet_drawables(sheet)?
        .into_iter()
        .map(|item| (item.message_type, item.comment_id.is_some()))
        .collect();
    found.sort_unstable();
    Ok(found)
}

fn leaf_drawables(bytes: &[u8]) -> TestResult<Vec<(u32, bool)>> {
    let package = Package::from_bytes(bytes)?;
    let mut found = Vec::new();
    for item in package.sheet_drawable_list(0)? {
        let message_type = match item.kind() {
            Kind::TextBox | Kind::Shape => 2_011,
            Kind::Image => 3_005,
            Kind::Movie | Kind::Audio => 3_007,
            Kind::Chart => 5_021,
            _ => continue,
        };
        let comment = package.sheet_drawable_comment(0, item.selector())?;
        found.push((message_type, comment.is_some()));
    }
    found.sort_unstable();
    Ok(found)
}

#[test]
fn numbers_sheet_drawable_list_parity() -> TestResult {
    for kind in KINDS {
        for fixture in [with(kind, &[])?, with(kind, &THREAD)?] {
            assert_eq!(
                host_drawables(&fixture.bytes)?,
                leaf_drawables(&fixture.bytes)?
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
