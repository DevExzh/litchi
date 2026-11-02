//! Pages body drawable comment parity (wave W4-P, record 0413): the host
//! `set_/clear_drawable_comment`, `*_drawable_comment_reply`,
//! `drawable_comment*` and `drawables` routes against
//! `litchi_pages::Package::edit_body_drawable_comment`.

use litchi_iwa::pages;
use litchi_iwa_common::comment::DrawableId;
use litchi_pages::{DrawableSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::pages_drawable_cases::{IMAGES, movie_fixture, shape_fixture};
use crate::pages_text_box_cases::{host_text_box, text_box_fixture};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    TextBox,
    Shape,
    Image,
    Movie,
}

const KINDS: [Kind; 4] = [Kind::TextBox, Kind::Shape, Kind::Image, Kind::Movie];

#[derive(Clone, Debug)]
enum Change {
    Set(&'static str),
    Clear,
    AddReply(&'static str),
    SetReply(usize, &'static str),
    RemoveReply(usize),
}

fn selector(kind: Kind) -> DrawableSelector {
    match kind {
        Kind::TextBox => DrawableSelector::text_box(0),
        Kind::Shape => DrawableSelector::shape(0),
        Kind::Image => DrawableSelector::image(0),
        Kind::Movie => DrawableSelector::movie(0),
    }
}

fn host_drawable(e: &pages::Draft, kind: Kind) -> TestResult<u64> {
    Ok(match kind {
        Kind::TextBox => host_text_box(e, 0)?,
        Kind::Shape => {
            e.body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id
        },
        Kind::Image => {
            e.body_images()?
                .first()
                .ok_or("no image")?
                .drawable_object_id
        },
        Kind::Movie => {
            e.body_movies()?
                .first()
                .ok_or("no movie")?
                .drawable_object_id
        },
    })
}

fn host_change(source: &[u8], kind: Kind, change: &Change) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let d = host_drawable(&e, kind)?;
    let id = DrawableId::try_from(d)?;
    let reply = |e: &pages::Draft, position: usize| {
        e.drawable_comment_replies(id)
            .map(|replies| replies.get(position).map(|reply| reply.storage_id))
    };
    match change {
        Change::Set(body) => e.set_drawable_comment(d, *body)?,
        Change::Clear => e.clear_drawable_comment(d)?,
        Change::AddReply(body) => {
            e.add_drawable_comment_reply(id, *body)?;
        },
        Change::SetReply(position, body) => {
            let reply = reply(&e, *position)?.ok_or("no reply")?;
            e.set_drawable_comment_reply(id, reply, *body)?;
        },
        Change::RemoveReply(position) => {
            let reply = reply(&e, *position)?.ok_or("no reply")?;
            e.remove_drawable_comment_reply(id, reply)?;
        },
    }
    Ok(e.to_bytes()?)
}

fn leaf_change(source: &[u8], kind: Kind, change: &Change) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_drawable_comment(selector(kind))?;
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
        .apply_body_drawable_comment(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_read(bytes: &[u8], kind: Kind) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = host_drawable(&e, kind)?;
    let id = DrawableId::try_from(d)?;
    let comment = e.drawable_comment(d)?.map(|comment| {
        (
            comment.comment.text,
            e.drawable_comment_replies(id)
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
        .body_drawable_comment(selector(kind))?
        .map(|comment| (comment.body().to_owned(), Some(comment.replies().to_vec())));
    Ok(format!("{comment:?}"))
}

fn run(name: &'static str, input: &Fixture, kind: Kind, change: &Change) -> TestResult {
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, input)
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

/// The fixture of `kind` after host `changes`.
fn with(kind: Kind, changes: &[Change]) -> TestResult<Fixture> {
    let mut bytes = match kind {
        Kind::TextBox => text_box_fixture()?.bytes,
        Kind::Shape => shape_fixture()?.bytes,
        Kind::Image => fixture(IMAGES)?.bytes,
        Kind::Movie => movie_fixture()?.bytes,
    };
    for change in changes {
        bytes = host_change(&bytes, kind, change)?;
    }
    Ok(Fixture {
        path: "Pages drawable fixture + host drawable comment",
        bytes,
    })
}

const THREAD: [Change; 3] = [
    Change::Set("Root"),
    Change::AddReply("First reply"),
    Change::AddReply("Second reply"),
];

#[test]
fn pages_drawable_comment_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        let name = "pages_drawable_comment_set_parity";
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
fn pages_drawable_comment_clear_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        let name = "pages_drawable_comment_clear_parity";
        run(name, &with(kind, &THREAD[..1])?, kind, &Change::Clear)?;
        run(name, &with(kind, &THREAD)?, kind, &Change::Clear)?;
    }
    Ok(())
}

#[test]
fn pages_drawable_comment_reply_add_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        let name = "pages_drawable_comment_reply_add_parity";
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
fn pages_drawable_comment_reply_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        for position in [0, 1] {
            run(
                "pages_drawable_comment_reply_set_parity",
                &with(kind, &THREAD)?,
                kind,
                &Change::SetReply(position, "Edited reply"),
            )?;
        }
    }
    Ok(())
}

#[test]
fn pages_drawable_comment_reply_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    for kind in KINDS {
        for position in [0, 1] {
            run(
                "pages_drawable_comment_reply_remove_parity",
                &with(kind, &THREAD)?,
                kind,
                &Change::RemoveReply(position),
            )?;
        }
    }
    Ok(())
}

#[test]
fn pages_drawable_comment_read_parity() -> TestResult {
    for kind in KINDS {
        for fixture in [
            with(kind, &[])?,
            with(kind, &THREAD[..1])?,
            with(kind, &THREAD)?,
        ] {
            assert_eq!(
                host_read(&fixture.bytes, kind)?,
                leaf_read(&fixture.bytes, kind)?,
                "{kind:?}"
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

/// The commentable drawables as (message type, has a comment), sorted.
fn host_drawables(bytes: &[u8]) -> TestResult<Vec<(u32, bool)>> {
    let e = pages::Draft::from_bytes(bytes)?;
    let mut found: Vec<(u32, bool)> = e
        .drawables()?
        .into_iter()
        .map(|item| (item.message_type, item.comment_id.is_some()))
        .collect();
    found.sort_unstable();
    Ok(found)
}

fn leaf_drawables(bytes: &[u8]) -> TestResult<Vec<(u32, bool)>> {
    let package = Package::from_bytes(bytes)?;
    let mut found = Vec::new();
    for (make, message_type) in [
        (
            DrawableSelector::text_box as fn(usize) -> DrawableSelector,
            2_011,
        ),
        (DrawableSelector::shape, 2_011),
        (DrawableSelector::image, 3_005),
        (DrawableSelector::movie, 3_007),
        (DrawableSelector::audio, 3_007),
    ] {
        let mut index = 0;
        while let Ok(comment) = package.body_drawable_comment(make(index)) {
            found.push((message_type, comment.is_some()));
            index += 1;
        }
    }
    found.sort_unstable();
    Ok(found)
}

#[test]
fn pages_drawable_list_parity() -> TestResult {
    for kind in KINDS {
        for fixture in [with(kind, &[])?, with(kind, &THREAD)?] {
            assert_eq!(
                host_drawables(&fixture.bytes)?,
                leaf_drawables(&fixture.bytes)?,
                "{kind:?}"
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0413)"]
fn drawable_comment_set_cost_abba() -> TestResult {
    let change = Change::Set("A drawable comment");
    crate::pages_cost_probes::abba(
        &with(Kind::TextBox, &[])?,
        |source| host_change(source, Kind::TextBox, &change),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_comment(selector(Kind::TextBox))?;
            edit.set("A drawable comment")?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
