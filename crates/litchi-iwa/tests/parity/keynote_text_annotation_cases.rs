//! Keynote text-box annotation parity (wave W1-C, record 0372): hyperlinks
//! of the host `*_slide_text_box_hyperlink*` families against
//! `slide::drawable::Edit::{add,update,remove}_hyperlink` and
//! `Package::slide_drawable_hyperlinks` on the shared W1-D smart-field
//! engine. New fields get a random text-attribute UUID on the host and a
//! derived one in the leaf, so that one string is masked.

use litchi_iwa::keynote;
use litchi_iwa_protos::tswp;
use litchi_iwa_text::hyperlink::TextHyperlinkTarget;
use litchi_iwa_text::position::{TextPosition, TextRange};
use litchi_iwa_text_wire::smart_field::{self, TextAttributeUuid};
use litchi_keynote::slide::drawable::Kind;
use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector};
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_drawable_cases::{host_id, leaf_position};
use crate::known_fields::{Decoded, KnownFields, decode_as, top_level_fields};
use crate::leaf_bytes;

const SLIDE: usize = 0;
const TEXT_BOX: &str = "iwork/keynote/basic.key";

fn masked_hyperlink(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tswp::HyperlinkFieldArchive::decode(data)?;
    if let Some(field) = message.super_.as_mut()
        && field
            .text_attribute_uuid_string
            .as_deref()
            .is_some_and(|value| TextAttributeUuid::new(value).is_ok())
    {
        field.text_attribute_uuid_string = Some("<uuid>".to_owned());
    }
    let mut decoded = decode_as::<tswp::HyperlinkFieldArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn masked(known: KnownFields) -> KnownFields {
    known.with(
        smart_field::HYPERLINK_MESSAGE_TYPE,
        "TSWP.HyperlinkFieldArchive (text-attribute UUID masked)",
        masked_hyperlink,
    )
}

fn range(start: u32, end: u32) -> TestResult<TextRange> {
    Ok(TextRange::new(
        TextPosition::from_utf16_code_units(start),
        TextPosition::from_utf16_code_units(end),
    )?)
}

#[derive(Clone, Debug)]
enum Operation {
    Add(TextRange, TextHyperlinkTarget),
    Update(usize, TextRange, TextHyperlinkTarget),
    Remove(usize),
}

fn host_route(source: &[u8], operation: &Operation) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let id = host_id(&editor, source, leaf_position(source, Kind::TextBox)?)?;
    match operation.clone() {
        Operation::Add(range, target) => {
            editor.add_slide_text_box_hyperlink(SLIDE, id, range, target)?;
        },
        Operation::Update(index, range, target) => {
            let link = editor.slide_text_box_hyperlinks(SLIDE, id)?[index].id;
            editor.update_slide_text_box_hyperlink(SLIDE, id, link, range, target)?;
        },
        Operation::Remove(index) => {
            let link = editor.slide_text_box_hyperlinks(SLIDE, id)?[index].id;
            editor.remove_slide_text_box_hyperlink(SLIDE, id, link)?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_route(source: &[u8], operation: &Operation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let (slide, drawable) = (
        SlideSelector::index(SLIDE),
        DrawableSelector::index(leaf_position(source, Kind::TextBox)?),
    );
    let mut edit = package.edit_slide_drawables();
    match operation.clone() {
        Operation::Add(range, target) => edit.add_hyperlink(slide, drawable, range, target)?,
        Operation::Update(index, range, target) => {
            edit.update_hyperlink(slide, drawable, Position::new(index), range, target)?
        },
        Operation::Remove(index) => edit.remove_hyperlink(slide, drawable, Position::new(index))?,
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_drawables(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_reads(bytes: &[u8]) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let id = host_id(&editor, bytes, leaf_position(bytes, Kind::TextBox)?)?;
    Ok(format!(
        "{:?}",
        editor
            .slide_text_box_hyperlinks(SLIDE, id)?
            .into_iter()
            .map(|link| (link.range, link.target))
            .collect::<Vec<_>>()
    ))
}

fn leaf_reads(bytes: &[u8]) -> TestResult<String> {
    let position = leaf_position(bytes, Kind::TextBox)?;
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?
            .slide_drawable_hyperlinks(SLIDE, DrawableSelector::index(position))?
    ))
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    operation: &Operation,
) -> TestResult {
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| host_route(source, operation))
        .leaf(|source| leaf_route(source, operation))
        .read("host hyperlinks", host_reads)
        .read("leaf hyperlinks", leaf_reads)
        .known_fields(masked)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()
        .map_err(|error| format!("{operation:?}: {error}"))?;
    Ok(())
}

fn linked() -> TestResult<Fixture> {
    let input = fixture(TEXT_BOX)?;
    Ok(Fixture {
        path: input.path,
        bytes: host_route(
            &input.bytes,
            &Operation::Add(
                range(0, 4)?,
                TextHyperlinkTarget::new("https://example.com/litchi")?,
            ),
        )?,
    })
}

#[test]
fn keynote_slide_text_box_hyperlink_add_parity() -> TestResult {
    run(
        "keynote_slide_text_box_hyperlink_add_parity",
        Tier::B,
        Evidence::Queued,
        &fixture(TEXT_BOX)?,
        &Operation::Add(
            range(0, 4)?,
            TextHyperlinkTarget::new("https://example.com/litchi")?,
        ),
    )
}

#[test]
fn keynote_slide_text_box_hyperlink_update_parity() -> TestResult {
    run(
        "keynote_slide_text_box_hyperlink_update_parity",
        Tier::B,
        Evidence::Queued,
        &linked()?,
        &Operation::Update(
            0,
            range(1, 3)?,
            TextHyperlinkTarget::new("mailto:a@example.com")?,
        ),
    )
}

#[test]
fn keynote_slide_text_box_hyperlink_remove_parity() -> TestResult {
    run(
        "keynote_slide_text_box_hyperlink_remove_parity",
        Tier::B,
        Evidence::Queued,
        &linked()?,
        &Operation::Remove(0),
    )
}

#[test]
fn keynote_slide_text_box_hyperlink_read_parity() -> TestResult {
    for input in [fixture(TEXT_BOX)?, linked()?] {
        assert_eq!(host_reads(&input.bytes)?, leaf_reads(&input.bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

fn masked_highlight(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tswp::HighlightArchive::decode(data)?;
    if message
        .text_attribute_uuid_string
        .as_deref()
        .is_some_and(|uuid| TextAttributeUuid::new(uuid).is_ok())
    {
        message.text_attribute_uuid_string = Some("<uuid>".to_owned());
    }
    let mut decoded = decode_as::<tswp::HighlightArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn masked_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    use litchi_iwa_protos::{tsd, tsp};
    let mut message = tsd::CommentStorageArchive::decode(data)?;
    if message.creation_date.is_some() {
        message.creation_date = Some(tsp::Date { seconds: 0.0 });
    }
    if message.storage_uuid.is_some() {
        message.storage_uuid = Some(tsp::Uuid { lower: 1, upper: 1 });
    }
    let mut decoded = decode_as::<tsd::CommentStorageArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn masked_comments(known: KnownFields) -> KnownFields {
    known
        .with(
            litchi_iwa_text_wire::annotation::HIGHLIGHT_MESSAGE_TYPE,
            "TSWP.HighlightArchive (text-attribute UUID masked)",
            masked_highlight,
        )
        .with(
            litchi_iwa_text_wire::annotation::COMMENT_STORAGE_MESSAGE_TYPE,
            "TSD.CommentStorageArchive (creation date and storage UUID masked)",
            masked_comment_storage,
        )
}

#[derive(Clone, Debug)]
enum Annotation {
    AddHighlight(TextRange),
    UpdateHighlight(usize, TextRange),
    RemoveHighlight(usize),
    AddComment(TextRange, &'static str),
    UpdateComment(usize, TextRange, &'static str),
    RemoveComment(usize),
    AddReply(usize, &'static str),
    UpdateReply(usize, usize, &'static str),
    RemoveReply(usize, usize),
}

fn when() -> TestResult<litchi_iwa_text::date_time::Instant> {
    Ok(litchi_iwa_text::date_time::Instant::from_reference_date_seconds(800_000_000.0)?)
}

fn host_annotate(source: &[u8], operation: &Annotation) -> TestResult<Vec<u8>> {
    use litchi_iwa_text::comment::{TextCommentBody, TextCommentReplyBody};
    let mut editor = keynote::Draft::from_bytes(source)?;
    let id = host_id(&editor, source, leaf_position(source, Kind::TextBox)?)?;
    let comment = |editor: &keynote::Draft, index: usize| -> TestResult<_> {
        Ok(editor.slide_text_box_comments(SLIDE, id)?[index].id())
    };
    match operation.clone() {
        Annotation::AddHighlight(range) => {
            editor.add_slide_text_box_highlight(SLIDE, id, range)?;
        },
        Annotation::UpdateHighlight(index, range) => {
            let highlight = editor.slide_text_box_highlights(SLIDE, id)?[index].id;
            editor.update_slide_text_box_highlight(SLIDE, id, highlight, range)?;
        },
        Annotation::RemoveHighlight(index) => {
            let highlight = editor.slide_text_box_highlights(SLIDE, id)?[index].id;
            editor.remove_slide_text_box_highlight(SLIDE, id, highlight)?;
        },
        Annotation::AddComment(range, body) => {
            editor.add_slide_text_box_comment(SLIDE, id, range, TextCommentBody::new(body)?)?;
        },
        Annotation::UpdateComment(index, range, body) => {
            let target = comment(&editor, index)?;
            editor.update_slide_text_box_comment(
                SLIDE,
                id,
                target,
                range,
                TextCommentBody::new(body)?,
            )?;
        },
        Annotation::RemoveComment(index) => {
            let target = comment(&editor, index)?;
            editor.remove_slide_text_box_comment(SLIDE, id, target)?;
        },
        Annotation::AddReply(index, body) => {
            let target = comment(&editor, index)?;
            editor.add_slide_text_box_comment_reply(
                SLIDE,
                id,
                target,
                TextCommentReplyBody::new(body)?,
            )?;
        },
        Annotation::UpdateReply(index, reply, body) => {
            let target = comment(&editor, index)?;
            let reply = editor.slide_text_box_comment_replies(SLIDE, id, target)?[reply].id();
            editor.update_slide_text_box_comment_reply(
                SLIDE,
                id,
                target,
                reply,
                TextCommentReplyBody::new(body)?,
            )?;
        },
        Annotation::RemoveReply(index, reply) => {
            let target = comment(&editor, index)?;
            let reply = editor.slide_text_box_comment_replies(SLIDE, id, target)?[reply].id();
            editor.remove_slide_text_box_comment_reply(SLIDE, id, target, reply)?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_annotate(source: &[u8], operation: &Annotation) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let (slide, drawable) = (
        SlideSelector::index(SLIDE),
        DrawableSelector::index(leaf_position(source, Kind::TextBox)?),
    );
    let mut edit = package.edit_slide_drawables();
    match operation.clone() {
        Annotation::AddHighlight(range) => edit.add_highlight(slide, drawable, range, when()?)?,
        Annotation::UpdateHighlight(index, range) => {
            edit.update_highlight(slide, drawable, Position::new(index), range)?
        },
        Annotation::RemoveHighlight(index) => {
            edit.remove_highlight(slide, drawable, Position::new(index))?
        },
        Annotation::AddComment(range, body) => {
            edit.add_comment(slide, drawable, range, body, when()?)?
        },
        Annotation::UpdateComment(index, range, body) => {
            edit.update_comment(slide, drawable, Position::new(index), range, body)?
        },
        Annotation::RemoveComment(index) => {
            edit.remove_comment(slide, drawable, Position::new(index))?
        },
        Annotation::AddReply(index, body) => {
            edit.add_comment_reply(slide, drawable, Position::new(index), body, when()?)?
        },
        Annotation::UpdateReply(index, reply, body) => edit.update_comment_reply(
            slide,
            drawable,
            Position::new(index),
            Position::new(reply),
            body,
        )?,
        Annotation::RemoveReply(index, reply) => {
            edit.remove_comment_reply(slide, drawable, Position::new(index), Position::new(reply))?
        },
    };
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_drawables(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_annotations(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let id = host_id(&editor, bytes, leaf_position(bytes, Kind::TextBox)?)?;
    let mut reads = Vec::new();
    for highlight in editor.slide_text_box_highlights(SLIDE, id)? {
        reads.push(format!("highlight {:?}", highlight.range));
    }
    for comment in editor.slide_text_box_comments(SLIDE, id)? {
        reads.push(format!(
            "comment {:?} {:?} {} {}",
            comment.range(),
            comment.body(),
            comment.reply_count(),
            comment.author().is_some()
        ));
        for reply in editor.slide_text_box_comment_replies(SLIDE, id, comment.id())? {
            reads.push(format!(
                "reply {:?} {}",
                reply.body(),
                reply.author().is_some()
            ));
        }
    }
    Ok(reads)
}

fn leaf_annotations(bytes: &[u8]) -> TestResult<Vec<String>> {
    let package = Package::from_bytes(bytes)?;
    let (slide, drawable) = (
        SlideSelector::index(SLIDE),
        DrawableSelector::index(leaf_position(bytes, Kind::TextBox)?),
    );
    let mut reads = Vec::new();
    for highlight in package.slide_drawable_text_highlights(slide, drawable)? {
        reads.push(format!("highlight {:?}", highlight.range));
    }
    for (index, comment) in package
        .slide_drawable_text_comments(slide, drawable)?
        .into_iter()
        .enumerate()
    {
        reads.push(format!(
            "comment {:?} {:?} {} {}",
            comment.range(),
            comment.body(),
            comment.reply_count(),
            comment.author().is_some()
        ));
        for reply in
            package.slide_drawable_text_comment_replies(slide, drawable, Position::new(index))?
        {
            reads.push(format!(
                "reply {:?} {}",
                reply.body(),
                reply.author().is_some()
            ));
        }
    }
    Ok(reads)
}

fn annotate(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    operation: &Annotation,
) -> TestResult {
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| host_annotate(source, operation))
        .leaf(|source| leaf_annotate(source, operation))
        .read("host annotations", host_annotations)
        .read("leaf annotations", leaf_annotations)
        .known_fields(masked_comments)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()
        .map_err(|error| format!("{operation:?}: {error}"))?;
    Ok(())
}

fn annotated(operations: &[Annotation]) -> TestResult<Fixture> {
    let input = fixture(TEXT_BOX)?;
    let mut bytes = input.bytes;
    for operation in operations {
        bytes = host_annotate(&bytes, operation)?;
    }
    Ok(Fixture {
        path: input.path,
        bytes,
    })
}

#[test]
fn keynote_slide_text_box_highlight_add_parity() -> TestResult {
    annotate(
        "keynote_slide_text_box_highlight_add_parity",
        Tier::B,
        Evidence::Queued,
        &fixture(TEXT_BOX)?,
        &Annotation::AddHighlight(range(0, 4)?),
    )
}

#[test]
fn keynote_slide_text_box_highlight_update_parity() -> TestResult {
    annotate(
        "keynote_slide_text_box_highlight_update_parity",
        Tier::B,
        Evidence::Queued,
        &annotated(&[Annotation::AddHighlight(range(0, 4)?)])?,
        &Annotation::UpdateHighlight(0, range(1, 3)?),
    )
}

#[test]
fn keynote_slide_text_box_highlight_remove_parity() -> TestResult {
    annotate(
        "keynote_slide_text_box_highlight_remove_parity",
        Tier::B,
        Evidence::Queued,
        &annotated(&[Annotation::AddHighlight(range(0, 4)?)])?,
        &Annotation::RemoveHighlight(0),
    )
}

#[test]
fn keynote_slide_text_box_comment_add_parity() -> TestResult {
    annotate(
        "keynote_slide_text_box_comment_add_parity",
        Tier::B,
        Evidence::Queued,
        &fixture(TEXT_BOX)?,
        &Annotation::AddComment(range(0, 4)?, "Check this"),
    )
}

#[test]
fn keynote_slide_text_box_comment_update_parity() -> TestResult {
    annotate(
        "keynote_slide_text_box_comment_update_parity",
        Tier::B,
        Evidence::Queued,
        &annotated(&[Annotation::AddComment(range(0, 4)?, "Check this")])?,
        &Annotation::UpdateComment(0, range(1, 3)?, "Checked"),
    )
}

#[test]
fn keynote_slide_text_box_comment_remove_parity() -> TestResult {
    annotate(
        "keynote_slide_text_box_comment_remove_parity",
        Tier::B,
        Evidence::Queued,
        &annotated(&[
            Annotation::AddComment(range(0, 4)?, "Check this"),
            Annotation::AddReply(0, "Agreed"),
        ])?,
        &Annotation::RemoveComment(0),
    )
}

#[test]
fn keynote_slide_text_box_comment_reply_add_parity() -> TestResult {
    annotate(
        "keynote_slide_text_box_comment_reply_add_parity",
        Tier::B,
        Evidence::Queued,
        &annotated(&[Annotation::AddComment(range(0, 4)?, "Check this")])?,
        &Annotation::AddReply(0, "Agreed"),
    )
}

#[test]
fn keynote_slide_text_box_comment_reply_update_parity() -> TestResult {
    annotate(
        "keynote_slide_text_box_comment_reply_update_parity",
        Tier::B,
        Evidence::Queued,
        &annotated(&[
            Annotation::AddComment(range(0, 4)?, "Check this"),
            Annotation::AddReply(0, "Agreed"),
        ])?,
        &Annotation::UpdateReply(0, 0, "Strongly agreed"),
    )
}

#[test]
fn keynote_slide_text_box_comment_reply_remove_parity() -> TestResult {
    annotate(
        "keynote_slide_text_box_comment_reply_remove_parity",
        Tier::B,
        Evidence::Queued,
        &annotated(&[
            Annotation::AddComment(range(0, 4)?, "Check this"),
            Annotation::AddReply(0, "Agreed"),
        ])?,
        &Annotation::RemoveReply(0, 0),
    )
}

#[test]
fn keynote_slide_text_box_annotation_read_parity() -> TestResult {
    for input in [
        fixture(TEXT_BOX)?,
        annotated(&[Annotation::AddHighlight(range(0, 2)?)])?,
        annotated(&[
            Annotation::AddComment(range(0, 4)?, "Check this"),
            Annotation::AddReply(0, "Agreed"),
        ])?,
    ] {
        assert_eq!(
            host_annotations(&input.bytes)?,
            leaf_annotations(&input.bytes)?
        );
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
