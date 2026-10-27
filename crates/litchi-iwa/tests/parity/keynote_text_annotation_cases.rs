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

fn masked_date_time(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tswp::DateTimeSmartFieldArchive::decode(data)?;
    if let Some(field) = message.super_.as_mut()
        && field
            .text_attribute_uuid_string
            .as_deref()
            .is_some_and(|uuid| TextAttributeUuid::new(uuid).is_ok())
    {
        field.text_attribute_uuid_string = Some("<uuid>".to_owned());
    }
    let mut decoded = decode_as::<tswp::DateTimeSmartFieldArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn date_time_settings() -> TestResult<[litchi_iwa_text::date_time::Settings; 2]> {
    use litchi_iwa_text::date_time::{
        Format as DateFormat, FormatterStyle, Instant, LocaleIdentifier, Settings, UpdatePlan,
    };
    Ok([
        Settings::fixed(
            DateFormat::new("yyyy-MM-dd")?,
            LocaleIdentifier::new("en_US")?,
            Instant::from_reference_date_seconds(735_000_000.5)?,
        ),
        Settings::new(
            None,
            Some(LocaleIdentifier::new("fr_FR")?),
            Some(FormatterStyle::Long),
            Some(FormatterStyle::Short),
            Some(UpdatePlan::Automatic),
            Some(true),
            None,
        )?,
    ])
}

#[derive(Clone, Debug)]
enum Field {
    Add(TextRange, litchi_iwa_text::date_time::Settings),
    Insert(
        TextPosition,
        &'static str,
        litchi_iwa_text::date_time::Settings,
    ),
    Update(usize, TextRange, litchi_iwa_text::date_time::Settings),
    Remove(usize),
}

fn host_field(source: &[u8], operation: &Field) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let id = host_id(&editor, source, leaf_position(source, Kind::TextBox)?)?;
    match operation.clone() {
        Field::Add(range, settings) => {
            editor.add_slide_text_box_date_time_field(SLIDE, id, range, settings)?;
        },
        Field::Insert(position, display, settings) => {
            editor.insert_slide_text_box_date_time_field(
                SLIDE,
                id,
                position,
                litchi_iwa_text::date_time::DisplayText::new(display)?,
                settings,
            )?;
        },
        Field::Update(index, range, settings) => {
            let field = editor.slide_text_box_date_time_fields(SLIDE, id)?[index].id;
            editor.update_slide_text_box_date_time_field(SLIDE, id, field, range, settings)?;
        },
        Field::Remove(index) => {
            let field = editor.slide_text_box_date_time_fields(SLIDE, id)?[index].id;
            editor.remove_slide_text_box_date_time_field(SLIDE, id, field)?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_field(source: &[u8], operation: &Field) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let (slide, drawable) = (
        SlideSelector::index(SLIDE),
        DrawableSelector::index(leaf_position(source, Kind::TextBox)?),
    );
    let mut edit = package.edit_slide_drawables();
    match operation.clone() {
        Field::Add(range, settings) => {
            edit.add_date_time_field(slide, drawable, range, settings)?
        },
        Field::Insert(position, display, settings) => edit.insert_date_time_field(
            slide,
            drawable,
            position,
            &litchi_iwa_text::date_time::DisplayText::new(display)?,
            settings,
        )?,
        Field::Update(index, range, settings) => {
            edit.update_date_time_field(slide, drawable, Position::new(index), range, settings)?
        },
        Field::Remove(index) => {
            edit.remove_date_time_field(slide, drawable, Position::new(index))?
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

fn host_date_fields(bytes: &[u8]) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let id = host_id(&editor, bytes, leaf_position(bytes, Kind::TextBox)?)?;
    Ok(format!(
        "{:?}",
        editor
            .slide_text_box_date_time_fields(SLIDE, id)?
            .into_iter()
            .map(|field| (field.range, field.settings))
            .collect::<Vec<_>>()
    ))
}

fn leaf_date_fields(bytes: &[u8]) -> TestResult<String> {
    let position = leaf_position(bytes, Kind::TextBox)?;
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?
            .slide_drawable_date_time_fields(SLIDE, DrawableSelector::index(position))?
    ))
}

fn field_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    operation: &Field,
) -> TestResult {
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| host_field(source, operation))
        .leaf(|source| leaf_field(source, operation))
        .read("host fields", host_date_fields)
        .read("leaf fields", leaf_date_fields)
        .known_fields(|known| {
            known.with(
                smart_field::DATE_TIME_MESSAGE_TYPE,
                "TSWP.DateTimeSmartFieldArchive (text-attribute UUID masked)",
                masked_date_time,
            )
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()
        .map_err(|error| format!("{operation:?}: {error}"))?;
    Ok(())
}

fn dated() -> TestResult<Fixture> {
    let input = fixture(TEXT_BOX)?;
    Ok(Fixture {
        path: input.path,
        bytes: host_field(
            &input.bytes,
            &Field::Add(range(0, 4)?, date_time_settings()?[0].clone()),
        )?,
    })
}

#[test]
fn keynote_slide_text_box_date_time_add_parity() -> TestResult {
    for settings in date_time_settings()? {
        field_case(
            "keynote_slide_text_box_date_time_add_parity",
            Tier::B,
            Evidence::Queued,
            &fixture(TEXT_BOX)?,
            &Field::Add(range(0, 4)?, settings),
        )?;
    }
    Ok(())
}

#[test]
fn keynote_slide_text_box_date_time_insert_parity() -> TestResult {
    field_case(
        "keynote_slide_text_box_date_time_insert_parity",
        Tier::B,
        Evidence::Queued,
        &fixture(TEXT_BOX)?,
        &Field::Insert(
            TextPosition::from_utf16_code_units(2),
            "2026-10-26",
            date_time_settings()?[0].clone(),
        ),
    )
}

#[test]
fn keynote_slide_text_box_date_time_update_parity() -> TestResult {
    field_case(
        "keynote_slide_text_box_date_time_update_parity",
        Tier::B,
        Evidence::Queued,
        &dated()?,
        &Field::Update(0, range(1, 3)?, date_time_settings()?[1].clone()),
    )
}

#[test]
fn keynote_slide_text_box_date_time_remove_parity() -> TestResult {
    field_case(
        "keynote_slide_text_box_date_time_remove_parity",
        Tier::B,
        Evidence::Queued,
        &dated()?,
        &Field::Remove(0),
    )
}

#[test]
fn keynote_slide_text_box_date_time_read_parity() -> TestResult {
    for input in [fixture(TEXT_BOX)?, dated()?] {
        assert_eq!(
            host_date_fields(&input.bytes)?,
            leaf_date_fields(&input.bytes)?
        );
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

const DROP_CAP_FIXTURE: &str = "iwork/keynote/chart-grid-authored-native.key";

fn drop_cap_samples() -> TestResult<[litchi_iwa_text::paragraph::drop_cap::DropCap; 2]> {
    use litchi_iwa_text::paragraph::drop_cap::{
        CharacterCount, DropCap, LineCount, Outdent, Padding, RaisedLines, Wrap,
    };
    Ok([
        DropCap::new(LineCount::new(3)?, CharacterCount::new(1)?),
        DropCap::new(LineCount::new(4)?, CharacterCount::new(2)?)
            .with_raised_lines(RaisedLines::new(1)?)
            .with_wrap(Wrap::Contour)
            .with_padding(Padding::from_points(6.0)?)
            .with_outdent(Outdent::from_ratio(0.25)?),
    ])
}

/// The first (slide, leaf drawable position, host id) text box of the
/// fixture whose first paragraph takes a Drop Cap on the host.
fn drop_cap_target(bytes: &[u8]) -> TestResult<(usize, usize, u64)> {
    let package = Package::from_bytes(bytes)?;
    let editor = keynote::Draft::from_bytes(bytes)?;
    for slide in 0..package.slides()?.len() {
        let list = package.slide_drawable_list(SlideSelector::index(slide))?;
        for (position, drawable) in list.iter().enumerate() {
            if drawable.kind() != Kind::TextBox || drawable.text().is_none_or(str::is_empty) {
                continue;
            }
            let z = drawable.z_position().ok_or("no z")?;
            let id = editor.slide_drawable_order(slide)?[z.get()];
            let mut probe = keynote::Draft::from_bytes(bytes)?;
            if probe
                .set_slide_text_box_paragraph_drop_cap(
                    slide,
                    id,
                    TextPosition::ZERO,
                    drop_cap_samples()?[0],
                )
                .is_ok()
            {
                return Ok((slide, position, id));
            }
        }
    }
    Err("no text box takes a Drop Cap".into())
}

fn host_drop_cap(
    source: &[u8],
    value: Option<litchi_iwa_text::paragraph::drop_cap::DropCap>,
) -> TestResult<Vec<u8>> {
    let (slide, _, id) = drop_cap_target(source)?;
    let mut editor = keynote::Draft::from_bytes(source)?;
    match value {
        Some(value) => {
            editor.set_slide_text_box_paragraph_drop_cap(slide, id, TextPosition::ZERO, value)?
        },
        None => {
            editor.remove_slide_text_box_paragraph_drop_cap(slide, id, TextPosition::ZERO)?;
        },
    }
    Ok(editor.to_bytes()?)
}

fn leaf_drop_cap(
    source: &[u8],
    value: Option<litchi_iwa_text::paragraph::drop_cap::DropCap>,
) -> TestResult<LeafRun> {
    let (slide, position, _) = drop_cap_target(source)?;
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_drawables();
    edit.set_drop_cap(
        SlideSelector::index(slide),
        DrawableSelector::index(position),
        TextPosition::ZERO,
        value,
    )?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_drawables(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_drop_caps(bytes: &[u8]) -> TestResult<String> {
    let (slide, _, id) = drop_cap_target(bytes)?;
    let editor = keynote::Draft::from_bytes(bytes)?;
    Ok(format!(
        "{:?} {:?}",
        editor
            .slide_text_box_paragraph_drop_caps(slide, id)?
            .into_iter()
            .map(|placement| (placement.paragraph, placement.drop_cap))
            .collect::<Vec<_>>(),
        editor.slide_text_box_paragraph_drop_cap(slide, id, TextPosition::ZERO)?
    ))
}

fn leaf_drop_caps(bytes: &[u8]) -> TestResult<String> {
    let (slide, position, _) = drop_cap_target(bytes)?;
    let caps = Package::from_bytes(bytes)?.slide_drawable_drop_caps(
        SlideSelector::index(slide),
        DrawableSelector::index(position),
    )?;
    let first = caps
        .iter()
        .find(|(paragraph, _)| *paragraph == TextPosition::ZERO)
        .map(|(_, value)| *value);
    Ok(format!("{caps:?} {first:?}"))
}

fn drop_cap_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    value: Option<litchi_iwa_text::paragraph::drop_cap::DropCap>,
) -> TestResult {
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| host_drop_cap(source, value))
        .leaf(|source| leaf_drop_cap(source, value))
        .read("host drop caps", host_drop_caps)
        .read("leaf drop caps", leaf_drop_caps)
        .known_fields(|known| {
            known.with(
                litchi_iwa_text_wire::drop_cap::MESSAGE_TYPE,
                "TSWP.DropCapStyleArchive",
                decode_as::<tswp::DropCapStyleArchive>,
            )
        })
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()?;
    Ok(())
}

/// The Drop Cap fixture with one host-added text box on slide 0 (the
/// fixture's styled text lives in placeholders).
fn drop_cap_input() -> TestResult<Fixture> {
    let input = fixture(DROP_CAP_FIXTURE)?;
    let mut editor = keynote::Draft::from_bytes(&input.bytes)?;
    editor.add_slide_text_box(
        0,
        "Drop caps open this paragraph",
        litchi_iwa::shapes::DrawablePoint { x: 80.0, y: 120.0 },
        litchi_iwa::shapes::DrawableSize {
            width: 400.0,
            height: 120.0,
        },
    )?;
    Ok(Fixture {
        path: input.path,
        bytes: editor.to_bytes()?,
    })
}

fn capped() -> TestResult<Fixture> {
    let input = drop_cap_input()?;
    Ok(Fixture {
        path: input.path,
        bytes: host_drop_cap(&input.bytes, Some(drop_cap_samples()?[0]))?,
    })
}

#[test]
fn keynote_slide_text_box_drop_cap_set_parity() -> TestResult {
    for value in drop_cap_samples()? {
        drop_cap_case(
            "keynote_slide_text_box_drop_cap_set_parity",
            Tier::B,
            Evidence::Queued,
            &drop_cap_input()?,
            Some(value),
        )?;
    }
    drop_cap_case(
        "keynote_slide_text_box_drop_cap_set_parity",
        Tier::B,
        Evidence::Queued,
        &capped()?,
        Some(drop_cap_samples()?[1]),
    )
}

#[test]
fn keynote_slide_text_box_drop_cap_remove_parity() -> TestResult {
    drop_cap_case(
        "keynote_slide_text_box_drop_cap_remove_parity",
        Tier::B,
        Evidence::Queued,
        &capped()?,
        None,
    )
}

#[test]
fn keynote_slide_text_box_drop_cap_read_parity() -> TestResult {
    for input in [drop_cap_input()?, capped()?] {
        assert_eq!(host_drop_caps(&input.bytes)?, leaf_drop_caps(&input.bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

fn host_language(
    source: &[u8],
    change: &litchi_iwa_text_wire::language::Change,
) -> TestResult<Vec<u8>> {
    use litchi_iwa_text_wire::language::Change;
    let mut editor = keynote::Draft::from_bytes(source)?;
    let id = host_id(&editor, source, leaf_position(source, Kind::TextBox)?)?;
    match change.clone() {
        Change::Set { position, language } => {
            editor.set_slide_text_box_text_language(SLIDE, id, position, language)?;
        },
        Change::RemoveBoundary(position) => {
            editor.remove_slide_text_box_text_language_boundary(SLIDE, id, position)?;
        },
        Change::Reset => {
            editor.reset_slide_text_box_text_languages(SLIDE, id)?;
        },
        other => return Err(format!("unmapped change {other:?}").into()),
    }
    Ok(editor.to_bytes()?)
}

fn leaf_language(
    source: &[u8],
    change: &litchi_iwa_text_wire::language::Change,
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_drawables();
    edit.set_text_language(
        SlideSelector::index(SLIDE),
        DrawableSelector::index(leaf_position(source, Kind::TextBox)?),
        change.clone(),
    )?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_drawables(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn host_languages(bytes: &[u8]) -> TestResult<String> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let id = host_id(&editor, bytes, leaf_position(bytes, Kind::TextBox)?)?;
    Ok(format!(
        "{:?} {:?}",
        editor.slide_text_box_text_languages(SLIDE, id)?,
        editor.slide_text_box_text_language(SLIDE, id, TextPosition::from_utf16_code_units(2))?
    ))
}

fn leaf_languages(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let (slide, drawable) = (
        SlideSelector::index(SLIDE),
        DrawableSelector::index(leaf_position(bytes, Kind::TextBox)?),
    );
    Ok(format!(
        "{:?} {:?}",
        package.slide_drawable_text_languages(slide, drawable)?,
        package.slide_drawable_text_language(
            slide,
            drawable,
            TextPosition::from_utf16_code_units(2)
        )?
    ))
}

fn language_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    input: &Fixture,
    change: &litchi_iwa_text_wire::language::Change,
) -> TestResult {
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| host_language(source, change))
        .leaf(|source| leaf_language(source, change))
        .read("host languages", host_languages)
        .read("leaf languages", leaf_languages)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .run()
        .map_err(|error| format!("{change:?}: {error}"))?;
    Ok(())
}

fn set_language(at: u32, tag: &str) -> TestResult<litchi_iwa_text_wire::language::Change> {
    Ok(litchi_iwa_text_wire::language::Change::Set {
        position: TextPosition::from_utf16_code_units(at),
        language: litchi_iwa_text::language::TextLanguage::tag(tag)?,
    })
}

fn tagged() -> TestResult<Fixture> {
    let input = fixture(TEXT_BOX)?;
    let mut bytes = input.bytes;
    for (at, tag) in [(0, "en"), (3, "fr")] {
        bytes = host_language(&bytes, &set_language(at, tag)?)?;
    }
    Ok(Fixture {
        path: input.path,
        bytes,
    })
}

#[test]
fn keynote_slide_text_box_language_set_parity() -> TestResult {
    let (input, base) = (fixture(TEXT_BOX)?, tagged()?);
    for (fixture, at, tag) in [
        (&input, 0, "de"),
        (&input, 3, "ja"),
        (&base, 3, "es"),
        (&base, 0, "fr"),
    ] {
        language_case(
            "keynote_slide_text_box_language_set_parity",
            Tier::B,
            Evidence::Queued,
            fixture,
            &set_language(at, tag)?,
        )?;
    }
    Ok(())
}

#[test]
fn keynote_slide_text_box_language_remove_parity() -> TestResult {
    language_case(
        "keynote_slide_text_box_language_remove_parity",
        Tier::B,
        Evidence::Queued,
        &tagged()?,
        &litchi_iwa_text_wire::language::Change::RemoveBoundary(
            TextPosition::from_utf16_code_units(3),
        ),
    )
}

#[test]
fn keynote_slide_text_box_language_reset_parity() -> TestResult {
    language_case(
        "keynote_slide_text_box_language_reset_parity",
        Tier::B,
        Evidence::Queued,
        &tagged()?,
        &litchi_iwa_text_wire::language::Change::Reset,
    )
}

#[test]
fn keynote_slide_text_box_language_read_parity() -> TestResult {
    for input in [fixture(TEXT_BOX)?, tagged()?] {
        assert_eq!(host_languages(&input.bytes)?, leaf_languages(&input.bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
