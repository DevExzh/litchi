//! Pages body text-box languages, Drop Caps, hyperlinks, Date & Time fields, highlights and comments parity (wave W4-P, record 0413).

use litchi_iwa::pages;
use litchi_pages::DrawableSelector;
use litchi_pages::Package;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::known_fields::decode_as;
use crate::leaf_bytes;
use crate::pages_text_box_cases::{host_text_box, text_box_fixture as authored_fixture};
use litchi_iwa_protos::tswp;

// ---- languages -----------------------------------------------------------

use litchi_pages::drawable::language::{TextLanguage, TextPosition};

#[derive(Clone, Debug)]
enum LanguageChange {
    Set(usize, &'static str),
    Remove(usize),
    Reset,
}

fn host_languages(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = host_text_box(&e, 0)?;
    Ok(format!("{:?}", e.text_box_text_languages(d)?))
}

fn leaf_languages(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?.body_drawable_text_languages(DrawableSelector::text_box(0))?
    ))
}

fn host_language_change(source: &[u8], change: &LanguageChange) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let d = host_text_box(&e, 0)?;
    match change {
        LanguageChange::Set(position, tag) => e.set_text_box_text_language(
            d,
            TextPosition::from_utf16_index(*position)?,
            TextLanguage::tag(*tag)?,
        )?,
        LanguageChange::Remove(position) => drop(e.remove_text_box_text_language_boundary(
            d,
            TextPosition::from_utf16_index(*position)?,
        )?),
        LanguageChange::Reset => drop(e.reset_text_box_text_languages(d)?),
    }
    Ok(e.to_bytes()?)
}

fn language_case(name: &'static str, input: &Fixture, change: LanguageChange) -> TestResult {
    let change = &change;
    Case::new(name, Format::Pages, Tier::A, Evidence::Queued, input)
        .host(|source| host_language_change(source, change))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_body_drawable_text_languages(DrawableSelector::text_box(0))?;
            match change {
                LanguageChange::Set(position, tag) => {
                    edit.set(
                        TextPosition::from_utf16_index(*position)?,
                        TextLanguage::tag(*tag)?,
                    )?;
                },
                LanguageChange::Remove(position) => {
                    edit.remove_boundary(TextPosition::from_utf16_index(*position)?)?;
                },
                LanguageChange::Reset => {
                    edit.reset()?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_drawable_text_languages(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host languages", host_languages)
        .read("leaf languages", leaf_languages)
        .run()?;
    Ok(())
}

/// The authored fixture with French from 6 set by the host.
fn with_boundary() -> TestResult<Fixture> {
    let input = authored_fixture()?;
    Ok(Fixture {
        path: "authored + host French boundary at 6",
        bytes: host_language_change(&input.bytes, &LanguageChange::Set(6, "fr"))?,
    })
}

#[test]
fn pages_text_box_language_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    language_case(
        "pages_text_box_language_set_parity",
        &authored_fixture()?,
        LanguageChange::Set(6, "fr"),
    )?;
    language_case(
        "pages_text_box_language_set_parity",
        &with_boundary()?,
        LanguageChange::Set(0, "de"),
    )
}

#[test]
fn pages_text_box_language_remove_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    language_case(
        "pages_text_box_language_remove_parity",
        &with_boundary()?,
        LanguageChange::Remove(6),
    )
}

#[test]
fn pages_text_box_language_reset_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    language_case(
        "pages_text_box_language_reset_parity",
        &with_boundary()?,
        LanguageChange::Reset,
    )
}

#[test]
fn pages_text_box_language_read_parity() -> TestResult {
    for input in [authored_fixture()?, with_boundary()?] {
        assert_eq!(host_languages(&input.bytes)?, leaf_languages(&input.bytes)?);
        let e = pages::Draft::from_bytes(&input.bytes)?;
        let d = host_text_box(&e, 0)?;
        for position in [0, 3, 6, 8] {
            let position = TextPosition::from_utf16_index(position)?;
            assert_eq!(
                e.text_box_text_language(d, position)?,
                Package::from_bytes(&input.bytes)?
                    .body_drawable_text_language(DrawableSelector::text_box(0), position)?
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- Drop Caps -------------------------------------------------------------

use litchi_pages::drawable::drop_cap::{
    CharacterCount, DropCap, LineCount, Outdent, Padding, RaisedLines, Wrap,
};

/// `chart-grid-authored-native.numbers` (its theme has a Drop Cap base) with
/// one host text box of two paragraphs.
fn drop_cap_fixture() -> TestResult<Fixture> {
    let mut editor = pages::Draft::from_bytes(
        &crate::harness::fixture("iwork/pages/chart-grid-authored-native.pages")?.bytes,
    )?;
    let anchor = editor.body_text()?.encode_utf16().count();
    editor.add_text_box(
        anchor,
        "First paragraph\nSecond paragraph",
        litchi_iwa::shapes::DrawablePoint { x: 300.0, y: 200.0 },
        litchi_iwa::shapes::DrawableSize {
            width: 220.0,
            height: 90.0,
        },
    )?;
    Ok(Fixture {
        path: "chart-grid-authored-native.pages + add_text_box",
        bytes: editor.to_bytes()?,
    })
}

fn drop_cap_samples() -> TestResult<[DropCap; 2]> {
    Ok([
        DropCap::new(LineCount::new(3)?, CharacterCount::new(1)?),
        DropCap::new(LineCount::new(4)?, CharacterCount::new(2)?)
            .with_raised_lines(RaisedLines::new(1)?)
            .with_wrap(Wrap::Contour)
            .with_padding(Padding::from_points(6.0)?)
            .with_outdent(Outdent::from_ratio(0.25)?),
    ])
}

fn host_drop_caps(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = host_text_box(&e, 0)?;
    Ok(format!(
        "{:?}",
        e.text_box_paragraph_drop_caps(d)?
            .into_iter()
            .map(|placement| (placement.paragraph, placement.drop_cap))
            .collect::<Vec<_>>()
    ))
}

fn leaf_drop_caps(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?.body_drawable_drop_caps(DrawableSelector::text_box(0))?
    ))
}

fn host_drop_cap(
    source: &[u8],
    paragraph: TextPosition,
    value: Option<DropCap>,
) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let d = host_text_box(&e, 0)?;
    match value {
        Some(value) => e.set_text_box_paragraph_drop_cap(d, paragraph, value)?,
        None => drop(e.remove_text_box_paragraph_drop_cap(d, paragraph)?),
    }
    Ok(e.to_bytes()?)
}

fn drop_cap_case(
    name: &'static str,
    tier: Tier,
    input: &Fixture,
    paragraph: TextPosition,
    value: Option<DropCap>,
) -> TestResult {
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, input)
        .host(|source| host_drop_cap(source, paragraph, value))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_drop_caps(DrawableSelector::text_box(0))?;
            match value {
                Some(value) => {
                    edit.set(paragraph, value)?;
                },
                None => {
                    edit.remove(paragraph)?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_drawable_drop_caps(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host drop caps", host_drop_caps)
        .read("leaf drop caps", leaf_drop_caps)
        .known_fields(|known| {
            known.with(
                10_024,
                "TSWP.DropCapStyleArchive",
                decode_as::<tswp::DropCapStyleArchive>,
            )
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn second_paragraph() -> TestResult<TextPosition> {
    Ok(TextPosition::from_utf16_index(
        "First paragraph\n".encode_utf16().count(),
    )?)
}

#[test]
fn pages_text_box_drop_cap_set_parity() -> TestResult {
    let input = drop_cap_fixture()?;
    let [first, second] = drop_cap_samples()?;
    for paragraph in [TextPosition::ZERO, second_paragraph()?] {
        drop_cap_case(
            "pages_text_box_drop_cap_set_parity",
            Tier::B,
            &input,
            paragraph,
            Some(first),
        )?;
    }
    // A second value on the paragraph's own variation rewrites it in place.
    let set = Fixture {
        path: input.path,
        bytes: host_drop_cap(&input.bytes, TextPosition::ZERO, Some(first))?,
    };
    drop_cap_case(
        "pages_text_box_drop_cap_set_parity",
        Tier::B,
        &set,
        TextPosition::ZERO,
        Some(second),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_drop_cap_remove_parity() -> TestResult {
    let input = drop_cap_fixture()?;
    let [first, _] = drop_cap_samples()?;
    let set = Fixture {
        path: input.path,
        bytes: host_drop_cap(&input.bytes, TextPosition::ZERO, Some(first))?,
    };
    drop_cap_case(
        "pages_text_box_drop_cap_remove_parity",
        Tier::B,
        &set,
        TextPosition::ZERO,
        None,
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_drop_cap_read_parity() -> TestResult {
    let input = drop_cap_fixture()?;
    let [first, _] = drop_cap_samples()?;
    let set = host_drop_cap(&input.bytes, second_paragraph()?, Some(first))?;
    for bytes in [input.bytes.clone(), set] {
        assert_eq!(host_drop_caps(&bytes)?, leaf_drop_caps(&bytes)?);
        let e = pages::Draft::from_bytes(&bytes)?;
        let d = host_text_box(&e, 0)?;
        let leaf =
            Package::from_bytes(&bytes)?.body_drawable_drop_caps(DrawableSelector::text_box(0))?;
        for paragraph in [TextPosition::ZERO, second_paragraph()?] {
            assert_eq!(
                e.text_box_paragraph_drop_cap(d, paragraph)?,
                leaf.iter()
                    .find(|(at, _)| *at == paragraph)
                    .map(|(_, cap)| *cap)
            );
        }
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- hyperlinks ------------------------------------------------------------

use litchi_pages::drawable::hyperlink::{TextHyperlinkTarget, TextRange};

#[derive(Clone, Debug)]
enum LinkChange {
    Add(TextRange, &'static str),
    Update(usize, TextRange, &'static str),
    Remove(usize),
}

fn range(start: usize, end: usize) -> TestResult<TextRange> {
    Ok(TextRange::new(
        TextPosition::from_utf16_index(start)?,
        TextPosition::from_utf16_index(end)?,
    )?)
}

fn host_links(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = host_text_box(&e, 0)?;
    Ok(format!(
        "{:?}",
        e.text_box_hyperlinks(d)?
            .into_iter()
            .map(|link| (link.range, link.target))
            .collect::<Vec<_>>()
    ))
}

fn leaf_links(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?
            .body_drawable_hyperlinks(DrawableSelector::text_box(0))?
            .into_iter()
            .map(|link| (link.range(), link.target().clone()))
            .collect::<Vec<_>>()
    ))
}

fn host_link(source: &[u8], change: &LinkChange) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let d = host_text_box(&e, 0)?;
    let links = e.text_box_hyperlinks(d)?;
    match change {
        LinkChange::Add(range, target) => {
            drop(e.add_text_box_hyperlink(d, *range, TextHyperlinkTarget::new(target)?)?);
        },
        LinkChange::Update(position, range, target) => drop(e.update_text_box_hyperlink(
            d,
            links.get(*position).ok_or("host link")?.id,
            *range,
            TextHyperlinkTarget::new(target)?,
        )?),
        LinkChange::Remove(position) => {
            drop(e.remove_text_box_hyperlink(d, links.get(*position).ok_or("host link")?.id)?)
        },
    }
    Ok(e.to_bytes()?)
}

fn link_case(name: &'static str, tier: Tier, input: &Fixture, change: LinkChange) -> TestResult {
    let change = &change;
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, input)
        .host(|source| host_link(source, change))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_hyperlinks(DrawableSelector::text_box(0))?;
            match change {
                LinkChange::Add(range, target) => {
                    edit.add(*range, TextHyperlinkTarget::new(target)?)?;
                },
                LinkChange::Update(position, range, target) => {
                    edit.update(*position, *range, TextHyperlinkTarget::new(target)?)?;
                },
                LinkChange::Remove(position) => {
                    edit.remove(*position)?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_drawable_hyperlinks(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host hyperlinks", host_links)
        .read("leaf hyperlinks", leaf_links);
    if tier == Tier::B {
        case = case
            .known_fields(crate::text_annotation_cases::masked)
            .leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn with_link() -> TestResult<Fixture> {
    let input = authored_fixture()?;
    Ok(Fixture {
        path: "authored + host hyperlink over 0..5",
        bytes: host_link(
            &input.bytes,
            &LinkChange::Add(range(0, 5)?, "https://example.com/litchi"),
        )?,
    })
}

#[test]
fn pages_text_box_hyperlink_add_parity() -> TestResult {
    link_case(
        "pages_text_box_hyperlink_add_parity",
        Tier::B,
        &authored_fixture()?,
        LinkChange::Add(range(0, 5)?, "https://example.com/litchi"),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_hyperlink_update_parity() -> TestResult {
    link_case(
        "pages_text_box_hyperlink_update_parity",
        Tier::A,
        &with_link()?,
        LinkChange::Update(0, range(2, 8)?, "mailto:a@example.com"),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_hyperlink_remove_parity() -> TestResult {
    link_case(
        "pages_text_box_hyperlink_remove_parity",
        Tier::A,
        &with_link()?,
        LinkChange::Remove(0),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_hyperlink_read_parity() -> TestResult {
    for input in [authored_fixture()?, with_link()?] {
        assert_eq!(host_links(&input.bytes)?, leaf_links(&input.bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- Date & Time fields ------------------------------------------------------

use litchi_pages::drawable::date_time::DisplayText;

#[derive(Clone, Debug)]
enum FieldChange {
    Add(TextRange, usize),
    Insert(TextPosition, &'static str, usize),
    Update(usize, TextRange, usize),
    Remove(usize),
}

fn host_dates(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = host_text_box(&e, 0)?;
    Ok(format!(
        "{:?}",
        e.text_box_date_time_fields(d)?
            .into_iter()
            .map(|field| (field.range, field.settings))
            .collect::<Vec<_>>()
    ))
}

fn leaf_dates(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:?}",
        Package::from_bytes(bytes)?
            .body_drawable_date_time_fields(DrawableSelector::text_box(0))?
            .into_iter()
            .map(|field| (field.range(), field.settings().clone()))
            .collect::<Vec<_>>()
    ))
}

fn host_date(source: &[u8], change: &FieldChange) -> TestResult<Vec<u8>> {
    let samples = crate::text_field_cases::settings()?;
    let mut e = pages::Draft::from_bytes(source)?;
    let d = host_text_box(&e, 0)?;
    let fields = e.text_box_date_time_fields(d)?;
    match change {
        FieldChange::Add(range, value) => {
            drop(e.add_text_box_date_time_field(d, *range, samples[*value].clone())?)
        },
        FieldChange::Insert(position, display, value) => {
            drop(e.insert_text_box_date_time_field(
                d,
                *position,
                DisplayText::new(display)?,
                samples[*value].clone(),
            )?);
        },
        FieldChange::Update(position, range, value) => {
            drop(e.update_text_box_date_time_field(
                d,
                fields.get(*position).ok_or("host field")?.id,
                *range,
                samples[*value].clone(),
            )?);
        },
        FieldChange::Remove(position) => drop(
            e.remove_text_box_date_time_field(d, fields.get(*position).ok_or("host field")?.id)?,
        ),
    }
    Ok(e.to_bytes()?)
}

fn date_case(name: &'static str, tier: Tier, input: &Fixture, change: FieldChange) -> TestResult {
    let change = &change;
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, input)
        .host(|source| host_date(source, change))
        .leaf(|source| {
            let samples = crate::text_field_cases::settings()?;
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_body_drawable_date_time_fields(DrawableSelector::text_box(0))?;
            match change {
                FieldChange::Add(range, value) => {
                    edit.add(*range, samples[*value].clone())?;
                },
                FieldChange::Insert(position, display, value) => {
                    edit.insert(
                        *position,
                        DisplayText::new(display)?,
                        samples[*value].clone(),
                    )?;
                },
                FieldChange::Update(position, range, value) => {
                    edit.update(*position, *range, samples[*value].clone())?;
                },
                FieldChange::Remove(position) => {
                    edit.remove(*position)?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_drawable_date_time_fields(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host fields", host_dates)
        .read("leaf fields", leaf_dates);
    if tier == Tier::B {
        case = case
            .known_fields(crate::text_field_cases::masked)
            .leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn with_date() -> TestResult<Fixture> {
    let input = authored_fixture()?;
    Ok(Fixture {
        path: "authored + host date-time field over 0..5",
        bytes: host_date(&input.bytes, &FieldChange::Add(range(0, 5)?, 0))?,
    })
}

#[test]
fn pages_text_box_date_time_add_parity() -> TestResult {
    for value in [0, 1] {
        date_case(
            "pages_text_box_date_time_add_parity",
            Tier::B,
            &authored_fixture()?,
            FieldChange::Add(range(0, 5)?, value),
        )?;
    }
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_date_time_insert_parity() -> TestResult {
    for position in [0, 9] {
        date_case(
            "pages_text_box_date_time_insert_parity",
            Tier::B,
            &authored_fixture()?,
            FieldChange::Insert(TextPosition::from_utf16_index(position)?, "2026-10-26", 0),
        )?;
    }
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_date_time_update_parity() -> TestResult {
    date_case(
        "pages_text_box_date_time_update_parity",
        Tier::A,
        &with_date()?,
        FieldChange::Update(0, range(2, 8)?, 1),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_date_time_remove_parity() -> TestResult {
    date_case(
        "pages_text_box_date_time_remove_parity",
        Tier::A,
        &with_date()?,
        FieldChange::Remove(0),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_date_time_read_parity() -> TestResult {
    for input in [authored_fixture()?, with_date()?] {
        assert_eq!(host_dates(&input.bytes)?, leaf_dates(&input.bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- highlights, comments, replies --------------------------------------------

use litchi_iwa_text::comment::{TextCommentBody, TextCommentReplyBody};

#[derive(Clone, Debug)]
enum NoteChange {
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

fn host_notes(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = host_text_box(&e, 0)?;
    let highlights: Vec<_> = e
        .text_box_highlights(d)?
        .into_iter()
        .map(|item| item.range)
        .collect();
    let mut comments = Vec::new();
    for comment in e.text_box_comments(d)? {
        let replies: Vec<String> = e
            .text_box_comment_replies(d, comment.id())?
            .into_iter()
            .map(|reply| reply.body().as_str().to_owned())
            .collect();
        comments.push((comment.range(), comment.body().as_str().to_owned(), replies));
    }
    Ok(format!("{highlights:?}\n{comments:?}"))
}

fn leaf_notes(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let t = DrawableSelector::text_box(0);
    let highlights: Vec<_> = package
        .body_drawable_highlights(t)?
        .into_iter()
        .map(|item| item.range())
        .collect();
    let comments: Vec<_> = package
        .body_drawable_comments(t)?
        .into_iter()
        .map(|item| {
            (
                item.range(),
                item.body().to_owned(),
                item.replies().to_vec(),
            )
        })
        .collect();
    Ok(format!("{highlights:?}\n{comments:?}"))
}

fn host_note(source: &[u8], change: &NoteChange) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let d = host_text_box(&e, 0)?;
    let highlights = e.text_box_highlights(d)?;
    let comments = e.text_box_comments(d)?;
    let comment = |position: usize| -> TestResult<_> {
        Ok(comments.get(position).ok_or("host comment")?.id())
    };
    match change {
        NoteChange::AddHighlight(range) => drop(e.add_text_box_highlight(d, *range)?),
        NoteChange::UpdateHighlight(p, range) => drop(e.update_text_box_highlight(
            d,
            highlights.get(*p).ok_or("host highlight")?.id,
            *range,
        )?),
        NoteChange::RemoveHighlight(p) => {
            e.remove_text_box_highlight(d, highlights.get(*p).ok_or("host highlight")?.id)?;
        },
        NoteChange::AddComment(range, body) => {
            drop(e.add_text_box_comment(d, *range, TextCommentBody::new(body)?)?)
        },
        NoteChange::UpdateComment(p, range, body) => {
            drop(e.update_text_box_comment(d, comment(*p)?, *range, TextCommentBody::new(body)?)?)
        },
        NoteChange::RemoveComment(p) => drop(e.remove_text_box_comment(d, comment(*p)?)?),
        NoteChange::AddReply(p, body) => {
            drop(e.add_text_box_comment_reply(d, comment(*p)?, TextCommentReplyBody::new(body)?)?)
        },
        NoteChange::UpdateReply(p, r, body) => {
            let id = comment(*p)?;
            let reply = e.text_box_comment_replies(d, id)?;
            drop(e.update_text_box_comment_reply(
                d,
                id,
                reply.get(*r).ok_or("host reply")?.id(),
                TextCommentReplyBody::new(body)?,
            )?);
        },
        NoteChange::RemoveReply(p, r) => {
            let id = comment(*p)?;
            let reply = e.text_box_comment_replies(d, id)?;
            drop(e.remove_text_box_comment_reply(
                d,
                id,
                reply.get(*r).ok_or("host reply")?.id(),
            )?);
        },
    }
    Ok(e.to_bytes()?)
}

fn note_case(name: &'static str, tier: Tier, input: &Fixture, change: NoteChange) -> TestResult {
    let change = &change;
    let mut case = Case::new(name, Format::Pages, tier, Evidence::Queued, input)
        .host(|source| host_note(source, change))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_annotations(DrawableSelector::text_box(0))?;
            match change {
                NoteChange::AddHighlight(range) => {
                    edit.add_highlight(*range)?;
                },
                NoteChange::UpdateHighlight(p, range) => {
                    edit.update_highlight(*p, *range)?;
                },
                NoteChange::RemoveHighlight(p) => {
                    edit.remove_highlight(*p)?;
                },
                NoteChange::AddComment(range, body) => {
                    edit.add_comment(*range, body)?;
                },
                NoteChange::UpdateComment(p, range, body) => {
                    edit.update_comment(*p, *range, body)?;
                },
                NoteChange::RemoveComment(p) => {
                    edit.remove_comment(*p)?;
                },
                NoteChange::AddReply(p, body) => {
                    edit.add_reply(*p, body)?;
                },
                NoteChange::UpdateReply(p, r, body) => {
                    edit.update_reply(*p, *r, body)?;
                },
                NoteChange::RemoveReply(p, r) => {
                    edit.remove_reply(*p, *r)?;
                },
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_drawable_annotations(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host annotations", host_notes)
        .read("leaf annotations", leaf_notes);
    if tier == Tier::B {
        case = case
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
            .leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn noted(change: NoteChange) -> TestResult<Fixture> {
    let input = authored_fixture()?;
    Ok(Fixture {
        path: "authored + host annotation",
        bytes: host_note(&input.bytes, &change)?,
    })
}

#[test]
fn pages_text_box_highlight_add_parity() -> TestResult {
    note_case(
        "pages_text_box_highlight_add_parity",
        Tier::B,
        &authored_fixture()?,
        NoteChange::AddHighlight(range(0, 5)?),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_highlight_update_parity() -> TestResult {
    note_case(
        "pages_text_box_highlight_update_parity",
        Tier::A,
        &noted(NoteChange::AddHighlight(range(0, 5)?))?,
        NoteChange::UpdateHighlight(0, range(2, 8)?),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_highlight_remove_parity() -> TestResult {
    note_case(
        "pages_text_box_highlight_remove_parity",
        Tier::B,
        &noted(NoteChange::AddHighlight(range(0, 5)?))?,
        NoteChange::RemoveHighlight(0),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_comment_add_parity() -> TestResult {
    note_case(
        "pages_text_box_comment_add_parity",
        Tier::B,
        &authored_fixture()?,
        NoteChange::AddComment(range(0, 5)?, "Check this"),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_comment_update_parity() -> TestResult {
    note_case(
        "pages_text_box_comment_update_parity",
        Tier::A,
        &noted(NoteChange::AddComment(range(0, 5)?, "Check this"))?,
        NoteChange::UpdateComment(0, range(2, 8)?, "Checked"),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_comment_remove_parity() -> TestResult {
    note_case(
        "pages_text_box_comment_remove_parity",
        Tier::B,
        &noted(NoteChange::AddComment(range(0, 5)?, "Check this"))?,
        NoteChange::RemoveComment(0),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_comment_reply_parity() -> TestResult {
    let input = noted(NoteChange::AddComment(range(0, 5)?, "Check this"))?;
    note_case(
        "pages_text_box_comment_reply_parity",
        Tier::B,
        &input,
        NoteChange::AddReply(0, "Done"),
    )?;
    let replied = Fixture {
        path: input.path,
        bytes: host_note(&input.bytes, &NoteChange::AddReply(0, "Done"))?,
    };
    note_case(
        "pages_text_box_comment_reply_parity",
        Tier::B,
        &replied,
        NoteChange::UpdateReply(0, 0, "Done twice"),
    )?;
    note_case(
        "pages_text_box_comment_reply_parity",
        Tier::B,
        &replied,
        NoteChange::RemoveReply(0, 0),
    )?;
    let _ = Evidence::Queued;
    Ok(())
}

#[test]
fn pages_text_box_annotation_read_parity() -> TestResult {
    let commented = noted(NoteChange::AddComment(range(0, 5)?, "Check this"))?;
    let replied = host_note(&commented.bytes, &NoteChange::AddReply(0, "Done"))?;
    let highlighted = host_note(&replied, &NoteChange::AddHighlight(range(6, 9)?))?;
    for bytes in [
        authored_fixture()?.bytes,
        commented.bytes.clone(),
        replied,
        highlighted,
    ] {
        assert_eq!(host_notes(&bytes)?, leaf_notes(&bytes)?);
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- number attachments ------------------------------------------------------

use litchi_iwa_common::comment::DrawableId;
use litchi_pages::number_attachment::{
    Target, TextNumberAttachmentKind, TextNumberAttachmentSettings,
};

fn attachment(kind: TextNumberAttachmentKind) -> TextNumberAttachmentSettings {
    TextNumberAttachmentSettings::new(kind)
}

fn host_box_attachments(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let d = DrawableId::try_from(host_text_box(&e, 0)?)?;
    Ok(format!(
        "{:#?}",
        e.text_box_number_attachments(d)?
            .into_iter()
            .map(|item| (item.position, item.settings))
            .collect::<Vec<_>>()
    ))
}

fn leaf_box_attachments(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:#?}",
        Package::from_bytes(bytes)?
            .number_attachments(Target::Drawable(DrawableSelector::text_box(0)))?
            .into_iter()
            .map(|item| (item.position(), item.settings().clone()))
            .collect::<Vec<_>>()
    ))
}

#[derive(Clone, Copy)]
enum AttachmentChange {
    Insert(usize),
    Update,
    Remove,
}

fn host_attachment_change(source: &[u8], change: AttachmentChange) -> TestResult<Vec<u8>> {
    let mut e = pages::Draft::from_bytes(source)?;
    let d = DrawableId::try_from(host_text_box(&e, 0)?)?;
    match change {
        AttachmentChange::Insert(at) => drop(e.insert_text_box_number_attachment(
            d,
            TextPosition::from_utf16_index(at)?,
            attachment(TextNumberAttachmentKind::PageNumber),
        )?),
        AttachmentChange::Update => {
            let id = e.text_box_number_attachments(d)?.first().ok_or("none")?.id;
            drop(e.update_text_box_number_attachment(
                d,
                id,
                attachment(TextNumberAttachmentKind::PageCount),
            )?);
        },
        AttachmentChange::Remove => {
            let id = e.text_box_number_attachments(d)?.first().ok_or("none")?.id;
            drop(e.remove_text_box_number_attachment(d, id)?);
        },
    }
    Ok(e.to_bytes()?)
}

fn numbered_box() -> TestResult<Fixture> {
    let input = authored_fixture()?;
    Ok(Fixture {
        path: "basic.pages + two add_text_box + host page number at 3",
        bytes: host_attachment_change(&input.bytes, AttachmentChange::Insert(3))?,
    })
}

fn attachment_case(name: &'static str, input: &Fixture, change: AttachmentChange) -> TestResult {
    Case::new(name, Format::Pages, Tier::A, Evidence::Queued, input)
        .host(|source| host_attachment_change(source, change))
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_number_attachments(Target::Drawable(DrawableSelector::text_box(0)))?;
            match change {
                AttachmentChange::Insert(at) => drop(edit.insert(
                    TextPosition::from_utf16_index(at)?,
                    attachment(TextNumberAttachmentKind::PageNumber),
                )?),
                AttachmentChange::Update => drop(edit.update(
                    litchi_core::Position::new(0),
                    attachment(TextNumberAttachmentKind::PageCount),
                )?),
                AttachmentChange::Remove => drop(edit.remove(litchi_core::Position::new(0))?),
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_number_attachments(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host text-box attachments", host_box_attachments)
        .read("leaf text-box attachments", leaf_box_attachments)
        .run()?;
    Ok(())
}

#[test]
fn pages_text_box_number_attachment_insert_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    attachment_case(
        "pages_text_box_number_attachment_insert_parity",
        &authored_fixture()?,
        AttachmentChange::Insert(3),
    )
}

#[test]
fn pages_text_box_number_attachment_update_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    attachment_case(
        "pages_text_box_number_attachment_update_parity",
        &numbered_box()?,
        AttachmentChange::Update,
    )
}

#[test]
fn pages_text_box_number_attachment_remove_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    attachment_case(
        "pages_text_box_number_attachment_remove_parity",
        &numbered_box()?,
        AttachmentChange::Remove,
    )
}

#[test]
fn pages_text_box_number_attachment_read_parity() -> TestResult {
    for input in [authored_fixture()?, numbered_box()?] {
        assert_eq!(
            host_box_attachments(&input.bytes)?,
            leaf_box_attachments(&input.bytes)?
        );
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

// ---- cost probes (record 0413 §5) ------------------------------------------

#[test]
#[ignore = "callgrind cost probe (record 0413)"]
fn text_box_hyperlink_add_cost_abba() -> TestResult {
    let change = LinkChange::Add(range(0, 5)?, "https://example.com/litchi");
    crate::pages_cost_probes::abba(
        &authored_fixture()?,
        |source| host_link(source, &change),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_hyperlinks(DrawableSelector::text_box(0))?;
            edit.add(
                range(0, 5)?,
                TextHyperlinkTarget::new("https://example.com/litchi")?,
            )?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0413)"]
fn text_box_comment_add_cost_abba() -> TestResult {
    let change = NoteChange::AddComment(range(0, 5)?, "Check this");
    crate::pages_cost_probes::abba(
        &authored_fixture()?,
        |source| host_note(source, &change),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_drawable_annotations(DrawableSelector::text_box(0))?;
            edit.add_comment(range(0, 5)?, "Check this")?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0413)"]
fn text_box_number_attachment_insert_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &authored_fixture()?,
        |source| host_attachment_change(source, AttachmentChange::Insert(3)),
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_number_attachments(Target::Drawable(DrawableSelector::text_box(0)))?;
            edit.insert(
                TextPosition::from_utf16_index(3)?,
                attachment(TextNumberAttachmentKind::PageNumber),
            )?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
