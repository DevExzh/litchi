//! Pages body text parity cases (wave W1-B, record 0371): the host's
//! whole-body text verbs against the leaf section-text family
//! (`litchi_pages::section::text`).

use litchi_iwa::pages;
use litchi_pages::{Package, SectionSelector, TextPosition, TextSpan};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const BASIC: &str = "iwork/pages/basic.pages";
const SECTIONS: &str = "iwork/pages/body-sections-unicode.pages";

fn host_body_text(bytes: &[u8]) -> TestResult<String> {
    Ok(pages::Draft::from_bytes(bytes)?.body_text()?)
}

fn leaf_sections(bytes: &[u8]) -> TestResult<Vec<String>> {
    let package = Package::from_bytes(bytes)?;
    (0..package.sections().len())
        .map(|index| {
            Ok(package
                .section_text(SectionSelector::index(index))?
                .to_owned())
        })
        .collect()
}

fn leaf_commit(commit: &litchi_pages::section::text::Commit) -> TestResult<LeafRun> {
    let restored = commit
        .package()
        .apply_section_text(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

#[test]
fn pages_set_body_text_parity() -> TestResult {
    let input = fixture(BASIC)?;
    Case::new(
        "pages_set_body_text_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        editor.set_body_text("Replaced body — ünïcode")?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_body_text()?;
        edit.set("Replaced body — ünïcode")?;
        leaf_commit(&edit.commit()?)
    })
    .read("host body_text", host_body_text)
    .read("leaf section_text", leaf_sections)
    .run()?;
    Ok(())
}

#[test]
fn pages_clear_body_parity() -> TestResult {
    let input = fixture(BASIC)?;
    Case::new(
        "pages_clear_body_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        editor.clear_body()?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_body_text()?;
        edit.clear()?;
        leaf_commit(&edit.commit()?)
    })
    .read("host body_text", host_body_text)
    .read("leaf section_text", leaf_sections)
    .run()?;
    Ok(())
}

/// `replace_body_text` takes a body-absolute UTF-16 range that may not
/// consume a section break, so it always falls inside one section; the leaf
/// edits that section with the section-relative span.
#[test]
fn pages_replace_body_text_parity() -> TestResult {
    let input = fixture(SECTIONS)?;
    let body = host_body_text(&input.bytes)?;
    let units = body.encode_utf16().collect::<Vec<_>>();
    // The second section starts after the first section-break marker.
    let marker = units
        .iter()
        .position(|unit| *unit == 0x0004)
        .ok_or("fixture has one section")?;
    let section_start = marker + 1;
    let (start, end) = (section_start + 1, section_start + 3);
    Case::new(
        "pages_replace_body_text_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        editor.replace_body_text(start..end, "Ω inserted")?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_section_text(SectionSelector::index(1))?;
        edit.replace(
            TextSpan::new(
                TextPosition::from_utf16_index(start - section_start)?,
                TextPosition::from_utf16_index(end - section_start)?,
            )?,
            "Ω inserted",
        )?;
        leaf_commit(&edit.commit()?)
    })
    .read("host body_text", host_body_text)
    .read("leaf section_text", leaf_sections)
    .run()?;
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn text_set_body_text_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(BASIC)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.set_body_text("Replaced body — ünïcode")?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_text()?;
            edit.set("Replaced body — ünïcode")?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

/// `TSWP.BookmarkFieldArchive` with its random `text_attribute_uuid_string`
/// replaced by a fixed placeholder: both routes draw a fresh UUID for a new
/// bookmark, as Pages does, so only its presence and form are compared.
fn bookmark_without_uuid(data: &[u8]) -> Result<crate::known_fields::Decoded, prost::DecodeError> {
    use prost::Message as _;
    let mut bookmark = litchi_iwa_protos::tswp::BookmarkFieldArchive::decode(data)?;
    if let Some(smart) = bookmark.super_.as_mut()
        && let Some(uuid) = smart.text_attribute_uuid_string.as_mut()
    {
        assert_eq!(uuid.len(), 36, "text-attribute UUID form");
        *uuid = "00000000-0000-0000-0000-000000000000".to_owned();
    }
    crate::known_fields::decode_as::<litchi_iwa_protos::tswp::BookmarkFieldArchive>(
        &bookmark.encode_to_vec(),
    )
}

fn host_bookmarks(bytes: &[u8]) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    Ok(format!(
        "{:#?}",
        editor
            .body_bookmarks()?
            .into_iter()
            .map(|bookmark| (bookmark.range, bookmark.settings))
            .collect::<Vec<_>>()
    ))
}

fn leaf_bookmarks(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:#?}",
        Package::from_bytes(bytes)?
            .body_bookmarks()?
            .into_iter()
            .map(|bookmark| (bookmark.range(), bookmark.settings().clone()))
            .collect::<Vec<_>>()
    ))
}

fn bookmark_settings(name: &str) -> TestResult<litchi_pages::bookmark::Settings> {
    Ok(litchi_pages::bookmark::Settings::new().with_name(litchi_pages::bookmark::Name::new(name)?))
}

/// `basic.pages` with one host-authored bookmark "Intro" over UTF-16 0..4.
fn bookmarked_fixture() -> TestResult<crate::harness::Fixture> {
    let input = fixture(BASIC)?;
    let mut editor = pages::Draft::from_bytes(&input.bytes)?;
    editor.add_body_bookmark(
        litchi_pages::bookmark::TextRange::from_utf16_indexes(0, 4)?,
        bookmark_settings("Intro")?,
    )?;
    Ok(crate::harness::Fixture {
        path: "iwork/pages/basic.pages + host bookmark Intro 0..4",
        bytes: editor.to_bytes()?,
    })
}

enum BookmarkOp {
    Add(usize, usize, &'static str),
    Update(usize, usize, &'static str),
    Remove,
}

fn bookmark_case(name: &'static str, tier: Tier, evidence: Evidence, op: BookmarkOp) -> TestResult {
    let input = match op {
        BookmarkOp::Add(..) => fixture(BASIC)?,
        _ => bookmarked_fixture()?,
    };
    let op = &op;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .known_fields(|known| known.with(2_035, "TSWP.BookmarkFieldArchive", bookmark_without_uuid))
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            match *op {
                BookmarkOp::Add(start, end, label) => drop(editor.add_body_bookmark(
                    litchi_pages::bookmark::TextRange::from_utf16_indexes(start, end)?,
                    bookmark_settings(label)?,
                )?),
                BookmarkOp::Update(start, end, label) => {
                    let id = editor.body_bookmarks()?.first().ok_or("no bookmark")?.id;
                    drop(editor.update_body_bookmark(
                        id,
                        litchi_pages::bookmark::TextRange::from_utf16_indexes(start, end)?,
                        bookmark_settings(label)?,
                    )?);
                },
                BookmarkOp::Remove => {
                    let id = editor.body_bookmarks()?.first().ok_or("no bookmark")?.id;
                    drop(editor.remove_body_bookmark(id)?);
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_bookmarks()?;
            match *op {
                BookmarkOp::Add(start, end, label) => drop(edit.add(
                    litchi_pages::bookmark::TextRange::from_utf16_indexes(start, end)?,
                    bookmark_settings(label)?,
                )?),
                BookmarkOp::Update(start, end, label) => drop(edit.update(
                    0,
                    litchi_pages::bookmark::TextRange::from_utf16_indexes(start, end)?,
                    bookmark_settings(label)?,
                )?),
                BookmarkOp::Remove => drop(edit.remove("Intro")?),
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_bookmarks(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host body_bookmarks", host_bookmarks)
        .read("leaf body_bookmarks", leaf_bookmarks)
        .read("leaf text", |bytes| Ok(Package::from_bytes(bytes)?.text()?))
        .run()?;
    Ok(())
}

#[test]
fn pages_body_bookmark_add_parity() -> TestResult {
    bookmark_case(
        "pages_body_bookmark_add_parity",
        Tier::B,
        Evidence::Queued,
        BookmarkOp::Add(2, 7, "Methods"),
    )
}

#[test]
fn pages_body_bookmark_update_parity() -> TestResult {
    bookmark_case(
        "pages_body_bookmark_update_parity",
        Tier::A,
        Evidence::Queued,
        BookmarkOp::Update(1, 6, "Renamed"),
    )
}

#[test]
fn pages_body_bookmark_remove_parity() -> TestResult {
    bookmark_case(
        "pages_body_bookmark_remove_parity",
        Tier::A,
        Evidence::Queued,
        BookmarkOp::Remove,
    )
}

/// `TSWP.DateTimeSmartFieldArchive` with its random text-attribute UUID
/// replaced by a fixed placeholder (see [`bookmark_without_uuid`]).
fn date_time_without_uuid(data: &[u8]) -> Result<crate::known_fields::Decoded, prost::DecodeError> {
    use prost::Message as _;
    let mut field = litchi_iwa_protos::tswp::DateTimeSmartFieldArchive::decode(data)?;
    if let Some(smart) = field.super_.as_mut()
        && let Some(uuid) = smart.text_attribute_uuid_string.as_mut()
    {
        assert_eq!(uuid.len(), 36, "text-attribute UUID form");
        *uuid = "00000000-0000-0000-0000-000000000000".to_owned();
    }
    crate::known_fields::decode_as::<litchi_iwa_protos::tswp::DateTimeSmartFieldArchive>(
        &field.encode_to_vec(),
    )
}

fn date_settings(seconds: f64) -> TestResult<litchi_pages::date_time::Settings> {
    use litchi_pages::date_time::{Format, Instant, LocaleIdentifier, Settings};
    Ok(Settings::fixed(
        Format::new("yyyy-MM-dd")?,
        LocaleIdentifier::new("en_US")?,
        Instant::from_reference_date_seconds(seconds)?,
    ))
}

fn host_fields(bytes: &[u8]) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    Ok(format!(
        "{:#?}",
        editor
            .body_date_time_fields()?
            .into_iter()
            .map(|field| (field.range, field.settings))
            .collect::<Vec<_>>()
    ))
}

fn leaf_fields(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:#?}",
        Package::from_bytes(bytes)?
            .body_date_time_fields()?
            .into_iter()
            .map(|field| (field.range(), field.settings().clone()))
            .collect::<Vec<_>>()
    ))
}

/// `basic.pages` with one host-authored date field over UTF-16 0..4.
fn dated_fixture() -> TestResult<crate::harness::Fixture> {
    let input = fixture(BASIC)?;
    let mut editor = pages::Draft::from_bytes(&input.bytes)?;
    editor.add_body_date_time_field(
        litchi_pages::date_time::TextRange::from_utf16_indexes(0, 4)?,
        date_settings(800_000_000.0)?,
    )?;
    Ok(crate::harness::Fixture {
        path: "iwork/pages/basic.pages + host date field 0..4",
        bytes: editor.to_bytes()?,
    })
}

enum FieldOp {
    Add(usize, usize),
    Insert(usize, &'static str),
    Update(usize, usize),
    Remove,
}

fn date_time_case(name: &'static str, tier: Tier, evidence: Evidence, op: FieldOp) -> TestResult {
    use litchi_pages::date_time::{DisplayText, TextRange};
    let input = match op {
        FieldOp::Add(..) | FieldOp::Insert(..) => fixture(BASIC)?,
        _ => dated_fixture()?,
    };
    let op = &op;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .known_fields(|known| {
            known.with(
                2_034,
                "TSWP.DateTimeSmartFieldArchive",
                date_time_without_uuid,
            )
        })
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            match *op {
                FieldOp::Add(start, end) => drop(editor.add_body_date_time_field(
                    TextRange::from_utf16_indexes(start, end)?,
                    date_settings(820_000_000.0)?,
                )?),
                FieldOp::Insert(at, text) => drop(editor.insert_body_date_time_field(
                    litchi_pages::TextPosition::from_utf16_index(at)?,
                    DisplayText::new(text)?,
                    date_settings(820_000_000.0)?,
                )?),
                FieldOp::Update(start, end) => {
                    let id = editor
                        .body_date_time_fields()?
                        .first()
                        .ok_or("no field")?
                        .id;
                    drop(editor.update_body_date_time_field(
                        id,
                        TextRange::from_utf16_indexes(start, end)?,
                        date_settings(830_000_000.0)?,
                    )?);
                },
                FieldOp::Remove => {
                    let id = editor
                        .body_date_time_fields()?
                        .first()
                        .ok_or("no field")?
                        .id;
                    drop(editor.remove_body_date_time_field(id)?);
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_date_time_fields()?;
            match *op {
                FieldOp::Add(start, end) => drop(edit.add(
                    TextRange::from_utf16_indexes(start, end)?,
                    date_settings(820_000_000.0)?,
                )?),
                FieldOp::Insert(at, text) => drop(edit.insert(
                    litchi_pages::TextPosition::from_utf16_index(at)?,
                    DisplayText::new(text)?,
                    date_settings(820_000_000.0)?,
                )?),
                FieldOp::Update(start, end) => drop(edit.update(
                    litchi_core::Position::new(0),
                    TextRange::from_utf16_indexes(start, end)?,
                    date_settings(830_000_000.0)?,
                )?),
                FieldOp::Remove => drop(edit.remove(litchi_core::Position::new(0))?),
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_date_time_fields(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        })
        .read("host body_date_time_fields", host_fields)
        .read("leaf body_date_time_fields", leaf_fields)
        .read("leaf text", |bytes| Ok(Package::from_bytes(bytes)?.text()?))
        .run()?;
    Ok(())
}

#[test]
fn pages_body_date_time_add_parity() -> TestResult {
    date_time_case(
        "pages_body_date_time_add_parity",
        Tier::B,
        Evidence::Queued,
        FieldOp::Add(2, 7),
    )
}

#[test]
fn pages_body_date_time_insert_parity() -> TestResult {
    date_time_case(
        "pages_body_date_time_insert_parity",
        Tier::B,
        Evidence::Queued,
        FieldOp::Insert(3, "2026-10-26"),
    )
}

#[test]
fn pages_body_date_time_update_parity() -> TestResult {
    date_time_case(
        "pages_body_date_time_update_parity",
        Tier::A,
        Evidence::Queued,
        FieldOp::Update(1, 6),
    )
}

#[test]
fn pages_body_date_time_remove_parity() -> TestResult {
    date_time_case(
        "pages_body_date_time_remove_parity",
        Tier::A,
        Evidence::Queued,
        FieldOp::Remove,
    )
}

fn page_number() -> litchi_pages::number_attachment::TextNumberAttachmentSettings {
    litchi_pages::number_attachment::TextNumberAttachmentSettings::new(
        litchi_pages::number_attachment::TextNumberAttachmentKind::PageNumber,
    )
}

fn page_count() -> litchi_pages::number_attachment::TextNumberAttachmentSettings {
    litchi_pages::number_attachment::TextNumberAttachmentSettings::new(
        litchi_pages::number_attachment::TextNumberAttachmentKind::PageCount,
    )
}

fn host_attachments(bytes: &[u8]) -> TestResult<String> {
    let editor = pages::Draft::from_bytes(bytes)?;
    Ok(format!(
        "{:#?}",
        editor
            .body_number_attachments()?
            .into_iter()
            .map(|attachment| (attachment.position, attachment.settings))
            .collect::<Vec<_>>()
    ))
}

fn leaf_attachments(bytes: &[u8]) -> TestResult<String> {
    Ok(format!(
        "{:#?}",
        Package::from_bytes(bytes)?
            .number_attachments(litchi_pages::number_attachment::Target::Body)?
            .into_iter()
            .map(|attachment| (attachment.position(), attachment.settings().clone()))
            .collect::<Vec<_>>()
    ))
}

/// `basic.pages` with one host-authored page number at UTF-16 2.
fn numbered_fixture() -> TestResult<crate::harness::Fixture> {
    let input = fixture(BASIC)?;
    let mut editor = pages::Draft::from_bytes(&input.bytes)?;
    editor.insert_body_number_attachment(
        litchi_pages::TextPosition::from_utf16_index(2)?,
        page_number(),
    )?;
    Ok(crate::harness::Fixture {
        path: "iwork/pages/basic.pages + host page number at 2",
        bytes: editor.to_bytes()?,
    })
}

enum AttachmentOp {
    Insert(usize),
    Update,
    Remove,
}

fn attachment_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    op: AttachmentOp,
) -> TestResult {
    use litchi_pages::number_attachment::Target;
    let input = match op {
        AttachmentOp::Insert(_) => fixture(BASIC)?,
        _ => numbered_fixture()?,
    };
    let op = &op;
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            match *op {
                AttachmentOp::Insert(at) => drop(editor.insert_body_number_attachment(
                    litchi_pages::TextPosition::from_utf16_index(at)?,
                    page_count(),
                )?),
                AttachmentOp::Update => {
                    let id = editor.body_number_attachments()?.first().ok_or("none")?.id;
                    drop(editor.update_body_number_attachment(id, page_count())?);
                },
                AttachmentOp::Remove => {
                    let id = editor.body_number_attachments()?.first().ok_or("none")?.id;
                    drop(editor.remove_body_number_attachment(id)?);
                },
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_number_attachments(Target::Body)?;
            match *op {
                AttachmentOp::Insert(at) => drop(edit.insert(
                    litchi_pages::TextPosition::from_utf16_index(at)?,
                    page_count(),
                )?),
                AttachmentOp::Update => {
                    edit.update(litchi_core::Position::new(0), page_count())?;
                },
                AttachmentOp::Remove => {
                    edit.remove(litchi_core::Position::new(0))?;
                },
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
        .read("host body_number_attachments", host_attachments)
        .read("leaf number_attachments", leaf_attachments)
        .read("leaf text", |bytes| Ok(Package::from_bytes(bytes)?.text()?))
        .run()?;
    Ok(())
}

#[test]
fn pages_body_number_attachment_insert_parity() -> TestResult {
    attachment_case(
        "pages_body_number_attachment_insert_parity",
        Tier::A,
        Evidence::Queued,
        AttachmentOp::Insert(3),
    )
}

#[test]
fn pages_body_number_attachment_update_parity() -> TestResult {
    attachment_case(
        "pages_body_number_attachment_update_parity",
        Tier::A,
        Evidence::Queued,
        AttachmentOp::Update,
    )
}

/// Removal deletes the U+FFFC, its table entry, the storage's header
/// reference and the attachment object with its metadata on both routes
/// (the host's pruning gap was fixed in record 0371).
#[test]
fn pages_body_number_attachment_remove_parity() -> TestResult {
    attachment_case(
        "pages_body_number_attachment_remove_parity",
        Tier::A,
        Evidence::Queued,
        AttachmentOp::Remove,
    )
}

fn first_header() -> litchi_pages::HeaderFooterSelector<'static> {
    litchi_pages::HeaderFooterSelector::index(
        0,
        litchi_pages::Template::First,
        litchi_pages::Kind::Header,
        0,
    )
}

fn header_attachment_case(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    op: AttachmentOp,
) -> TestResult {
    use litchi_pages::number_attachment::Target;
    let mut input = fixture(BASIC)?;
    if !matches!(op, AttachmentOp::Insert(_)) {
        let mut editor = pages::Draft::from_bytes(&input.bytes)?;
        editor.insert_header_footer_number_attachment(
            first_header(),
            litchi_pages::TextPosition::from_utf16_index(0)?,
            page_number(),
        )?;
        input = crate::harness::Fixture {
            path: "iwork/pages/basic.pages + host page number in the first header",
            bytes: editor.to_bytes()?,
        };
    }
    let host_read = |bytes: &[u8]| -> TestResult<String> {
        let editor = pages::Draft::from_bytes(bytes)?;
        Ok(format!(
            "{:#?}",
            editor
                .header_footer_number_attachments(first_header())?
                .into_iter()
                .map(|attachment| (attachment.position, attachment.settings))
                .collect::<Vec<_>>()
        ))
    };
    let leaf_read = |bytes: &[u8]| -> TestResult<String> {
        Ok(format!(
            "{:#?}",
            Package::from_bytes(bytes)?
                .number_attachments(Target::HeaderFooter(first_header()))?
                .into_iter()
                .map(|attachment| (attachment.position(), attachment.settings().clone()))
                .collect::<Vec<_>>()
        ))
    };
    Case::new(name, Format::Pages, tier, evidence, &input)
        .host(|source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            if !matches!(op, AttachmentOp::Insert(_)) {
                let id = editor
                    .header_footer_number_attachments(first_header())?
                    .first()
                    .ok_or("none")?
                    .id;
                if matches!(op, AttachmentOp::Remove) {
                    drop(editor.remove_header_footer_number_attachment(first_header(), id)?);
                } else {
                    drop(editor.update_header_footer_number_attachment(
                        first_header(),
                        id,
                        page_count(),
                    )?);
                }
            } else {
                drop(editor.insert_header_footer_number_attachment(
                    first_header(),
                    litchi_pages::TextPosition::from_utf16_index(0)?,
                    page_count(),
                )?);
            }
            Ok(editor.to_bytes()?)
        })
        .leaf(|source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_number_attachments(Target::HeaderFooter(first_header()))?;
            if matches!(op, AttachmentOp::Remove) {
                edit.remove(litchi_core::Position::new(0))?;
            } else if matches!(op, AttachmentOp::Update) {
                edit.update(litchi_core::Position::new(0), page_count())?;
            } else {
                edit.insert(
                    litchi_pages::TextPosition::from_utf16_index(0)?,
                    page_count(),
                )?;
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
        .read("host header_footer_number_attachments", host_read)
        .read("leaf number_attachments", leaf_read)
        .run()?;
    Ok(())
}

#[test]
fn pages_header_number_attachment_insert_parity() -> TestResult {
    header_attachment_case(
        "pages_header_number_attachment_insert_parity",
        Tier::A,
        Evidence::Queued,
        AttachmentOp::Insert(0),
    )
}

#[test]
fn pages_header_number_attachment_update_parity() -> TestResult {
    header_attachment_case(
        "pages_header_number_attachment_update_parity",
        Tier::A,
        Evidence::Queued,
        AttachmentOp::Update,
    )
}

#[test]
fn pages_header_number_attachment_remove_parity() -> TestResult {
    header_attachment_case(
        "pages_header_number_attachment_remove_parity",
        Tier::A,
        Evidence::Queued,
        AttachmentOp::Remove,
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn text_bookmark_add_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(BASIC)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.add_body_bookmark(
                litchi_pages::bookmark::TextRange::from_utf16_indexes(2, 7)?,
                bookmark_settings("Methods")?,
            )?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_bookmarks()?;
            edit.add(
                litchi_pages::bookmark::TextRange::from_utf16_indexes(2, 7)?,
                bookmark_settings("Methods")?,
            )?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn text_date_time_insert_cost_abba() -> TestResult {
    use litchi_pages::date_time::DisplayText;
    crate::pages_cost_probes::abba(
        &fixture(BASIC)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.insert_body_date_time_field(
                litchi_pages::TextPosition::from_utf16_index(3)?,
                DisplayText::new("2026-10-26")?,
                date_settings(820_000_000.0)?,
            )?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_date_time_fields()?;
            edit.insert(
                litchi_pages::TextPosition::from_utf16_index(3)?,
                DisplayText::new("2026-10-26")?,
                date_settings(820_000_000.0)?,
            )?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn text_number_attachment_insert_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(BASIC)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.insert_body_number_attachment(
                litchi_pages::TextPosition::from_utf16_index(3)?,
                page_count(),
            )?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_number_attachments(litchi_pages::number_attachment::Target::Body)?;
            edit.insert(
                litchi_pages::TextPosition::from_utf16_index(3)?,
                page_count(),
            )?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn text_number_attachment_remove_cost_abba() -> TestResult {
    let mut editor = pages::Draft::from_bytes(&fixture(BASIC)?.bytes)?;
    editor.insert_body_number_attachment(
        litchi_pages::TextPosition::from_utf16_index(3)?,
        page_count(),
    )?;
    let input = crate::harness::Fixture {
        path: "iwork/pages/basic.pages + host page count at 3",
        bytes: editor.to_bytes()?,
    };
    crate::pages_cost_probes::abba(
        &input,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            let id = editor.body_number_attachments()?.first().ok_or("none")?.id;
            drop(editor.remove_body_number_attachment(id)?);
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit =
                package.edit_number_attachments(litchi_pages::number_attachment::Target::Body)?;
            edit.remove(litchi_core::Position::new(0))?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}

/// `body_storage` and `body_text` against the leaf reads of the same names,
/// on native fixtures with sections, tables, media and footnote markers.
#[test]
fn pages_body_storage_read_parity() -> TestResult {
    let input = fixture(BASIC)?;
    Case::new(
        "pages_body_storage_read_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(pages::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host body text", |bytes: &[u8]| {
        Ok(pages::Draft::from_bytes(bytes)?.body_text()?)
    })
    .read("leaf body_storage text", |bytes: &[u8]| {
        Ok(Package::from_bytes(bytes)?.body_storage()?.into_text())
    })
    .run()?;
    for path in [
        BASIC,
        "iwork/pages/body-sections-unicode.pages",
        "iwork/pages/body-table-catalog-native.pages",
        "iwork/pages/image-adjustments-native.pages",
        "iwork/pages/chart-data-native.pages",
    ] {
        let document = fixture(path)?;
        let host = pages::Draft::from_bytes(&document.bytes)?;
        let leaf = Package::from_bytes(&document.bytes)?;
        assert_eq!(host.body_text()?, leaf.body_storage()?.text(), "{path}");
        assert_eq!(host.body_storage()?.storage, leaf.body_storage()?, "{path}");
    }
    Ok(())
}
