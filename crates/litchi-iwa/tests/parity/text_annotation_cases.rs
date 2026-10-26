//! Ranged smart-field parity (wave W1-D, record 0374): hyperlinks and
//! bookmarks of the host `text::Draft` against
//! `litchi_iwa_text_wire::{ranged, smart_field}` published by
//! [`text_glue`](crate::text_glue).
//!
//! The host draws each new smart field's text-attribute UUID at random, so
//! creation cases compare the new object with that one string masked (its
//! shape is still checked) and are tier B; moves, setting changes and
//! removals start from a host-created field and are tier A.

use litchi_iwa::text::{Draft, TextStorageId};
use litchi_iwa_common::WireLimits;
use litchi_iwa_protos::tswp;
use litchi_iwa_text::bookmark::{
    Id as BookmarkId, Name as BookmarkName, Settings as BookmarkSettings, Visibility,
};
use litchi_iwa_text::hyperlink::{TextHyperlinkTarget, raw as hyperlink_raw};
use litchi_iwa_text::position::{TextPosition, TextRange};
use litchi_iwa_text_wire::ranged::{self, Table};
use litchi_iwa_text_wire::smart_field::{self, TextAttributeUuid};
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::{Decoded, KnownFields, decode_as, top_level_fields};
use crate::text_glue::Workspace;
use crate::text_style_cases::FIXTURES;

const LIMITS: WireLimits = WireLimits::DEFAULT;

/// A fixed UUID the leaf route stamps on the fields it creates.
const LEAF_UUID: &str = "4c495443-4849-5731-442d-746578740001";

fn uuid_shaped(value: &str) -> bool {
    TextAttributeUuid::new(value).is_ok()
}

fn masked_hyperlink(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tswp::HyperlinkFieldArchive::decode(data)?;
    if let Some(field) = message.super_.as_mut()
        && field
            .text_attribute_uuid_string
            .as_deref()
            .is_some_and(uuid_shaped)
    {
        field.text_attribute_uuid_string = Some("<random uuid>".to_owned());
    }
    let mut decoded = decode_as::<tswp::HyperlinkFieldArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn masked_bookmark(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tswp::BookmarkFieldArchive::decode(data)?;
    if let Some(field) = message.super_.as_mut()
        && field
            .text_attribute_uuid_string
            .as_deref()
            .is_some_and(uuid_shaped)
    {
        field.text_attribute_uuid_string = Some("<random uuid>".to_owned());
    }
    let mut decoded = decode_as::<tswp::BookmarkFieldArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn masked(known: KnownFields) -> KnownFields {
    known
        .with(
            smart_field::HYPERLINK_MESSAGE_TYPE,
            "TSWP.HyperlinkFieldArchive (text-attribute UUID masked)",
            masked_hyperlink,
        )
        .with(
            smart_field::BOOKMARK_MESSAGE_TYPE,
            "TSWP.BookmarkFieldArchive (text-attribute UUID masked)",
            masked_bookmark,
        )
}

/// One smart-field kind: its table, message type and settings value.
#[derive(Clone, Debug, PartialEq)]
pub enum Field {
    Hyperlink(TextHyperlinkTarget),
    Bookmark(BookmarkSettings),
}

impl Field {
    const fn table(&self) -> Table {
        match self {
            Self::Hyperlink(_) => Table::SmartField,
            Self::Bookmark(_) => Table::Bookmark,
        }
    }

    const fn message_type(&self) -> u32 {
        match self {
            Self::Hyperlink(_) => smart_field::HYPERLINK_MESSAGE_TYPE,
            Self::Bookmark(_) => smart_field::BOOKMARK_MESSAGE_TYPE,
        }
    }

    fn encode(&self, uuid: &TextAttributeUuid) -> TestResult<Vec<u8>> {
        Ok(match self {
            Self::Hyperlink(target) => smart_field::encode_hyperlink(uuid, target, LIMITS)?,
            Self::Bookmark(settings) => smart_field::encode_bookmark(uuid, settings, LIMITS)?,
        })
    }

    fn patch(&self, payload: &[u8]) -> TestResult<Vec<u8>> {
        Ok(match self {
            Self::Hyperlink(target) => smart_field::hyperlink_with_target(payload, target, LIMITS)?,
            Self::Bookmark(settings) => {
                smart_field::bookmark_with_settings(payload, settings, LIMITS)?
            },
        })
    }
}

/// What one case does to a storage.
#[derive(Clone, Debug)]
pub enum Operation {
    Add(TextRange, Field),
    Update(u64, TextRange, Field),
    Remove(u64, Field),
}

/// A read of every field of one kind: `(object, range, settings)`.
pub type Fields = Vec<(u64, TextRange, String)>;

pub fn host_fields(bytes: &[u8], storage: u64, bookmarks: bool) -> TestResult<Fields> {
    let draft = Draft::from_bytes(bytes)?;
    let id = TextStorageId::try_from(storage)?;
    Ok(if bookmarks {
        draft
            .text_bookmarks(id)?
            .into_iter()
            .map(|bookmark| {
                (
                    bookmark.id.get(),
                    bookmark.range,
                    format!("{:?}", bookmark.settings),
                )
            })
            .collect()
    } else {
        draft
            .text_hyperlinks(id)?
            .into_iter()
            .map(|link| {
                (
                    hyperlink_raw::object_id(link.id),
                    link.range,
                    format!("{:?}", link.target),
                )
            })
            .collect()
    })
}

pub fn leaf_fields(bytes: &[u8], storage: u64, bookmarks: bool) -> TestResult<Fields> {
    let workspace = Workspace::open(bytes)?;
    let located = workspace.storage(storage)?;
    let table = if bookmarks {
        Table::Bookmark
    } else {
        Table::SmartField
    };
    let mut fields = Vec::new();
    for entry in ranged::entries(&located.payload, table, LIMITS)? {
        let objects = workspace.messages(if bookmarks {
            smart_field::BOOKMARK_MESSAGE_TYPE
        } else {
            smart_field::HYPERLINK_MESSAGE_TYPE
        })?;
        let Some(object) = objects
            .iter()
            .find(|message| message.object == entry.object() && message.entry == located.entry)
        else {
            continue;
        };
        let settings = if bookmarks {
            format!("{:?}", smart_field::bookmark(&object.payload, LIMITS)?.1)
        } else {
            format!("{:?}", smart_field::hyperlink(&object.payload, LIMITS)?.1)
        };
        fields.push((entry.object(), entry.range(), settings));
    }
    Ok(fields)
}

pub fn host_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut draft = Draft::from_bytes(source)?;
    let id = TextStorageId::try_from(storage)?;
    match operation.clone() {
        Operation::Add(range, Field::Hyperlink(target)) => {
            draft.add_text_hyperlink(id, range, target)?;
        },
        Operation::Add(range, Field::Bookmark(settings)) => {
            draft.add_text_bookmark(id, range, settings)?;
        },
        Operation::Update(object, range, Field::Hyperlink(target)) => {
            draft.update_text_hyperlink(
                id,
                hyperlink_raw::from_object_id(object)?,
                range,
                target,
            )?;
        },
        Operation::Update(object, range, Field::Bookmark(settings)) => {
            draft.update_text_bookmark(id, BookmarkId::new(object)?, range, settings)?;
        },
        Operation::Remove(object, Field::Hyperlink(_)) => {
            draft.remove_text_hyperlink(id, hyperlink_raw::from_object_id(object)?)?;
        },
        Operation::Remove(object, Field::Bookmark(_)) => {
            draft.remove_text_bookmark(id, BookmarkId::new(object)?)?;
        },
    }
    Ok(draft.to_bytes()?)
}

pub fn leaf_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    let located = workspace.storage(storage)?;
    match operation {
        Operation::Add(range, field) => {
            let object = workspace.next_identifier()?;
            let rewrite = ranged::rewrite(
                &located.payload,
                field.table(),
                &ranged::Change::Add {
                    object,
                    range: *range,
                },
                LIMITS,
            )?;
            let (added, removed) = (rewrite.added(), rewrite.removed());
            workspace.replace_with_references(&located, rewrite.into_payload(), added, removed)?;
            let payload = field.encode(&TextAttributeUuid::new(LEAF_UUID)?)?;
            workspace.insert_plain(&located.entry, object, field.message_type(), payload)?;
            workspace.raise_watermark(object)?;
        },
        Operation::Update(object, range, field) => {
            let rewrite = ranged::rewrite(
                &located.payload,
                field.table(),
                &ranged::Change::Move {
                    object: *object,
                    range: *range,
                },
                LIMITS,
            )?;
            workspace.replace_with_references(&located, rewrite.into_payload(), None, None)?;
            let current = workspace.message(*object, field.message_type())?;
            workspace.replace_object_message(
                &current.entry,
                *object,
                field.message_type(),
                field.patch(&current.payload)?,
            )?;
        },
        Operation::Remove(object, field) => {
            let rewrite = ranged::rewrite(
                &located.payload,
                field.table(),
                &ranged::Change::Remove { object: *object },
                LIMITS,
            )?;
            let removed = rewrite.removed();
            workspace.replace_with_references(&located, rewrite.into_payload(), None, removed)?;
            workspace.remove_object(&located.entry, *object)?;
            workspace.forget_object(&located.entry, *object)?;
        },
    }
    workspace.into_bytes()
}

pub fn run(
    name: &'static str,
    format: Format,
    tier: Tier,
    input: &Fixture,
    storage: u64,
    operation: &Operation,
) -> TestResult {
    let bookmarks = matches!(
        operation,
        Operation::Add(_, Field::Bookmark(_))
            | Operation::Update(_, _, Field::Bookmark(_))
            | Operation::Remove(_, Field::Bookmark(_))
    );
    let mut case = Case::new(name, format, tier, Evidence::Queued, input)
        .host(|source| host_route(source, storage, operation))
        .leaf(|source| Ok(LeafRun::output(leaf_route(source, storage, operation)?)))
        .read("host fields", move |bytes| {
            host_fields(bytes, storage, bookmarks)
        })
        .read("leaf fields", move |bytes| {
            leaf_fields(bytes, storage, bookmarks)
        });
    if tier == Tier::B {
        case = case.known_fields(masked);
    }
    case.run()?;
    Ok(())
}

/// The first `count` UTF-16 code units of the storage that end on a scalar
/// boundary (empty storages have none).
pub fn leading_range(bytes: &[u8], storage: u64, count: usize) -> TestResult<Option<TextRange>> {
    let draft = Draft::from_bytes(bytes)?;
    let text = draft
        .storage(TextStorageId::try_from(storage)?)?
        .storage
        .text()
        .to_owned();
    let mut end = 0u32;
    for character in text.chars().take(count) {
        end += u32::try_from(character.len_utf16())?;
    }
    if end == 0 {
        return Ok(None);
    }
    Ok(Some(TextRange::new(
        TextPosition::from_utf16_code_units(0),
        TextPosition::from_utf16_code_units(end),
    )?))
}

/// Storages with at least four characters of text.
pub fn text_storages(bytes: &[u8], limit: usize) -> TestResult<Vec<(u64, TextRange, TextRange)>> {
    let draft = Draft::from_bytes(bytes)?;
    let mut found = Vec::new();
    for storage in draft.storages()? {
        let id: u64 = storage.id.to_string().parse()?;
        let chars: Vec<char> = storage.storage.text().chars().collect();
        if chars.len() < 4 || chars.iter().take(4).any(|c| *c == '\u{fffc}') {
            continue;
        }
        let (Some(first), Some(wide)) =
            (leading_range(bytes, id, 2)?, leading_range(bytes, id, 4)?)
        else {
            continue;
        };
        if host_fields(bytes, id, false).is_err() || host_fields(bytes, id, true).is_err() {
            continue;
        }
        found.push((id, first, wide));
        if found.len() == limit {
            break;
        }
    }
    Ok(found)
}

fn samples() -> TestResult<[Field; 2]> {
    Ok([
        Field::Hyperlink(TextHyperlinkTarget::new("https://example.com/litchi")?),
        Field::Bookmark(
            BookmarkSettings::new()
                .with_name(BookmarkName::new("Chapter")?)
                .with_visibility(Visibility::Hidden),
        ),
    ])
}

fn second(field: &Field) -> TestResult<Field> {
    Ok(match field {
        Field::Hyperlink(_) => Field::Hyperlink(TextHyperlinkTarget::new("mailto:a@example.com")?),
        Field::Bookmark(_) => Field::Bookmark(BookmarkSettings::new()),
    })
}

/// The host-created field of `field`'s kind on `storage`: the input of
/// the update and removal cases.
fn created(
    input: &Fixture,
    storage: u64,
    range: TextRange,
    field: &Field,
) -> TestResult<(Fixture, u64)> {
    let bytes = host_route(&input.bytes, storage, &Operation::Add(range, field.clone()))?;
    let bookmarks = matches!(field, Field::Bookmark(_));
    let object = host_fields(&bytes, storage, bookmarks)?
        .into_iter()
        .find(|(_, created, _)| *created == range)
        .ok_or("the host did not create the field")?
        .0;
    Ok((
        Fixture {
            path: input.path,
            bytes,
        },
        object,
    ))
}

#[test]
fn text_smart_field_add_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            for field in samples()? {
                run(
                    "text_smart_field_add_parity",
                    format,
                    Tier::B,
                    &input,
                    storage,
                    &Operation::Add(range, field.clone()),
                )
                .map_err(|e| format!("{path} storage {storage} {field:?}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 8, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_smart_field_update_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, wide) in text_storages(&input.bytes, 2)? {
            for field in samples()? {
                let (created, object) = created(&input, storage, range, &field)?;
                for (target, value) in [
                    (wide, field.clone()),
                    (range, second(&field)?),
                    (wide, second(&field)?),
                ] {
                    run(
                        "text_smart_field_update_parity",
                        format,
                        Tier::A,
                        &created,
                        storage,
                        &Operation::Update(object, target, value.clone()),
                    )
                    .map_err(|e| format!("{path} storage {storage} {value:?}: {e}"))?;
                    cases += 1;
                }
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 24, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_smart_field_remove_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            for field in samples()? {
                let (created, object) = created(&input, storage, range, &field)?;
                run(
                    "text_smart_field_remove_parity",
                    format,
                    Tier::A,
                    &created,
                    storage,
                    &Operation::Remove(object, field.clone()),
                )
                .map_err(|e| format!("{path} storage {storage} {field:?}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 8, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_smart_field_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 4)? {
            for field in samples()? {
                let bookmarks = matches!(field, Field::Bookmark(_));
                let (created, _) = created(&input, storage, range, &field)?;
                for bytes in [&input.bytes, &created.bytes] {
                    assert_eq!(
                        host_fields(bytes, storage, bookmarks)?,
                        leaf_fields(bytes, storage, bookmarks)?,
                        "{path} storage {storage}"
                    );
                    reads += 1;
                }
            }
        }
    }
    assert!(reads >= 16, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

#[test]
fn text_smart_field_refusal_parity() -> TestResult {
    // An occupied range and a split surrogate pair are refused by both.
    let input = fixture(FIXTURES[0].0)?;
    let (storage, range, wide) = text_storages(&input.bytes, 1)?
        .into_iter()
        .next()
        .ok_or("no text storage")?;
    for field in samples()? {
        let (created, _) = created(&input, storage, range, &field)?;
        let overlap = Operation::Add(wide, field.clone());
        assert!(host_route(&created.bytes, storage, &overlap).is_err());
        assert!(leaf_route(&created.bytes, storage, &overlap).is_err());
    }
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
