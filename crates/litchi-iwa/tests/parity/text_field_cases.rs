//! Date & Time field and plain-text replacement parity (wave W1-D, record
//! 0374): the host `text::Draft` date-time family and `replace_text`/
//! `set_text` against `litchi_iwa_text_wire::{smart_field, ranged}` and the
//! text-wire storage rewrite, published by [`text_glue`](crate::text_glue).
//!
//! New fields carry the host's random text-attribute UUID, masked in the
//! comparison (tier B); updates, removals and replacements are tier A.

use litchi_iwa::text::{Draft, TextDateTimeFieldId, TextStorageId};
use litchi_iwa_common::WireLimits;
use litchi_iwa_protos::tswp;
use litchi_iwa_text::date_time::{
    DisplayText, Format, FormatterStyle, Instant, LocaleIdentifier, Settings, UpdatePlan,
};
use litchi_iwa_text::position::{TextPosition, TextRange};
use litchi_iwa_text_wire::ranged::{self, Table};
use litchi_iwa_text_wire::smart_field::{self, TextAttributeUuid};
use prost::Message;

use crate::harness::{
    Case, Evidence, Fixture, Format as FileFormat, LeafRun, TestResult, Tier, fixture,
};
use crate::known_fields::{Decoded, KnownFields, decode_as, top_level_fields};
use crate::text_annotation_cases::text_storages;
use crate::text_glue::Workspace;
use crate::text_style_cases::FIXTURES;

const LIMITS: WireLimits = WireLimits::DEFAULT;
const LEAF_UUID: &str = "4c495443-4849-5731-442d-746578740002";

fn masked_date_time(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tswp::DateTimeSmartFieldArchive::decode(data)?;
    if let Some(field) = message.super_.as_mut()
        && field
            .text_attribute_uuid_string
            .as_deref()
            .is_some_and(|uuid| TextAttributeUuid::new(uuid).is_ok())
    {
        field.text_attribute_uuid_string = Some("<random uuid>".to_owned());
    }
    let mut decoded = decode_as::<tswp::DateTimeSmartFieldArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn masked(known: KnownFields) -> KnownFields {
    known.with(
        smart_field::DATE_TIME_MESSAGE_TYPE,
        "TSWP.DateTimeSmartFieldArchive (text-attribute UUID masked)",
        masked_date_time,
    )
}

#[derive(Clone, Debug)]
pub enum Operation {
    Add(TextRange, Settings),
    Insert(TextPosition, DisplayText, Settings),
    Update(u64, TextRange, Settings),
    Remove(u64),
    ReplaceText(std::ops::Range<usize>, String),
    SetText(String),
}

pub type Fields = Vec<(u64, TextRange, String)>;

pub fn host_fields(bytes: &[u8], storage: u64) -> TestResult<(Fields, String)> {
    let draft = Draft::from_bytes(bytes)?;
    let id = TextStorageId::try_from(storage)?;
    let fields = draft
        .text_date_time_fields(id)?
        .into_iter()
        .map(|field| {
            (
                field.id.object_id(),
                field.range,
                format!("{:?}", field.settings),
            )
        })
        .collect();
    Ok((fields, draft.storage(id)?.storage.text().to_owned()))
}

pub fn leaf_fields(bytes: &[u8], storage: u64) -> TestResult<(Fields, String)> {
    let workspace = Workspace::open(bytes)?;
    let located = workspace.storage(storage)?;
    let objects = workspace.messages(smart_field::DATE_TIME_MESSAGE_TYPE)?;
    let mut fields = Vec::new();
    for entry in ranged::entries(&located.payload, Table::SmartField, LIMITS)? {
        if let Some(object) = objects
            .iter()
            .find(|message| message.object == entry.object() && message.entry == located.entry)
        {
            fields.push((
                entry.object(),
                entry.range(),
                format!("{:?}", smart_field::date_time(&object.payload, LIMITS)?.1),
            ));
        }
    }
    let text = litchi_iwa_text_wire::from_bytes(
        &located.payload,
        litchi_iwa_text_wire::Limits::default(),
    )?;
    Ok((fields, text.text().to_owned()))
}

pub fn host_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut draft = Draft::from_bytes(source)?;
    let id = TextStorageId::try_from(storage)?;
    match operation.clone() {
        Operation::Add(range, settings) => {
            draft.add_text_date_time_field(id, range, settings)?;
        },
        Operation::Insert(position, display, settings) => {
            draft.insert_text_date_time_field(id, position, display, settings)?;
        },
        Operation::Update(object, range, settings) => {
            draft.update_text_date_time_field(
                id,
                TextDateTimeFieldId::from_object_id(object)?,
                range,
                settings,
            )?;
        },
        Operation::Remove(object) => {
            draft.remove_text_date_time_field(id, TextDateTimeFieldId::from_object_id(object)?)?;
        },
        Operation::ReplaceText(range, text) => draft.replace_text(id, range, &text)?,
        Operation::SetText(text) => draft.set_text(id, &text)?,
    }
    Ok(draft.to_bytes()?)
}

fn add_field(
    workspace: &mut Workspace,
    storage: u64,
    range: TextRange,
    settings: &Settings,
) -> TestResult {
    let located = workspace.storage(storage)?;
    let object = workspace.next_identifier()?;
    let rewrite = ranged::rewrite(
        &located.payload,
        Table::SmartField,
        &ranged::Change::Add { object, range },
        LIMITS,
    )?;
    let (added, removed) = (rewrite.added(), rewrite.removed());
    workspace.replace_with_references(&located, rewrite.into_payload(), added, removed)?;
    let payload =
        smart_field::encode_date_time(&TextAttributeUuid::new(LEAF_UUID)?, settings, LIMITS)?;
    workspace.insert_plain(
        &located.entry,
        object,
        smart_field::DATE_TIME_MESSAGE_TYPE,
        payload,
    )?;
    workspace.raise_watermark(object)
}

pub fn leaf_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    match operation {
        Operation::Add(range, settings) => add_field(&mut workspace, storage, *range, settings)?,
        Operation::Insert(position, display, settings) => {
            let start = usize::try_from(position.utf16_index())?;
            let end = start + display.as_str().encode_utf16().count();
            let located = workspace.storage(storage)?;
            workspace.replace_text(&located, start..start, display.as_str())?;
            add_field(
                &mut workspace,
                storage,
                TextRange::from_utf16_indexes(start, end)?,
                settings,
            )?;
        },
        Operation::Update(object, range, settings) => {
            let located = workspace.storage(storage)?;
            let rewrite = ranged::rewrite(
                &located.payload,
                Table::SmartField,
                &ranged::Change::Move {
                    object: *object,
                    range: *range,
                },
                LIMITS,
            )?;
            workspace.replace_with_references(&located, rewrite.into_payload(), None, None)?;
            let current = workspace.message(*object, smart_field::DATE_TIME_MESSAGE_TYPE)?;
            workspace.replace_object_message(
                &current.entry,
                *object,
                smart_field::DATE_TIME_MESSAGE_TYPE,
                smart_field::date_time_with_settings(&current.payload, settings, LIMITS)?,
            )?;
        },
        Operation::Remove(object) => {
            let located = workspace.storage(storage)?;
            let rewrite = ranged::rewrite(
                &located.payload,
                Table::SmartField,
                &ranged::Change::Remove { object: *object },
                LIMITS,
            )?;
            let removed = rewrite.removed();
            workspace.replace_with_references(&located, rewrite.into_payload(), None, removed)?;
            workspace.remove_object(&located.entry, *object)?;
            workspace.forget_object(&located.entry, *object)?;
        },
        Operation::ReplaceText(range, text) => {
            let located = workspace.storage(storage)?;
            workspace.replace_text(&located, range.clone(), text)?;
        },
        Operation::SetText(text) => {
            let located = workspace.storage(storage)?;
            let units = litchi_iwa_text_wire::from_bytes(
                &located.payload,
                litchi_iwa_text_wire::Limits::default(),
            )?
            .text()
            .encode_utf16()
            .count();
            workspace.replace_text(&located, 0..units, text)?;
        },
    }
    workspace.into_bytes()
}

fn run(
    name: &'static str,
    format: FileFormat,
    tier: Tier,
    input: &Fixture,
    storage: u64,
    operation: &Operation,
) -> TestResult {
    let mut case = Case::new(name, format, tier, Evidence::Queued, input)
        .host(|source| host_route(source, storage, operation))
        .leaf(|source| Ok(LeafRun::output(leaf_route(source, storage, operation)?)))
        .read("host fields", move |bytes| host_fields(bytes, storage))
        .read("leaf fields", move |bytes| leaf_fields(bytes, storage));
    if tier == Tier::B {
        case = case.known_fields(masked);
    }
    case.run()?;
    Ok(())
}

fn settings() -> TestResult<[Settings; 2]> {
    Ok([
        Settings::fixed(
            Format::new("yyyy-MM-dd")?,
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

fn created(input: &Fixture, storage: u64, range: TextRange) -> TestResult<(Fixture, u64)> {
    let bytes = host_route(
        &input.bytes,
        storage,
        &Operation::Add(range, settings()?[0].clone()),
    )?;
    let object = host_fields(&bytes, storage)?
        .0
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
fn text_date_time_add_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            for value in settings()? {
                run(
                    "text_date_time_add_parity",
                    format,
                    Tier::B,
                    &input,
                    storage,
                    &Operation::Add(range, value.clone()),
                )
                .map_err(|e| format!("{path} storage {storage}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 8, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_date_time_insert_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            for position in [TextPosition::ZERO, range.end()] {
                run(
                    "text_date_time_insert_parity",
                    format,
                    Tier::B,
                    &input,
                    storage,
                    &Operation::Insert(
                        position,
                        DisplayText::new("2026-10-26")?,
                        settings()?[0].clone(),
                    ),
                )
                .map_err(|e| format!("{path} storage {storage}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 8, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_date_time_update_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, wide) in text_storages(&input.bytes, 2)? {
            let (created, object) = created(&input, storage, range)?;
            let [first, second] = settings()?;
            for (target, value) in [(wide, first), (range, second.clone()), (wide, second)] {
                run(
                    "text_date_time_update_parity",
                    format,
                    Tier::A,
                    &created,
                    storage,
                    &Operation::Update(object, target, value),
                )
                .map_err(|e| format!("{path} storage {storage}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 12, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_date_time_remove_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            let (created, object) = created(&input, storage, range)?;
            run(
                "text_date_time_remove_parity",
                format,
                Tier::A,
                &created,
                storage,
                &Operation::Remove(object),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 1;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 4, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_date_time_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 4)? {
            let (created, _) = created(&input, storage, range)?;
            for bytes in [&input.bytes, &created.bytes] {
                assert_eq!(host_fields(bytes, storage)?, leaf_fields(bytes, storage)?);
                reads += 1;
            }
        }
    }
    assert!(reads >= 8, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

#[test]
fn text_replace_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            let end = usize::try_from(range.end().utf16_index())?;
            for operation in [
                Operation::ReplaceText(0..end, "Litchi ✓".to_owned()),
                Operation::ReplaceText(end..end, " inserted".to_owned()),
                Operation::SetText("Replaced text\nSecond paragraph".to_owned()),
            ] {
                run(
                    "text_replace_parity",
                    format,
                    Tier::A,
                    &input,
                    storage,
                    &operation,
                )
                .map_err(|e| format!("{path} storage {storage} {operation:?}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 12, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_storage_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        let draft = Draft::from_bytes(&input.bytes)?;
        let workspace = Workspace::open(&input.bytes)?;
        for storage in draft.storages()? {
            let id: u64 = storage.id.to_string().parse()?;
            let located = workspace.storage(id)?;
            let leaf = litchi_iwa_text_wire::from_bytes(
                &located.payload,
                litchi_iwa_text_wire::Limits::default(),
            )?;
            assert_eq!(storage.storage.text(), leaf.text(), "{path} storage {id}");
            assert_eq!(
                draft.storage(storage.id)?.storage.text(),
                leaf.text(),
                "{path} storage {id}"
            );
            reads += 1;
        }
    }
    assert!(reads >= 8, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
