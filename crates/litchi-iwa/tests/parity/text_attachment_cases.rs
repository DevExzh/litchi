//! Number-attachment parity (wave W1-D, record 0374): the host
//! `text::Draft` number-attachment family against
//! `litchi_iwa_text_wire::number_attachment` and the storage text rewrite,
//! published by [`text_glue`](crate::text_glue). Attachments carry no
//! random value, so every case is tier A.

use litchi_iwa::text::{Draft, TextStorageId};
use litchi_iwa_common::WireLimits;
use litchi_iwa_text::number_attachment::{
    TextNumberAttachmentFormat, TextNumberAttachmentKind, TextNumberAttachmentSettings,
    TextNumberAttachmentText, raw,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::number_attachment;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::text_annotation_cases::text_storages;
use crate::text_glue::Workspace;
use crate::text_style_cases::FIXTURES;

const LIMITS: WireLimits = WireLimits::DEFAULT;

#[derive(Clone, Debug)]
pub enum Operation {
    Insert(TextPosition, TextNumberAttachmentSettings),
    Update(u64, TextNumberAttachmentSettings),
    Remove(u64),
}

pub type Attachments = (
    Vec<(u64, TextPosition, TextNumberAttachmentSettings)>,
    String,
);

pub fn host_attachments(bytes: &[u8], storage: u64) -> TestResult<Attachments> {
    let draft = Draft::from_bytes(bytes)?;
    let id = TextStorageId::try_from(storage)?;
    let attachments = draft
        .text_number_attachments(id)?
        .into_iter()
        .map(|attachment| {
            (
                raw::object_id(attachment.id),
                attachment.position,
                attachment.settings,
            )
        })
        .collect();
    Ok((attachments, draft.storage(id)?.storage.text().to_owned()))
}

pub fn leaf_attachments(bytes: &[u8], storage: u64) -> TestResult<Attachments> {
    let workspace = Workspace::open(bytes)?;
    let located = workspace.storage(storage)?;
    let objects = workspace.messages(number_attachment::MESSAGE_TYPE)?;
    let mut attachments = Vec::new();
    for anchor in number_attachment::anchors(&located.payload, LIMITS)? {
        if let Some(object) = objects
            .iter()
            .find(|message| message.object == anchor.object() && message.entry == located.entry)
        {
            attachments.push((
                anchor.object(),
                anchor.position(),
                number_attachment::settings(&object.payload, LIMITS)?,
            ));
        }
    }
    let text = litchi_iwa_text_wire::from_bytes(
        &located.payload,
        litchi_iwa_text_wire::Limits::default(),
    )?;
    Ok((attachments, text.text().to_owned()))
}

pub fn host_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut draft = Draft::from_bytes(source)?;
    let id = TextStorageId::try_from(storage)?;
    match operation.clone() {
        Operation::Insert(position, settings) => {
            draft.insert_text_number_attachment(id, position, settings)?;
        },
        Operation::Update(object, settings) => {
            draft.update_text_number_attachment(id, raw::from_object_id(object)?, settings)?;
        },
        Operation::Remove(object) => {
            draft.remove_text_number_attachment(id, raw::from_object_id(object)?)?;
        },
    }
    Ok(draft.to_bytes()?)
}

pub fn leaf_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    match operation {
        Operation::Insert(position, settings) => {
            let start = usize::try_from(position.utf16_index())?;
            let located = workspace.storage(storage)?;
            workspace.replace_text(&located, start..start, "\u{fffc}")?;
            let located = workspace.storage(storage)?;
            let object = workspace.next_identifier()?;
            let payload = number_attachment::anchor(&located.payload, *position, object, LIMITS)?;
            workspace.replace_with_references(&located, payload, Some(object), None)?;
            workspace.insert_plain(
                &located.entry,
                object,
                number_attachment::MESSAGE_TYPE,
                number_attachment::encode(settings, LIMITS)?,
            )?;
            workspace.raise_watermark(object)?;
        },
        Operation::Update(object, settings) => {
            let current = workspace.message(*object, number_attachment::MESSAGE_TYPE)?;
            workspace.replace_object_message(
                &current.entry,
                *object,
                number_attachment::MESSAGE_TYPE,
                number_attachment::with_settings(&current.payload, settings, LIMITS)?,
            )?;
        },
        Operation::Remove(object) => {
            let located = workspace.storage(storage)?;
            let anchor = number_attachment::anchors(&located.payload, LIMITS)?
                .into_iter()
                .find(|anchor| anchor.object() == *object)
                .ok_or("attachment not anchored")?;
            let start = usize::try_from(anchor.position().utf16_index())?;
            let orphans = workspace.replace_text_pruning(&located, start..start + 1, "")?;
            for orphan in orphans {
                if workspace
                    .messages(number_attachment::MESSAGE_TYPE)?
                    .iter()
                    .any(|message| message.object == orphan && message.entry == located.entry)
                {
                    workspace.remove_object(&located.entry, orphan)?;
                    workspace.forget_object(&located.entry, orphan)?;
                }
            }
        },
    }
    workspace.into_bytes()
}

fn run(
    name: &'static str,
    format: Format,
    input: &Fixture,
    storage: u64,
    operation: &Operation,
) -> TestResult {
    Case::new(name, format, Tier::A, Evidence::Queued, input)
        .host(|source| host_route(source, storage, operation))
        .leaf(|source| Ok(LeafRun::output(leaf_route(source, storage, operation)?)))
        .read("host attachments", move |bytes| {
            host_attachments(bytes, storage)
        })
        .read("leaf attachments", move |bytes| {
            leaf_attachments(bytes, storage)
        })
        .run()?;
    Ok(())
}

fn samples() -> TestResult<[TextNumberAttachmentSettings; 2]> {
    Ok([
        TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageNumber),
        TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageCount)
            .with_string_equivalent(TextNumberAttachmentText::new("12")?)
            .with_number_format(TextNumberAttachmentFormat::from_native_value(3))
            .with_string_value(TextNumberAttachmentText::new("XII")?),
    ])
}

fn inserted(input: &Fixture, storage: u64, position: TextPosition) -> TestResult<(Fixture, u64)> {
    let bytes = host_route(
        &input.bytes,
        storage,
        &Operation::Insert(position, samples()?[0].clone()),
    )?;
    let object = host_attachments(&bytes, storage)?
        .0
        .into_iter()
        .find(|(_, at, _)| *at == position)
        .ok_or("the host did not insert the attachment")?
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
fn text_number_attachment_insert_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            for (position, settings) in [TextPosition::ZERO, range.end()]
                .into_iter()
                .zip(samples()?)
            {
                run(
                    "text_number_attachment_insert_parity",
                    format,
                    &input,
                    storage,
                    &Operation::Insert(position, settings),
                )
                .map_err(|e| format!("{path} storage {storage}: {e}"))?;
                cases += 1;
            }
            let _ = Tier::A;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 8, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_number_attachment_update_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            let (base, object) = inserted(&input, storage, range.end())?;
            run(
                "text_number_attachment_update_parity",
                format,
                &base,
                storage,
                &Operation::Update(object, samples()?[1].clone()),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 1;
            let _ = Tier::A;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 4, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_number_attachment_remove_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 2)? {
            let (base, object) = inserted(&input, storage, range.end())?;
            run(
                "text_number_attachment_remove_parity",
                format,
                &base,
                storage,
                &Operation::Remove(object),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 1;
            let _ = Tier::A;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 4, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_number_attachment_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, range, _) in text_storages(&input.bytes, 4)? {
            let (base, _) = inserted(&input, storage, range.end())?;
            for bytes in [&input.bytes, &base.bytes] {
                assert_eq!(
                    host_attachments(bytes, storage)?,
                    leaf_attachments(bytes, storage)?
                );
                reads += 1;
            }
        }
    }
    assert!(reads >= 8, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
