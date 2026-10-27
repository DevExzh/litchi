//! Document language, locale and revision identifier of a fresh document
//! (record 0371; the migration host builder's `language`/`locale`).
//!
//! The host builder writes the language into the document archive, the
//! body storage's language table and the document settings, and the locale
//! into the document archive (its current and creation locale) and the
//! settings (creation and last locale). This rewrites exactly those fields
//! on the wire.

use litchi_iwa_common::wire::{
    patch_length_delimited_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};

use super::Package;
use super::body_drawable_graph::MessageRewrite;
use crate::PackageError;

const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const METADATA_COMPONENT: &str = "Index/Metadata.iwa";

/// The located registry message.
struct Located<'a> {
    component: std::sync::Arc<str>,
    object: u64,
    message_index: usize,
    payload: &'a [u8],
}
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const SETTINGS_MESSAGE_TYPE: u32 = 10_012;
/// `TP.DocumentArchive.super` (`TSA.DocumentArchive`).
const DOCUMENT_SUPER: u32 = 15;
/// `TSA.DocumentArchive.super` (`TSK.DocumentArchive`).
const TSA_SUPER: u32 = 1;
const TSA_DOCUMENT_LANGUAGE: u32 = 3;
const TSK_LOCALE: u32 = 4;
const TSK_CREATION_LOCALE: u32 = 9;
const STORAGE_TABLE_LANGUAGE: u32 = 19;
const ATTRIBUTE_TABLE_ENTRIES: u32 = 1;
const ATTRIBUTE_OBJECT: u32 = 2;
const SETTINGS_LANGUAGE: u32 = 21;
const SETTINGS_CREATION_LOCALE: u32 = 23;
const SETTINGS_LAST_LOCALE: u32 = 24;

type WireResult<T> = Result<T, litchi_iwa_common::Error>;

fn invalid(error: impl std::fmt::Debug) -> PackageError {
    PackageError::InvalidFormat(format!(
        "Pages language or locale rewrite failed: {error:?}"
    ))
}

/// The singular length-delimited field `number`, if present.
fn single(data: &[u8], number: u32) -> WireResult<Option<&[u8]>> {
    Ok(repeated_length_delimited_payloads(data, number)?
        .last()
        .copied())
}

/// Set the string field `number` to `value`, adding it when absent.
fn set(data: &[u8], number: u32, value: &str) -> WireResult<Vec<u8>> {
    let present = single(data, number)?.is_some();
    patch_length_delimited_field(data, number, present, Some(value.as_bytes()))
}

/// Set `number` only when it is present.
fn set_present(data: &[u8], number: u32, value: &str) -> WireResult<Vec<u8>> {
    if single(data, number)?.is_none() {
        return Ok(data.to_vec());
    }
    patch_length_delimited_field(data, number, true, Some(value.as_bytes()))
}

/// Replace the nested message `number` of `data` with `f(nested)`.
fn nested(
    data: &[u8],
    number: u32,
    f: impl FnOnce(&[u8]) -> WireResult<Vec<u8>>,
) -> WireResult<Vec<u8>> {
    let Some(inner) = single(data, number)? else {
        return Ok(data.to_vec());
    };
    let rewritten = f(inner)?;
    patch_length_delimited_field(data, number, true, Some(&rewritten))
}

fn document(payload: &[u8], language: Option<&str>, locale: Option<&str>) -> WireResult<Vec<u8>> {
    nested(payload, DOCUMENT_SUPER, |tsa| {
        let tsa = match language {
            Some(language) => set(tsa, TSA_DOCUMENT_LANGUAGE, language)?,
            None => tsa.to_vec(),
        };
        match locale {
            Some(locale) => nested(&tsa, TSA_SUPER, |tsk| {
                let tsk = set(tsk, TSK_LOCALE, locale)?;
                set(&tsk, TSK_CREATION_LOCALE, locale)
            }),
            None => Ok(tsa),
        }
    })
}

fn storage(payload: &[u8], language: &str) -> WireResult<Vec<u8>> {
    nested(payload, STORAGE_TABLE_LANGUAGE, |table| {
        let entries = repeated_length_delimited_payloads(table, ATTRIBUTE_TABLE_ENTRIES)?
            .into_iter()
            .map(|entry| set_present(entry, ATTRIBUTE_OBJECT, language))
            .collect::<WireResult<Vec<_>>>()?;
        rewrite_repeated_length_delimited_fields(table, ATTRIBUTE_TABLE_ENTRIES, &entries)
    })
}

fn settings(payload: &[u8], language: Option<&str>, locale: Option<&str>) -> WireResult<Vec<u8>> {
    let mut data = payload.to_vec();
    if let Some(language) = language {
        data = set(&data, SETTINGS_LANGUAGE, language)?;
    }
    if let Some(locale) = locale {
        data = set(&data, SETTINGS_CREATION_LOCALE, locale)?;
        data = set(&data, SETTINGS_LAST_LOCALE, locale)?;
    }
    Ok(data)
}

impl Package {
    /// The metadata rewrite naming `version` as the package revision
    /// identifier (`TSP.PackageMetadata.revision.identifier`), if it changes.
    pub(crate) fn revision_rewrite(
        &self,
        version: &str,
    ) -> Result<Option<MessageRewrite>, PackageError> {
        // The registry message, found without decoding the registry.
        let component = self
            .state
            .source
            .components()
            .get(METADATA_COMPONENT)
            .ok_or_else(|| invalid("the package has no metadata"))?;
        let mut found = None;
        for object in &component.archive().objects {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == METADATA_MESSAGE_TYPE {
                    if found.is_some() {
                        return Err(invalid("the package metadata repeats"));
                    }
                    found = Some((object.archive_info.identifier, message_index, &message.data));
                }
            }
        }
        let (identifier, message_index, payload) =
            found.ok_or_else(|| invalid("the package has no metadata"))?;
        let metadata = Located {
            component: std::sync::Arc::from(component.name()),
            object: identifier.ok_or_else(|| invalid("the metadata object has no identifier"))?,
            message_index,
            payload,
        };
        let present = single(metadata.payload, 2)
            .map_err(invalid)?
            .is_some_and(|revision| single(revision, 2).ok().flatten().is_some());
        let data = litchi_iwa_common::wire::patch_nested_length_delimited_field(
            metadata.payload,
            &[2, 2],
            present,
            Some(version.as_bytes()),
        )
        .map_err(invalid)?;
        if data == metadata.payload {
            return Ok(None);
        }
        let object = std::num::NonZeroU64::new(metadata.object)
            .ok_or_else(|| invalid("the metadata object has no identifier"))?;
        Ok(Some(MessageRewrite {
            component: std::sync::Arc::clone(&metadata.component),
            object,
            message_index: metadata.message_index,
            message_type: METADATA_MESSAGE_TYPE,
            data,
        }))
    }

    /// Add (or compose into an existing rewrite of the body storage) the
    /// rewrite replacing the whole body text by `text`.
    pub(crate) fn body_text_rewrite(
        &self,
        rewrites: &mut Vec<MessageRewrite>,
        text: &str,
    ) -> Result<(), PackageError> {
        let (storage, _) = super::body_storage_objects::body_storage(self).map_err(invalid)?;
        let limits = super::storage_rewrite_limits(self.state.source.limits())
            .map_err(|_error| invalid("storage limits"))?;
        let length = storage.text.encode_utf16().count();
        let existing = rewrites.iter().position(|rewrite| {
            rewrite.object == storage.identifier && rewrite.message_index == storage.message_index
        });
        let base = existing.map_or(storage.payload.as_slice(), |index| {
            rewrites[index].data.as_slice()
        });
        let data = litchi_iwa_text_wire::rewrite_storage_text(base, 0..length, text, limits)
            .map_err(invalid)?
            .into_bytes();
        match existing {
            Some(index) => rewrites[index].data = data,
            None => rewrites.push(MessageRewrite {
                component: storage.component,
                object: storage.identifier,
                message_index: storage.message_index,
                message_type: storage.message_type,
                data,
            }),
        }
        Ok(())
    }

    /// The rewrites writing the document `language` and `locale` where the
    /// host builder writes them. `None` keeps a value.
    pub(crate) fn language_and_locale_rewrites(
        &self,
        language: Option<&str>,
        locale: Option<&str>,
    ) -> Result<Vec<MessageRewrite>, PackageError> {
        let mut rewrites = Vec::new();
        for component in self.state.source.components().iter() {
            let mut name = None;
            for object in &component.archive().objects {
                for (message_index, message) in object.messages.iter().enumerate() {
                    let data = match message.type_ {
                        DOCUMENT_MESSAGE_TYPE => document(&message.data, language, locale),
                        STORAGE_MESSAGE_TYPE => match language {
                            Some(language) => storage(&message.data, language),
                            None => continue,
                        },
                        SETTINGS_MESSAGE_TYPE => settings(&message.data, language, locale),
                        _ => continue,
                    }
                    .map_err(invalid)?;
                    if data == message.data {
                        continue;
                    }
                    let object = object
                        .archive_info
                        .identifier
                        .and_then(std::num::NonZeroU64::new)
                        .ok_or_else(|| invalid("an object has no identifier"))?;
                    let component_name = std::sync::Arc::clone(
                        name.get_or_insert_with(|| std::sync::Arc::from(component.name())),
                    );
                    rewrites.push(MessageRewrite {
                        component: component_name,
                        object,
                        message_index,
                        message_type: message.type_,
                        data,
                    });
                }
            }
        }
        Ok(rewrites)
    }
}
