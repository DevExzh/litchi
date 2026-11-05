//! Document language and locale of a fresh presentation (record 0372; the
//! migration host builder's `language`/`locale`).
//!
//! The host writes the language into the document archive and every text
//! storage's language table, and the locale into the document archive (its
//! current and creation locale) and the calculation engine. This rewrites
//! exactly those fields on the wire, inside the builder's single
//! fresh-settings publication (`super::blank_layout`).

use super::structure::StructureError;
use super::structure::wire::{Fields, LENGTH_DELIMITED};

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
/// `KN.DocumentArchive.super` (`TSA.DocumentArchive`).
const DOCUMENT_SUPER: u32 = 3;
/// `TSA.DocumentArchive.super` (`TSK.DocumentArchive`).
const TSA_SUPER: u32 = 1;
const TSA_DOCUMENT_LANGUAGE: u32 = 3;
const TSK_LOCALE: u32 = 4;
const TSK_CREATION_LOCALE: u32 = 9;
const STORAGE_TABLE_LANGUAGE: u32 = 19;
const ATTRIBUTE_TABLE_ENTRIES: u32 = 1;
const ATTRIBUTE_OBJECT: u32 = 2;
const ENGINE_SAVED_LOCALE: u32 = 16;

/// Set the singular string field `number` to `new`; returns whether it
/// changed.
fn set_string(fields: &mut Fields, number: u32, new: &str) -> bool {
    let current = fields
        .unique(number)
        .ok()
        .flatten()
        .and_then(|field| field.message().ok().map(<[u8]>::to_vec));
    if current.as_deref() == Some(new.as_bytes()) {
        return false;
    }
    fields.set_message(number, new.as_bytes());
    true
}

/// Rewrite one message for the document `language` and `locale` where the
/// host builder writes them; `None` keeps a value or leaves an unrelated
/// message. Applied by `Package::with_fresh_settings` (`super::blank_layout`).
pub(super) fn rewrite(
    message_type: u32,
    payload: &[u8],
    language: Option<&str>,
    locale: Option<&str>,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<Vec<u8>>, StructureError> {
    match message_type {
        DOCUMENT_MESSAGE_TYPE if language.is_some() || locale.is_some() => {
            document(payload, language, locale, wire)
        },
        STORAGE_MESSAGE_TYPE => match language {
            Some(language) => storage(payload, language, wire),
            None => Ok(None),
        },
        CALCULATION_ENGINE_MESSAGE_TYPE => match locale {
            Some(locale) => engine(payload, locale, wire),
            None => Ok(None),
        },
        _ => Ok(None),
    }
}

fn document(
    payload: &[u8],
    language: Option<&str>,
    locale: Option<&str>,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<Vec<u8>>, StructureError> {
    let mut fields = Fields::parse(payload, wire)?;
    let Some(tsa) = fields.unique(DOCUMENT_SUPER)? else {
        return Ok(None);
    };
    let mut tsa = Fields::parse(tsa.message()?, wire)?;
    let mut changed = false;
    if let Some(language) = language {
        changed |= set_string(&mut tsa, TSA_DOCUMENT_LANGUAGE, language);
    }
    if let Some(locale) = locale
        && let Some(tsk) = tsa.unique(TSA_SUPER)?
    {
        let mut tsk = Fields::parse(tsk.message()?, wire)?;
        let mut tsk_changed = set_string(&mut tsk, TSK_LOCALE, locale);
        tsk_changed |= set_string(&mut tsk, TSK_CREATION_LOCALE, locale);
        if tsk_changed {
            tsa.set_message(TSA_SUPER, &tsk.encode());
            changed = true;
        }
    }
    if !changed {
        return Ok(None);
    }
    fields.set_message(DOCUMENT_SUPER, &tsa.encode());
    Ok(Some(fields.encode()))
}

fn storage(
    payload: &[u8],
    language: &str,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<Vec<u8>>, StructureError> {
    let mut fields = Fields::parse(payload, wire)?;
    let Some(table) = fields.unique(STORAGE_TABLE_LANGUAGE)? else {
        return Ok(None);
    };
    let mut table = Fields::parse(table.message()?, wire)?;
    let mut entries = Vec::new();
    let mut changed = false;
    for entry in table.all(ATTRIBUTE_TABLE_ENTRIES) {
        let mut entry = Fields::parse(entry.message()?, wire)?;
        if entry.unique(ATTRIBUTE_OBJECT)?.is_some() {
            changed |= set_string(&mut entry, ATTRIBUTE_OBJECT, language);
        }
        entries.push(entry.encode());
    }
    if !changed {
        return Ok(None);
    }
    table.replace_all(ATTRIBUTE_TABLE_ENTRIES, LENGTH_DELIMITED, &entries);
    fields.set_message(STORAGE_TABLE_LANGUAGE, &table.encode());
    Ok(Some(fields.encode()))
}

fn engine(
    payload: &[u8],
    locale: &str,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<Vec<u8>>, StructureError> {
    let mut fields = Fields::parse(payload, wire)?;
    if fields.unique(ENGINE_SAVED_LOCALE)?.is_none() {
        return Ok(None);
    }
    Ok(set_string(&mut fields, ENGINE_SAVED_LOCALE, locale).then(|| fields.encode()))
}
