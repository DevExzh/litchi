//! Document language and locale fields (record 0370i; the migration host
//! builder's `language` and `locale`).
//!
//! The host builder writes the language into the document archive
//! (`TSA.DocumentArchive.document_language`) and the locale into its current
//! and creation locale (`TSK.DocumentArchive.locale_identifier`,
//! `creation_locale_identifier`) and the calculation engine
//! (`saved_locale_identifier`). These rewrite exactly those fields on the
//! wire, for the one-pass builder synthesis.

use litchi_iwa_common::wire::{
    parse_wire_view, patch_length_delimited_field, transform_length_delimited_field,
};

use super::structure::{self, Fault, FaultResult};

/// `TN.DocumentArchive.super` (`TSA.DocumentArchive`).
const DOCUMENT_SUPER: u32 = 8;
/// `TSA.DocumentArchive.super` (`TSK.DocumentArchive`).
const TSA_SUPER: u32 = 1;
const TSA_DOCUMENT_LANGUAGE: u32 = 3;
const TSK_LOCALE: u32 = 4;
const TSK_CREATION_LOCALE: u32 = 9;
const ENGINE_SAVED_LOCALE: u32 = 16;

/// A [`Fault`] carried through the wire callbacks.
struct WireFault(Fault);

impl From<litchi_iwa_common::Error> for WireFault {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self(structure::map_wire(&error))
    }
}

fn has(data: &[u8], number: u32) -> Result<bool, WireFault> {
    let view = parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT)?;
    Ok(view.fields().any(|field| field.number() == number))
}

/// `data` with string field `number` set to `value`.
fn set(data: &[u8], number: u32, value: &str) -> Result<Vec<u8>, WireFault> {
    Ok(patch_length_delimited_field(
        data,
        number,
        has(data, number)?,
        Some(value.as_bytes()),
    )?)
}

/// A `TN.DocumentArchive` payload with `language` and `locale` written;
/// `None` keeps a value.
pub(super) fn document(
    data: &[u8],
    language: Option<&str>,
    locale: Option<&str>,
) -> FaultResult<Vec<u8>> {
    transform_length_delimited_field(data, DOCUMENT_SUPER, |tsa| {
        let mut tsa = match language {
            Some(language) => set(tsa, TSA_DOCUMENT_LANGUAGE, language)?,
            None => tsa.to_vec(),
        };
        if let Some(locale) = locale {
            tsa = transform_length_delimited_field(&tsa, TSA_SUPER, |tsk| {
                let tsk = set(tsk, TSK_LOCALE, locale)?;
                set(&tsk, TSK_CREATION_LOCALE, locale)
            })?;
        }
        Ok::<_, WireFault>(tsa)
    })
    .map_err(|error| error.0)
}

/// A `TSCE.CalculationEngineArchive` payload with its saved locale.
pub(super) fn engine(data: &[u8], locale: &str) -> FaultResult<Vec<u8>> {
    set(data, ENGINE_SAVED_LOCALE, locale).map_err(|error| error.0)
}
