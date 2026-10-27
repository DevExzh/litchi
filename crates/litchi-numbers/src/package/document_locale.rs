//! Document language and locale of a fresh workbook (record 0370i; the
//! migration host builder's `language` and `locale`).
//!
//! The host builder writes the language into the document archive
//! (`TSA.DocumentArchive.document_language`) and the locale into its current
//! and creation locale (`TSK.DocumentArchive.locale_identifier`,
//! `creation_locale_identifier`) and the calculation engine
//! (`saved_locale_identifier`). This rewrites exactly those fields on the
//! wire. A builder table of another size also records its extent in its
//! tile's `maxColumn` and `maxRow`, as the host builder writes them.

use litchi_iwa_common::wire::{
    parse_wire_view, patch_length_delimited_field, patch_varint_field,
    repeated_length_delimited_payloads, repeated_varint_values, transform_length_delimited_field,
};
use litchi_iwa_core::RawMessage;

use super::Package;
use super::structure::{self, Fault, FaultResult, Working, metadata::Transition};

/// `TN.DocumentArchive.super` (`TSA.DocumentArchive`).
const DOCUMENT_SUPER: u32 = 8;
/// `TSA.DocumentArchive.super` (`TSK.DocumentArchive`).
const TSA_SUPER: u32 = 1;
const TSA_DOCUMENT_LANGUAGE: u32 = 3;
const TSK_LOCALE: u32 = 4;
const TSK_CREATION_LOCALE: u32 = 9;
/// `TSCE.CalculationEngineArchive`.
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const ENGINE_SAVED_LOCALE: u32 = 16;
/// `TST.TableModelArchive` and its data store's tile list
/// (`base_data_store.tiles.tiles[].tile`).
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MAX_COLUMN: u32 = 1;
const TILE_MAX_ROW: u32 = 2;

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

fn document(data: &[u8], language: Option<&str>, locale: Option<&str>) -> FaultResult<Vec<u8>> {
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

/// The tiles of every table model of the workbook.
fn tiles(working: &Working<'_>) -> FaultResult<Vec<u64>> {
    let wire = |error: litchi_iwa_common::Error| structure::map_wire(&error);
    let mut output = Vec::new();
    for name in working.component_names() {
        for object in &working.archive(&name)?.objects {
            for message in &object.messages {
                if message.type_ != TABLE_MODEL_MESSAGE_TYPE {
                    continue;
                }
                for store in repeated_length_delimited_payloads(&message.data, 4).map_err(wire)? {
                    for list in repeated_length_delimited_payloads(store, 3).map_err(wire)? {
                        for entry in repeated_length_delimited_payloads(list, 1).map_err(wire)? {
                            for tile in
                                repeated_length_delimited_payloads(entry, 2).map_err(wire)?
                            {
                                if let Some(identifier) =
                                    repeated_varint_values(tile, 1).map_err(wire)?.last()
                                {
                                    output.push(*identifier);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    Ok(output)
}

fn rewrite(
    source: &Package,
    language: Option<&str>,
    locale: Option<&str>,
    extent: Option<(u32, u32)>,
) -> FaultResult<structure::Published> {
    let mut working = Working::new(source)?;
    working.keep_previews();
    if let Some((rows, columns)) = extent {
        let limits = working.archive_limits();
        for tile in tiles(&working)? {
            let object = working.object(tile)?;
            let [message] = object.messages.as_slice() else {
                return Err(Fault::Unsupported);
            };
            let message_type = message.type_;
            let data = patch_varint_field(
                &message.data,
                TILE_MAX_COLUMN,
                true,
                Some(u64::from(columns.saturating_sub(1))),
            )
            .and_then(|data| {
                patch_varint_field(
                    &data,
                    TILE_MAX_ROW,
                    true,
                    Some(u64::from(rows.saturating_sub(1))),
                )
            })
            .map_err(|error| structure::map_wire(&error))?;
            working
                .object_mut(tile)?
                .replace_message_preserving_header(
                    0,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                    limits,
                )
                .map_err(|error| structure::map_core(&error))?;
        }
    }
    if language.is_none() && locale.is_none() {
        return working.publish(&Transition::default());
    }
    let document_object = working.object(1)?;
    let index = structure::document_message(document_object)?;
    let data = document(&document_object.messages[index].data, language, locale)?;
    let message_type = document_object.messages[index].type_;
    let limits = working.archive_limits();
    working
        .object_mut(1)?
        .replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    if let Some(locale) = locale {
        let mut engines = Vec::new();
        for name in working.component_names() {
            for object in &working.archive(&name)?.objects {
                for (index, message) in object.messages.iter().enumerate() {
                    if message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE {
                        let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
                        engines.push((identifier, index));
                    }
                }
            }
        }
        for (identifier, index) in engines {
            let data = set(
                &working.object(identifier)?.messages[index].data,
                ENGINE_SAVED_LOCALE,
                locale,
            )
            .map_err(|error| error.0)?;
            working
                .object_mut(identifier)?
                .replace_message_preserving_header(
                    index,
                    RawMessage {
                        type_: CALCULATION_ENGINE_MESSAGE_TYPE,
                        data,
                    },
                    limits,
                )
                .map_err(|error| structure::map_core(&error))?;
        }
    }
    working.publish(&Transition::default())
}

impl Package {
    /// A reopened copy with the document `language` and `locale` written
    /// where the host builder writes them (`None` keeps a value), and with
    /// a builder table's `extent` (rows, columns) in its tiles.
    pub(crate) fn with_builder_options(
        &self,
        language: Option<&str>,
        locale: Option<&str>,
        extent: Option<(u32, u32)>,
    ) -> crate::Result<Self> {
        if language.is_none() && locale.is_none() && extent.is_none() {
            return Ok(self.clone());
        }
        rewrite(self, language, locale, extent)
            .map(|published| published.package)
            .map_err(|fault| {
                crate::Error::InvalidFormat(format!(
                    "the Numbers document language or locale could not be written: {fault:?}"
                ))
            })
    }
}
