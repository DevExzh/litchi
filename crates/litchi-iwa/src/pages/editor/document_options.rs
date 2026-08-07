//! Typed Pages document formatter options.

use crate::protobuf::tp::SettingsArchive;

use super::*;

mod wire;

use wire::{read_document_options_wire, write_document_options_wire};

const SETTINGS_REFERENCE_FIELD: u32 = 7;
pub(super) const SETTINGS_MESSAGE_TYPE: u32 = 10_012;

fn options_from_settings(settings: &SettingsArchive) -> DocumentOptions {
    DocumentOptions::new(
        settings.body,
        settings.headers,
        settings.footers,
        settings.facing_pages,
        settings.hyphenation,
        settings.use_ligatures,
    )
}

impl PagesEditor {
    /// Read the lossless options shown by Pages' Document formatter.
    pub fn document_options(&self) -> Result<DocumentOptions> {
        Ok(locate_settings(self.text.package())?.options)
    }

    /// Replace Pages' Document formatter options transactionally.
    pub fn set_document_options(&mut self, options: DocumentOptions) -> Result<()> {
        let location = locate_settings(self.text.package())?;
        if location.options == options {
            return Ok(());
        }

        let mut staged = self.text.package().clone();
        staged.update_archive(&location.archive_name, |archive| {
            let object = archive.object_mut(location.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages settings object {} is missing",
                    location.identifier
                ))
            })?;
            let message_index = settings_message_index(object, location.identifier)?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let settings = SettingsArchive::decode(original)?;
            let data = write_document_options_wire(original, &settings, options)?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;

        let verified = Self::from_package(staged)?;
        if verified.document_options()? != options {
            return Err(Error::InvalidFormat(
                "Pages document options failed round-trip validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

pub(super) struct SettingsLocation {
    pub(super) identifier: u64,
    pub(super) archive_name: String,
    pub(super) data: Vec<u8>,
    pub(super) settings: SettingsArchive,
    options: DocumentOptions,
}

pub(super) fn locate_settings(package: &IWorkPackage) -> Result<SettingsLocation> {
    let identifier = settings_identifier(package)?;
    let mut found = None;
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        if found.is_some() {
            return Err(Error::InvalidFormat(format!(
                "Pages settings object {identifier} occurs in more than one archive"
            )));
        }
        let message_index = settings_message_index(object, identifier)?;
        let original = object.messages[message_index].data.as_slice();
        let settings = SettingsArchive::decode(original)?;
        let options = read_document_options_wire(original, &settings)?;
        found = Some(SettingsLocation {
            identifier,
            archive_name: archive_name.to_owned(),
            data: original.to_vec(),
            settings,
            options,
        });
    }
    found.ok_or_else(|| {
        Error::InvalidFormat(format!("Pages settings object {identifier} is missing"))
    })
}

fn settings_identifier(package: &IWorkPackage) -> Result<u64> {
    let archive = package.archive(DOCUMENT_ARCHIVE_NAME)?;
    let object = archive.object(DOCUMENT_OBJECT_ID).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages root object {DOCUMENT_OBJECT_ID} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages root must contain exactly one TP.DocumentArchive payload, found {}",
            messages.len()
        )));
    };
    let document = DocumentArchive::decode(message.data.as_slice())?;
    let reference = document.settings.ok_or_else(|| {
        Error::InvalidFormat("Pages document has no settings reference".to_owned())
    })?;
    let raw_references =
        repeated_length_delimited_payloads(message.data.as_slice(), SETTINGS_REFERENCE_FIELD)?;
    let [raw_reference] = raw_references.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Pages root settings reference occurs {} times",
            raw_references.len()
        )));
    };
    if tsp::Reference::decode(*raw_reference)? != reference || reference.identifier == 0 {
        return Err(Error::InvalidFormat(
            "Pages root settings reference is invalid or inconsistent".to_owned(),
        ));
    }
    Ok(reference.identifier)
}

pub(super) fn settings_message_index(object: &ArchiveObject, identifier: u64) -> Result<usize> {
    let matches = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| (message.type_ == SETTINGS_MESSAGE_TYPE).then_some(index))
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [index] => Ok(*index),
        _ => Err(Error::InvalidFormat(format!(
            "Pages settings object {identifier} must contain exactly one payload, found {}",
            matches.len()
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn absent_options_use_native_effective_defaults() {
        let options = DocumentOptions::default();
        assert!(options.body_is_enabled());
        assert!(options.headers_are_enabled());
        assert!(options.footers_are_enabled());
        assert!(!options.uses_facing_pages());
        assert!(!options.uses_automatic_hyphenation());
        assert!(!options.uses_ligatures());
    }
}
