//! Typed Pages document formatter options.

use crate::protobuf::tp::SettingsArchive;

use super::*;

mod wire;

use wire::{read_document_options_wire, write_document_options_wire};

const SETTINGS_REFERENCE_FIELD: u32 = 7;
const SETTINGS_MESSAGE_TYPE: u32 = 10_012;

/// Lossless options exposed by Pages' Document formatter.
///
/// Every field retains its optional protobuf presence. The convenience methods
/// return the effective native defaults without erasing that distinction.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct PagesDocumentOptions {
    pub body_enabled: Option<bool>,
    pub headers_enabled: Option<bool>,
    pub footers_enabled: Option<bool>,
    pub facing_pages: Option<bool>,
    pub automatic_hyphenation: Option<bool>,
    pub ligatures_enabled: Option<bool>,
}

impl PagesDocumentOptions {
    /// Return whether the document body is effectively enabled.
    pub fn body_is_enabled(self) -> bool {
        self.body_enabled.unwrap_or(true)
    }

    /// Return whether headers are effectively enabled.
    pub fn headers_are_enabled(self) -> bool {
        self.headers_enabled.unwrap_or(true)
    }

    /// Return whether footers are effectively enabled.
    pub fn footers_are_enabled(self) -> bool {
        self.footers_enabled.unwrap_or(true)
    }

    /// Return whether facing-page layout is effectively enabled.
    pub fn uses_facing_pages(self) -> bool {
        self.facing_pages.unwrap_or(false)
    }

    /// Return whether automatic hyphenation is effectively enabled.
    pub fn uses_automatic_hyphenation(self) -> bool {
        self.automatic_hyphenation.unwrap_or(false)
    }

    /// Return whether typographic ligatures are effectively enabled.
    pub fn uses_ligatures(self) -> bool {
        self.ligatures_enabled.unwrap_or(false)
    }

    fn from_settings(settings: &SettingsArchive) -> Self {
        Self {
            body_enabled: settings.body,
            headers_enabled: settings.headers,
            footers_enabled: settings.footers,
            facing_pages: settings.facing_pages,
            automatic_hyphenation: settings.hyphenation,
            ligatures_enabled: settings.use_ligatures,
        }
    }
}

impl PagesEditor {
    /// Read the lossless options shown by Pages' Document formatter.
    pub fn document_options(&self) -> Result<PagesDocumentOptions> {
        Ok(locate_settings(self.text.package())?.options)
    }

    /// Replace Pages' Document formatter options transactionally.
    pub fn set_document_options(&mut self, options: PagesDocumentOptions) -> Result<()> {
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

struct SettingsLocation {
    identifier: u64,
    archive_name: String,
    options: PagesDocumentOptions,
}

fn locate_settings(package: &IWorkPackage) -> Result<SettingsLocation> {
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

fn settings_message_index(object: &ArchiveObject, identifier: u64) -> Result<usize> {
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
        let options = PagesDocumentOptions::default();
        assert!(options.body_is_enabled());
        assert!(options.headers_are_enabled());
        assert!(options.footers_are_enabled());
        assert!(!options.uses_facing_pages());
        assert!(!options.uses_automatic_hyphenation());
        assert!(!options.uses_ligatures());
    }
}
