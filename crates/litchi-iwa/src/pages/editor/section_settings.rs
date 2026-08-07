//! Typed, wire-preserving Pages section settings and backgrounds.

use super::*;

const INHERIT_HEADER_FOOTER_FIELD: u32 = 17;
const FIRST_PAGE_DIFFERENT_FIELD: u32 = 18;
const EVEN_ODD_PAGES_DIFFERENT_FIELD: u32 = 19;
const SECTION_START_FIELD: u32 = 20;
const PAGE_NUMBERING_FIELD: u32 = 21;
const STARTING_PAGE_NUMBER_FIELD: u32 = 22;
const SECTION_NAME_FIELD: u32 = 26;
const FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD: u32 = 28;
const BACKGROUND_FILL_FIELD: u32 = 30;

impl PagesEditor {
    /// Read the lossless settings payload of a reachable Pages section.
    pub fn section_settings(&self, section_id: u64) -> Result<Settings> {
        if !self
            .sections
            .iter()
            .any(|section| section.object_id == section_id)
        {
            return Err(Error::ParseError(format!(
                "Section {section_id} is not reachable from the Pages body"
            )));
        }
        decode_section_settings(&self.section_message_data(section_id)?)
    }

    /// Replace the settings stored directly on a reachable Pages section.
    ///
    /// The update is transactional and patches only changed protobuf fields,
    /// preserving unknown fields, raw background-fill bytes, and field order.
    pub fn set_section_settings(&mut self, section_id: u64, settings: Settings) -> Result<()> {
        validate_section_settings(&settings)?;
        let current = self.section_settings(section_id)?;
        if current == settings {
            return Ok(());
        }

        let mut staged = self.text.package().clone();
        let archive_name = find_section_archive(&staged, section_id)?;
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(section_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Pages section object {section_id} is missing"))
            })?;
            let message_indexes = object
                .messages
                .iter()
                .enumerate()
                .filter_map(|(index, message)| {
                    (message.type_ == SECTION_MESSAGE_TYPE).then_some(index)
                })
                .collect::<Vec<_>>();
            if message_indexes.len() != 1 {
                return Err(Error::InvalidFormat(format!(
                    "Pages section object {section_id} must have one section payload, found {}",
                    message_indexes.len()
                )));
            }
            let message_index = message_indexes[0];
            let original = object.messages[message_index].data.as_slice();
            let decoded_current = decode_section_settings(original)?;
            if decoded_current != current {
                return Err(Error::InvalidFormat(format!(
                    "Pages section {section_id} changed during mutation"
                )));
            }

            let mut data = original.to_vec();
            for (field_number, before, after) in [
                (
                    INHERIT_HEADER_FOOTER_FIELD,
                    current.inherit_previous_header_footer(),
                    settings.inherit_previous_header_footer(),
                ),
                (
                    FIRST_PAGE_DIFFERENT_FIELD,
                    current.first_page_different(),
                    settings.first_page_different(),
                ),
                (
                    EVEN_ODD_PAGES_DIFFERENT_FIELD,
                    current.even_odd_pages_different(),
                    settings.even_odd_pages_different(),
                ),
                (
                    FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD,
                    current.first_page_hides_header_footer(),
                    settings.first_page_hides_header_footer(),
                ),
            ] {
                if before != after {
                    data = patch_varint_field(
                        &data,
                        field_number,
                        before.is_some(),
                        after.map(u64::from),
                    )?;
                }
            }
            for (field_number, before, after) in [
                (
                    SECTION_START_FIELD,
                    current.start().map(Start::as_raw),
                    settings.start().map(Start::as_raw),
                ),
                (
                    PAGE_NUMBERING_FIELD,
                    current.page_numbering().map(PageNumbering::as_raw),
                    settings.page_numbering().map(PageNumbering::as_raw),
                ),
                (
                    STARTING_PAGE_NUMBER_FIELD,
                    current.starting_page_number().map(PageNumber::get),
                    settings.starting_page_number().map(PageNumber::get),
                ),
            ] {
                if before != after {
                    data = patch_varint_field(
                        &data,
                        field_number,
                        before.is_some(),
                        after.map(u64::from),
                    )?;
                }
            }
            if current.name() != settings.name() {
                data = patch_length_delimited_field(
                    &data,
                    SECTION_NAME_FIELD,
                    current.name().is_some(),
                    settings.name().map(str::as_bytes),
                )?;
            }
            if decode_section_settings(&data)? != settings {
                return Err(Error::InvalidFormat(format!(
                    "Pages section {section_id} settings patch failed validation"
                )));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: SECTION_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })?;
        *self = Self::from_package(staged)?;
        Ok(())
    }

    /// Read the exact encoded background-fill payload for a reachable section.
    ///
    /// This is an IWA-only bridge for the separate semantic background API;
    /// the payload never enters [`Settings`].
    pub(super) fn section_background_payload(&self, section_id: u64) -> Result<Option<Box<[u8]>>> {
        let data = self.section_message_data(section_id)?;
        raw_background_payload(&data)
    }

    /// Replace only the native background-fill field of a reachable section.
    ///
    /// The operation preserves every other section field and rechecks the
    /// source payload inside the transaction before publication.
    pub(super) fn set_section_background_payload(
        &mut self,
        section_id: u64,
        payload: Option<&[u8]>,
    ) -> Result<()> {
        if let Some(payload) = payload {
            tsd::FillArchive::decode(payload).map_err(|error| {
                Error::ParseError(format!(
                    "Pages section background fill is not a TSD.FillArchive: {error}"
                ))
            })?;
        }
        let current = self.section_background_payload(section_id)?;
        if current.as_deref() == payload {
            return Ok(());
        }

        let mut staged = self.text.package().clone();
        let archive_name = find_section_archive(&staged, section_id)?;
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(section_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Pages section object {section_id} is missing"))
            })?;
            let message_indexes = object
                .messages
                .iter()
                .enumerate()
                .filter_map(|(index, message)| {
                    (message.type_ == SECTION_MESSAGE_TYPE).then_some(index)
                })
                .collect::<Vec<_>>();
            if message_indexes.len() != 1 {
                return Err(Error::InvalidFormat(format!(
                    "Pages section object {section_id} must have one section payload, found {}",
                    message_indexes.len()
                )));
            }
            let message_index = message_indexes[0];
            let original = object.messages[message_index].data.as_slice();
            if raw_background_payload(original)?.as_deref() != current.as_deref() {
                return Err(Error::InvalidFormat(format!(
                    "Pages section {section_id} changed during mutation"
                )));
            }
            let data = patch_length_delimited_field(
                original,
                BACKGROUND_FILL_FIELD,
                current.is_some(),
                payload,
            )?;
            decode_section_settings(&data)?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: SECTION_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })?;
        *self = Self::from_package(staged)?;
        Ok(())
    }

    pub(super) fn section_message_data(&self, section_id: u64) -> Result<Vec<u8>> {
        if !self
            .sections
            .iter()
            .any(|section| section.object_id == section_id)
        {
            return Err(Error::ParseError(format!(
                "Section {section_id} is not reachable from the Pages body"
            )));
        }
        let archive_name = find_section_archive(self.text.package(), section_id)?;
        let archive = self.text.package().archive(&archive_name)?;
        let object = archive.object(section_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages section object {section_id} is missing"))
        })?;
        let messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == SECTION_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Pages section object {section_id} must have one section payload, found {}",
                messages.len()
            )));
        }
        Ok(messages[0].data.clone())
    }
}

fn validate_section_settings(settings: &Settings) -> Result<()> {
    settings
        .validate()
        .map_err(|error| Error::ParseError(error.to_string()))
}

fn decode_section_settings(data: &[u8]) -> Result<Settings> {
    let section = SectionArchive::decode(data)?;

    // Validate raw singularity and wire types even though prost accepts the
    // final occurrence of duplicate scalar fields.
    for (field_number, present, value) in [
        (
            INHERIT_HEADER_FOOTER_FIELD,
            section.inherit_previous_header_footer.is_some(),
            section.inherit_previous_header_footer.map(u64::from),
        ),
        (
            FIRST_PAGE_DIFFERENT_FIELD,
            section.section_template_first_page_different.is_some(),
            section.section_template_first_page_different.map(u64::from),
        ),
        (
            EVEN_ODD_PAGES_DIFFERENT_FIELD,
            section.section_template_even_odd_pages_different.is_some(),
            section
                .section_template_even_odd_pages_different
                .map(u64::from),
        ),
        (
            SECTION_START_FIELD,
            section.section_start_kind.is_some(),
            section.section_start_kind.map(u64::from),
        ),
        (
            PAGE_NUMBERING_FIELD,
            section.section_page_number_kind.is_some(),
            section.section_page_number_kind.map(u64::from),
        ),
        (
            STARTING_PAGE_NUMBER_FIELD,
            section.section_page_number_start.is_some(),
            section.section_page_number_start.map(u64::from),
        ),
        (
            FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD,
            section
                .section_template_first_page_hides_header_footer
                .is_some(),
            section
                .section_template_first_page_hides_header_footer
                .map(u64::from),
        ),
    ] {
        patch_varint_field(data, field_number, present, value)?;
    }
    patch_length_delimited_field(
        data,
        SECTION_NAME_FIELD,
        section.name.is_some(),
        section.name.as_deref().map(str::as_bytes),
    )?;

    let background_payloads = repeated_length_delimited_payloads(data, BACKGROUND_FILL_FIELD)?;
    if background_payloads.len() > 1 {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {BACKGROUND_FILL_FIELD} occurs {} times",
            background_payloads.len()
        )));
    }
    if background_payloads.is_empty() != section.background_fill.is_none() {
        return Err(Error::InvalidFormat(
            "Pages section background-fill presence changed during decoding".to_owned(),
        ));
    }
    if let Some(payload) = background_payloads.first().copied() {
        tsd::FillArchive::decode(payload)?;
    }
    patch_length_delimited_field(
        data,
        BACKGROUND_FILL_FIELD,
        section.background_fill.is_some(),
        background_payloads.first().copied(),
    )?;

    let mut settings = Settings::new();
    settings
        .set_name(section.name.map(String::into_boxed_str))
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
    settings.set_inherit_previous_header_footer(section.inherit_previous_header_footer);
    settings.set_first_page_different(section.section_template_first_page_different);
    settings.set_even_odd_pages_different(section.section_template_even_odd_pages_different);
    settings
        .set_start(section.section_start_kind.map(Start::from_raw))
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
    settings
        .set_page_numbering(
            section
                .section_page_number_kind
                .map(PageNumbering::from_raw),
        )
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
    settings.set_starting_page_number(
        section
            .section_page_number_start
            .map(|value| {
                PageNumber::new(value).map_err(|_| {
                    Error::InvalidFormat(format!(
                        "Pages section has invalid starting page number {value}"
                    ))
                })
            })
            .transpose()?,
    );
    settings.set_first_page_hides_header_footer(
        section.section_template_first_page_hides_header_footer,
    );
    Ok(settings)
}

fn raw_background_payload(data: &[u8]) -> Result<Option<Box<[u8]>>> {
    let section = SectionArchive::decode(data)?;
    let payloads = repeated_length_delimited_payloads(data, BACKGROUND_FILL_FIELD)?;
    if payloads.len() > 1 {
        return Err(Error::InvalidFormat(format!(
            "singular protobuf field {BACKGROUND_FILL_FIELD} occurs {} times",
            payloads.len()
        )));
    }
    if payloads.is_empty() != section.background_fill.is_none() {
        return Err(Error::InvalidFormat(
            "Pages section background-fill presence changed during decoding".to_owned(),
        ));
    }
    let Some(payload) = payloads.first().copied() else {
        return Ok(None);
    };
    tsd::FillArchive::decode(payload)?;
    Ok(Some(payload.to_vec().into_boxed_slice()))
}
