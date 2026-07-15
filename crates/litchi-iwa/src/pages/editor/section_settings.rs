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
    pub fn section_settings(&self, section_id: u64) -> Result<PagesSectionSettings> {
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
        decode_section_settings(messages[0].data.as_slice())
    }

    /// Replace the settings stored directly on a reachable Pages section.
    ///
    /// The update is transactional and patches only changed protobuf fields,
    /// preserving unknown fields, raw background-fill bytes, and field order.
    pub fn set_section_settings(
        &mut self,
        section_id: u64,
        settings: PagesSectionSettings,
    ) -> Result<()> {
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
                    current.inherit_previous_header_footer,
                    settings.inherit_previous_header_footer,
                ),
                (
                    FIRST_PAGE_DIFFERENT_FIELD,
                    current.first_page_different,
                    settings.first_page_different,
                ),
                (
                    EVEN_ODD_PAGES_DIFFERENT_FIELD,
                    current.even_odd_pages_different,
                    settings.even_odd_pages_different,
                ),
                (
                    FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD,
                    current.first_page_hides_header_footer,
                    settings.first_page_hides_header_footer,
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
                    current.start.map(PagesSectionStart::as_raw),
                    settings.start.map(PagesSectionStart::as_raw),
                ),
                (
                    PAGE_NUMBERING_FIELD,
                    current
                        .page_numbering
                        .map(PagesSectionPageNumbering::as_raw),
                    settings
                        .page_numbering
                        .map(PagesSectionPageNumbering::as_raw),
                ),
                (
                    STARTING_PAGE_NUMBER_FIELD,
                    current.starting_page_number.map(PagesPageNumber::get),
                    settings.starting_page_number.map(PagesPageNumber::get),
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
            if current.name != settings.name {
                data = patch_length_delimited_field(
                    &data,
                    SECTION_NAME_FIELD,
                    current.name.is_some(),
                    settings.name.as_deref().map(str::as_bytes),
                )?;
            }
            if current.background_fill_payload != settings.background_fill_payload {
                data = patch_length_delimited_field(
                    &data,
                    BACKGROUND_FILL_FIELD,
                    current.background_fill_payload.is_some(),
                    settings.background_fill_payload.as_deref(),
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
}

fn validate_section_settings(settings: &PagesSectionSettings) -> Result<()> {
    if settings.start.is_some_and(|start| !start.is_canonical()) {
        return Err(Error::ParseError(
            "Pages unknown section start must not alias a known start".to_owned(),
        ));
    }
    if settings
        .page_numbering
        .is_some_and(|numbering| !numbering.is_canonical())
    {
        return Err(Error::ParseError(
            "Pages unknown page-numbering behavior must not alias a known behavior".to_owned(),
        ));
    }
    if settings
        .name
        .as_deref()
        .is_some_and(|name| name.contains('\0'))
    {
        return Err(Error::ParseError(
            "Pages section names cannot contain NUL".to_owned(),
        ));
    }
    if let Some(payload) = settings.background_fill_payload.as_deref() {
        tsd::FillArchive::decode(payload).map_err(|error| {
            Error::ParseError(format!(
                "Pages section background fill is not a TSD.FillArchive: {error}"
            ))
        })?;
    }
    Ok(())
}

fn decode_section_settings(data: &[u8]) -> Result<PagesSectionSettings> {
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
    let background_fill_payload = background_payloads.first().map(|payload| payload.to_vec());
    if let Some(payload) = background_fill_payload.as_deref() {
        tsd::FillArchive::decode(payload)?;
    }
    patch_length_delimited_field(
        data,
        BACKGROUND_FILL_FIELD,
        section.background_fill.is_some(),
        background_fill_payload.as_deref(),
    )?;

    Ok(PagesSectionSettings {
        name: section.name,
        inherit_previous_header_footer: section.inherit_previous_header_footer,
        first_page_different: section.section_template_first_page_different,
        even_odd_pages_different: section.section_template_even_odd_pages_different,
        start: section.section_start_kind.map(PagesSectionStart::from_raw),
        page_numbering: section
            .section_page_number_kind
            .map(PagesSectionPageNumbering::from_raw),
        starting_page_number: section
            .section_page_number_start
            .map(|value| {
                PagesPageNumber::new(value).map_err(|_| {
                    Error::InvalidFormat(format!(
                        "Pages section has invalid starting page number {value}"
                    ))
                })
            })
            .transpose()?,
        first_page_hides_header_footer: section.section_template_first_page_hides_header_footer,
        background_fill_payload,
    })
}
