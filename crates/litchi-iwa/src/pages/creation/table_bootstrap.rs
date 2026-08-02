//! Transactional installation of the first native table into a scratch Pages package.

use super::*;
use crate::package_metadata::{
    add_component_external_reference, add_component_link, add_component_object_uuids,
    add_component_registration, component_identifier_for_entry,
};

const CALCULATION_ENGINE_ENTRY: &str = "Index/CalculationEngine.iwa";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in crate::pages) struct BootstrappedTableGraph {
    pub(in crate::pages) info_object_id: u64,
    pub(in crate::pages) model_object_id: u64,
    pub(in crate::pages) attachment_object_id: u64,
}

pub(in crate::pages) fn bootstrap_first_table_graph(
    package: &mut IWorkPackage,
    body_storage_id: u64,
    name: &str,
    rows: usize,
    columns: usize,
) -> Result<BootstrappedTableGraph> {
    let table = InitialPagesTable {
        name: name.to_owned(),
        rows,
        columns,
    };
    validate_initial_table(&table)?;
    require_scratch_table_slots(package, body_storage_id)?;

    let document = package.archive(DOCUMENT_ARCHIVE_ENTRY)?;
    let document_object = document
        .object(PagesObjectId::Document.value())
        .ok_or_else(|| {
            crate::Error::InvalidFormat("scratch Pages document root is missing".to_owned())
        })?;
    let document_message = document_object
        .messages
        .iter()
        .find(|message| message.type_ == PagesMessageType::Document.value())
        .ok_or_else(|| {
            crate::Error::InvalidFormat("scratch Pages document payload is missing".to_owned())
        })?;
    let decoded_document = tp::DocumentArchive::decode(document_message.data.as_slice())?;
    let language = decoded_document
        .super_
        .document_language
        .as_deref()
        .unwrap_or(DEFAULT_LANGUAGE);
    let locale = decoded_document
        .super_
        .super_
        .locale_identifier
        .as_deref()
        .unwrap_or(DEFAULT_LOCALE);
    install_initial_table_graph(package, &table, language, locale)?;
    patch_document_scaffolding(package)?;
    register_table_components(package)?;

    Ok(BootstrappedTableGraph {
        info_object_id: TABLE_INFO_OBJECT_ID,
        model_object_id: TABLE_MODEL_OBJECT_ID,
        attachment_object_id: TABLE_ATTACHMENT_OBJECT_ID,
    })
}

fn require_scratch_table_slots(package: &IWorkPackage, body_storage_id: u64) -> Result<()> {
    if body_storage_id != PagesObjectId::Body.value()
        || component_identifier_for_entry(package, DOCUMENT_ARCHIVE_ENTRY)?
            != Some(PagesObjectId::Document.value())
        || component_identifier_for_entry(package, STYLESHEET_ARCHIVE_ENTRY)?
            != Some(PagesObjectId::Stylesheet.value())
    {
        return Err(crate::Error::ParseError(
            "The first runtime Pages table requires a litchi-iwa scratch document or an existing native table template"
                .to_owned(),
        ));
    }
    if package.contains_entry(CALCULATION_ENGINE_ENTRY) {
        return Err(crate::Error::InvalidFormat(
            "table-less scratch Pages package already contains a CalculationEngine".to_owned(),
        ));
    }
    let reserved = TABLE_DOCUMENT_OBJECT_IDS
        .iter()
        .chain(TABLE_STYLE_OBJECT_IDS)
        .copied()
        .chain([
            TABLE_CALCULATION_ENGINE_OBJECT_ID,
            TABLE_FORMULA_OWNER_OBJECT_ID,
            TABLE_ATTACHMENT_OBJECT_ID,
        ])
        .collect::<std::collections::HashSet<_>>();
    for entry in package.iwa_entry_names() {
        for object in package.archive(entry)?.objects {
            if object
                .archive_info
                .identifier
                .is_some_and(|identifier| reserved.contains(&identifier))
            {
                return Err(crate::Error::InvalidFormat(format!(
                    "scratch Pages table object range collides in {entry}"
                )));
            }
        }
    }
    Ok(())
}

fn patch_document_scaffolding(package: &mut IWorkPackage) -> Result<()> {
    package.update_archive(DOCUMENT_ARCHIVE_ENTRY, |archive| {
        let document = archive
            .object_mut(PagesObjectId::Document.value())
            .ok_or_else(|| {
                crate::Error::InvalidFormat("scratch Pages document root is missing".to_owned())
            })?;
        let document_message_index = document
            .messages
            .iter()
            .position(|message| message.type_ == PagesMessageType::Document.value())
            .ok_or_else(|| {
                crate::Error::InvalidFormat("scratch Pages document payload is missing".to_owned())
            })?;
        let document_message_type = document.messages[document_message_index].type_;
        let mut decoded =
            tp::DocumentArchive::decode(document.messages[document_message_index].data.as_slice())?;
        decoded.super_.calculation_engine = Some(raw_reference(TABLE_CALCULATION_ENGINE_OBJECT_ID));
        decoded.super_.function_browser_state =
            Some(raw_reference(TABLE_FUNCTION_BROWSER_STATE_OBJECT_ID));
        decoded.super_.custom_format_list = Some(raw_reference(TABLE_CUSTOM_FORMAT_LIST_OBJECT_ID));
        document.replace_message(
            document_message_index,
            RawMessage {
                type_: document_message_type,
                data: decoded.encode_to_vec(),
            },
        )?;
        for identifier in [
            TABLE_CALCULATION_ENGINE_OBJECT_ID,
            TABLE_FUNCTION_BROWSER_STATE_OBJECT_ID,
            TABLE_CUSTOM_FORMAT_LIST_OBJECT_ID,
        ] {
            append_object_reference(document, identifier);
        }

        let theme = archive
            .object_mut(PagesObjectId::Theme.value())
            .ok_or_else(|| {
                crate::Error::InvalidFormat("scratch Pages theme is missing".to_owned())
            })?;
        let theme_message_index = theme
            .messages
            .iter()
            .position(|message| message.type_ == PagesMessageType::Theme.value())
            .ok_or_else(|| {
                crate::Error::InvalidFormat("scratch Pages theme payload is missing".to_owned())
            })?;
        let theme_message_type = theme.messages[theme_message_index].type_;
        let mut decoded =
            IWorkThemeArchive::decode(theme.messages[theme_message_index].data.as_slice())?;
        decoded.extensions.table = Some(tst::ThemePresetsArchive {
            table_style_presets: vec![raw_reference(TABLE_PRESET_OBJECT_ID)],
            ..Default::default()
        });
        theme.replace_message(
            theme_message_index,
            RawMessage {
                type_: theme_message_type,
                data: decoded.encode()?,
            },
        )?;
        append_object_reference(theme, TABLE_PRESET_OBJECT_ID);

        archive.insert_object(raw_object_with_id(
            TABLE_ATTACHMENT_OBJECT_ID,
            TABLE_ATTACHMENT_MESSAGE_TYPE,
            tswp::DrawableAttachmentArchive {
                drawable: Some(raw_reference(TABLE_INFO_OBJECT_ID)),
                h_offset_type: Some(TABLE_ATTACHMENT_OFFSET_TYPE),
                h_offset: Some(TABLE_ATTACHMENT_OFFSET_POINTS),
                v_offset_type: Some(TABLE_ATTACHMENT_OFFSET_TYPE),
                v_offset: Some(TABLE_ATTACHMENT_OFFSET_POINTS),
            }
            .encode_to_vec(),
            &[TABLE_INFO_OBJECT_ID],
        )?)?;
        Ok(())
    })
}

fn register_table_components(package: &mut IWorkPackage) -> Result<()> {
    let document_component = PagesObjectId::Document.value();
    let stylesheet_component = PagesObjectId::Stylesheet.value();
    add_component_object_uuids(
        package,
        document_component,
        &TABLE_DOCUMENT_OBJECT_IDS
            .iter()
            .copied()
            .chain(std::iter::once(TABLE_ATTACHMENT_OBJECT_ID))
            .collect::<Vec<_>>(),
    )?;
    add_component_object_uuids(package, stylesheet_component, TABLE_STYLE_OBJECT_IDS)?;

    let mut calculation = component_raw(
        TABLE_CALCULATION_ENGINE_OBJECT_ID,
        "CalculationEngine",
        &TABLE_CALCULATION_COMPONENT_VERSION,
    );
    calculation.object_uuid_map_entries = [
        TABLE_CALCULATION_ENGINE_OBJECT_ID,
        TABLE_FORMULA_OWNER_OBJECT_ID,
    ]
    .into_iter()
    .map(object_uuid_raw)
    .collect();
    calculation.external_references = vec![tsp::ComponentExternalReference {
        component_identifier: document_component,
        object_identifier: Some(TABLE_INFO_OBJECT_ID),
        is_weak: None,
    }];
    add_component_registration(package, &calculation)?;

    for &identifier in TABLE_STYLE_OBJECT_IDS {
        add_component_external_reference(
            package,
            document_component,
            stylesheet_component,
            identifier,
        )?;
    }
    add_component_link(
        package,
        document_component,
        TABLE_CALCULATION_ENGINE_OBJECT_ID,
    )?;
    Ok(())
}
