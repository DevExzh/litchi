//! Exact-source integration coverage for Pages body-table appearance.
//!
//! The body graph is shared with the neighboring header-settings fixture so
//! selector and attachment semantics stay identical.  This file adds the
//! table-style graph (direct style, preset/network, and inheritance) and
//! keeps every malformed case source-bound: a rejected read or edit must not
//! publish bytes.

use std::error::Error as StdError;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field};
use litchi_iwa_core::{Archive, SnappyStream};
use litchi_iwa_protos::{tsp, tst};
use litchi_pages::table::appearance::{
    Appearance, Banding, GridlineVisibility, Gridlines, RowSizing,
};
use litchi_pages::{BodyTableAppearanceError as Error, BodyTableSelector, Package};
use prost::Message as _;

#[path = "body_table_header_settings.rs"]
mod header_fixture;

mod body_fixture {
    use super::header_fixture::{object, reference, rewrite_document_archive, synthetic_package};
    use super::{StdError, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE};
    use litchi_iwa_archive::{
        Limits,
        package::{Catalog, EntryEdit},
    };
    use litchi_iwa_common::wire::append_varint_field;
    use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
    use litchi_iwa_protos::{tsp, tss, tst};
    use prost::Message as _;

    const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
    const FIRST_DRAWABLE_IDENTIFIER: u64 = 200;
    const FIRST_MODEL_IDENTIFIER: u64 = 300;
    const SECOND_MODEL_IDENTIFIER: u64 = 310;
    const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
    const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;

    const PARENT_STYLE_ID: u64 = 1_000;
    const CHILD_STYLE_ID: u64 = 1_001;
    const PRESET_ID: u64 = 1_010;
    const NETWORK_ID: u64 = 1_011;
    const STYLESHEET_ID: u64 = 1_012;
    const GENERIC_STYLE_BASE: u64 = 1_020;
    const STYLE_TYPE: u32 = 6_003;
    const PRESET_TYPE: u32 = 6_008;
    const NETWORK_TYPE: u32 = 6_247;
    const STYLESHEET_TYPE: u32 = 401;
    const METADATA_MEMBER: &str = "Index/Metadata.iwa";
    const METADATA_OBJECT_ID: u64 = 50_000;
    const METADATA_MESSAGE_TYPE: u32 = 11_006;
    const UNKNOWN_STYLE_FIELD: u32 = 90;
    const UNKNOWN_STYLE_VALUE: u64 = 0x0bad_cafe;

    fn properties(enabled: bool) -> tst::TableStylePropertiesArchive {
        tst::TableStylePropertiesArchive {
            banded_rows: Some(enabled),
            auto_resize: Some(enabled),
            h_strokes_visible: Some(!enabled),
            v_strokes_visible: Some(!enabled),
            table_hc_divider_visible: Some(enabled),
            table_hr_divider_visible: Some(!enabled),
            table_footer_divider_visible: Some(enabled),
            ..tst::TableStylePropertiesArchive::default()
        }
    }

    fn style_payload(
        identifier: u64,
        parent: Option<u64>,
        enabled: bool,
    ) -> Result<Vec<u8>, Box<dyn StdError>> {
        let mut payload = tst::TableStyleArchive {
            super_: tss::StyleArchive {
                name: Some(format!("Pages appearance style {identifier}")),
                style_identifier: Some(format!("pages-appearance-{identifier}")),
                parent: parent.map(reference),
                is_variation: Some(parent.is_some()),
                stylesheet: Some(reference(STYLESHEET_ID)),
            },
            override_count: Some(if parent.is_some() { 7 } else { 0 }),
            table_properties: Some(properties(enabled)),
        }
        .encode_to_vec();
        append_varint_field(&mut payload, UNKNOWN_STYLE_FIELD, UNKNOWN_STYLE_VALUE)?;
        Ok(payload)
    }

    fn style_object(
        identifier: u64,
        parent: Option<u64>,
        enabled: bool,
    ) -> Result<ArchiveObject, Box<dyn StdError>> {
        let mut object = object(
            identifier,
            STYLE_TYPE,
            style_payload(identifier, parent, enabled)?,
            &[STYLESHEET_ID],
        )?;
        let info = &mut object.archive_info.message_infos[0];
        info.object_references = vec![STYLESHEET_ID];
        let mut field = FieldInfo::new(vec![5]);
        field.object_references.push(STYLESHEET_ID);
        info.field_infos.push(field);
        Ok(object)
    }

    fn placeholder_style(identifier: u64) -> Result<ArchiveObject, Box<dyn StdError>> {
        object(
            identifier,
            2_022,
            tss::StyleArchive {
                name: Some(format!("Pages placeholder {identifier}")),
                style_identifier: Some(format!("pages-placeholder-{identifier}")),
                ..tss::StyleArchive::default()
            }
            .encode_to_vec(),
            &[],
        )
    }

    fn stylesheet_object(style_ids: &[u64]) -> Result<ArchiveObject, Box<dyn StdError>> {
        let mut payload = tss::StylesheetArchive {
            styles: style_ids.iter().copied().map(reference).collect(),
            identifier_to_style_map: style_ids
                .iter()
                .map(|identifier| tss::stylesheet_archive::IdentifiedStyleEntry {
                    identifier: format!("pages-style-{identifier}"),
                    style: reference(*identifier),
                })
                .collect(),
            is_locked: Some(false),
            can_cull_styles: Some(true),
            ..tss::StylesheetArchive::default()
        }
        .encode_to_vec();
        append_varint_field(&mut payload, 91, 0x1234)?;
        let mut object = object(STYLESHEET_ID, STYLESHEET_TYPE, payload, &[])?;
        object.archive_info.message_infos[0]
            .object_references
            .extend_from_slice(style_ids);
        Ok(object)
    }

    fn preset_objects() -> Result<Vec<ArchiveObject>, Box<dyn StdError>> {
        let generic_ids: Vec<u64> = (0..9).map(|index| GENERIC_STYLE_BASE + index).collect();
        let network = tst::TableStyleNetworkArchive {
            body_text_style: reference(generic_ids[0]),
            header_row_text_style: reference(generic_ids[1]),
            header_column_text_style: reference(generic_ids[2]),
            footer_row_text_style: reference(generic_ids[3]),
            body_cell_style: reference(generic_ids[4]),
            header_row_style: reference(generic_ids[5]),
            header_column_style: reference(generic_ids[6]),
            footer_row_style: reference(generic_ids[7]),
            table_style: reference(PARENT_STYLE_ID),
            ..tst::TableStyleNetworkArchive::default()
        }
        .encode_to_vec();
        let preset = tst::TableStylePresetArchive {
            index: Some(3),
            style_network: Some(reference(NETWORK_ID)),
            ..tst::TableStylePresetArchive::default()
        }
        .encode_to_vec();
        let mut network_object = object(NETWORK_ID, NETWORK_TYPE, network, &[])?;
        network_object.archive_info.message_infos[0]
            .object_references
            .extend(generic_ids.iter().copied());
        network_object.archive_info.message_infos[0]
            .object_references
            .push(PARENT_STYLE_ID);
        let preset_object = object(PRESET_ID, PRESET_TYPE, preset, &[NETWORK_ID])?;
        let mut objects = vec![preset_object, network_object];
        objects.extend(
            generic_ids
                .iter()
                .copied()
                .map(placeholder_style)
                .collect::<Result<Vec<_>, _>>()?,
        );
        Ok(objects)
    }

    fn rewrite_model_style(
        source: &[u8],
        first_style: u64,
        second_style: u64,
        first_preset: Option<u64>,
    ) -> Result<Vec<u8>, Box<dyn StdError>> {
        rewrite_document_archive(source, |archive| {
            for (identifier, style, preset) in [
                (FIRST_MODEL_IDENTIFIER, first_style, first_preset),
                (SECOND_MODEL_IDENTIFIER, second_style, None),
            ] {
                let model = archive
                    .object_mut(identifier)
                    .ok_or("missing table model")?;
                let message_index = model
                    .messages
                    .iter()
                    .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
                    .ok_or("missing table model message")?;
                let message_type = model.messages[message_index].type_;
                let mut native =
                    tst::TableModelArchive::decode(model.messages[message_index].data.as_slice())?;
                native.table_style = reference(style);
                native.table_style_preset = preset.map(reference);
                let mut data = native.encode_to_vec();
                append_varint_field(&mut data, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
                model.replace_message_preserving_header(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
                let info = &mut model.archive_info.message_infos[0];
                info.object_references.clear();
                if style != 0 {
                    info.object_references.push(style);
                }
                if let Some(preset) = preset {
                    info.object_references.push(preset);
                }
                info.field_infos.clear();
                if style != 0 {
                    let mut field = FieldInfo::new(vec![3]);
                    field.object_references.push(style);
                    info.field_infos.push(field);
                }
            }
            archive
                .objects
                .push(style_object(PARENT_STYLE_ID, None, false)?);
            if first_style == CHILD_STYLE_ID {
                archive
                    .objects
                    .push(style_object(CHILD_STYLE_ID, Some(PARENT_STYLE_ID), true)?);
            }
            let mut style_ids = vec![PARENT_STYLE_ID];
            if first_style == CHILD_STYLE_ID {
                style_ids.push(CHILD_STYLE_ID);
            }
            archive.objects.push(stylesheet_object(&style_ids)?);
            if first_preset.is_some() {
                archive.objects.extend(preset_objects()?);
            }
            Ok(())
        })
    }

    fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
        tsp::ObjectUuidMapEntry {
            identifier,
            uuid: tsp::Uuid {
                lower: identifier.saturating_add(10_000),
                upper: identifier.saturating_add(20_000),
            },
        }
    }

    fn metadata_package(source: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        let source_catalog = Catalog::from_bytes(source)?;
        let document = source_catalog
            .iter()
            .find(|entry| entry.name() == DOCUMENT_MEMBER)
            .ok_or("missing document member")?;
        let document_archive =
            Archive::parse(SnappyStream::decompress(document.data())?.as_bytes())?;
        let mut identifiers = document_archive
            .objects
            .iter()
            .map(|object| {
                object
                    .archive_info
                    .identifier
                    .ok_or("missing object identifier")
            })
            .collect::<Result<Vec<_>, _>>()?;
        identifiers.push(METADATA_OBJECT_ID);
        identifiers.sort_unstable();
        let metadata = tsp::PackageMetadata {
            last_object_identifier: METADATA_OBJECT_ID,
            save_token: Some(1),
            components: vec![tsp::ComponentInfo {
                identifier: 1,
                preferred_locator: "Document".to_owned(),
                locator: Some("Document".to_owned()),
                save_token: Some(1),
                object_uuid_map_entries: identifiers.iter().copied().map(uuid_entry).collect(),
                ..tsp::ComponentInfo::default()
            }],
            ..tsp::PackageMetadata::default()
        }
        .encode_to_vec();
        let metadata_archive = SnappyStream::compress(
            &Archive {
                objects: vec![object(METADATA_OBJECT_ID, 11_006, metadata, &[])?],
            }
            .to_bytes()?,
        )?;
        let catalog = Catalog::from_bytes(source)?;
        let mut members = catalog
            .iter()
            .filter(|entry| entry.name() != METADATA_MEMBER)
            .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
            .collect::<Vec<_>>();
        members.push((METADATA_MEMBER.to_owned(), metadata_archive));
        let refs = members
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice()))
            .collect::<Vec<_>>();
        Ok(litchi_iwa_archive::package::to_bytes(
            refs,
            Limits::default(),
        )?)
    }

    fn rewrite_metadata(
        source: &[u8],
        mutate: impl FnOnce(&mut tsp::PackageMetadata) -> Result<(), Box<dyn StdError>>,
    ) -> Result<Vec<u8>, Box<dyn StdError>> {
        let catalog = Catalog::from_bytes(source)?;
        let entry = catalog
            .iter()
            .find(|entry| entry.name() == METADATA_MEMBER)
            .ok_or("missing metadata member")?;
        let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
        let metadata_object = archive
            .object_mut(METADATA_OBJECT_ID)
            .ok_or("missing metadata object")?;
        let message_index = metadata_object
            .messages
            .iter()
            .position(|message| message.type_ == METADATA_MESSAGE_TYPE)
            .ok_or("missing metadata message")?;
        let message_type = metadata_object.messages[message_index].type_;
        let mut metadata =
            tsp::PackageMetadata::decode(metadata_object.messages[message_index].data.as_slice())?;
        mutate(&mut metadata)?;
        metadata_object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: message_type,
                data: metadata.encode_to_vec(),
            },
        )?;
        let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
        Ok(catalog.reassemble_to_bytes(
            &[EntryEdit::new(METADATA_MEMBER, compressed.as_slice())],
            Limits::default(),
        )?)
    }

    fn append_metadata_unknown(source: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        let catalog = Catalog::from_bytes(source)?;
        let entry = catalog
            .iter()
            .find(|entry| entry.name() == METADATA_MEMBER)
            .ok_or("missing metadata member")?;
        let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
        let metadata_object = archive
            .object_mut(METADATA_OBJECT_ID)
            .ok_or("missing metadata object")?;
        let message_index = metadata_object
            .messages
            .iter()
            .position(|message| message.type_ == METADATA_MESSAGE_TYPE)
            .ok_or("missing metadata message")?;
        let message_type = metadata_object.messages[message_index].type_;
        let mut data = metadata_object.messages[message_index].data.clone();
        append_varint_field(&mut data, 90, 0xdecafbad)?;
        metadata_object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
        Ok(catalog.reassemble_to_bytes(
            &[EntryEdit::new(METADATA_MEMBER, compressed.as_slice())],
            Limits::default(),
        )?)
    }

    pub(crate) fn without_metadata_member(source: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        Ok(
            Catalog::from_bytes(source)?.reassemble_with_deletions_to_bytes(
                &[],
                &[METADATA_MEMBER],
                Limits::default(),
            )?,
        )
    }

    pub(crate) fn split_style_component(source: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        const FOREIGN_MEMBER: &str = "Index/ForeignStyles.iwa";
        let catalog = Catalog::from_bytes(source)?;
        let document = catalog
            .iter()
            .find(|entry| entry.name() == DOCUMENT_MEMBER)
            .ok_or("missing document member")?;
        let mut archive = Archive::parse(SnappyStream::decompress(document.data())?.as_bytes())?;
        let mut kept = Vec::new();
        let mut moved = Vec::new();
        for object in archive.objects {
            let is_style = object
                .messages
                .iter()
                .any(|message| message.type_ == STYLE_TYPE || message.type_ == STYLESHEET_TYPE);
            if is_style {
                moved.push(object);
            } else {
                kept.push(object);
            }
        }
        archive.objects = kept;
        let document_data = SnappyStream::compress(&archive.to_bytes()?)?;
        let foreign_data = SnappyStream::compress(&Archive { objects: moved }.to_bytes()?)?;
        let mut members = catalog
            .iter()
            .map(|entry| {
                if entry.name() == DOCUMENT_MEMBER {
                    (entry.name().to_owned(), document_data.to_vec())
                } else {
                    (entry.name().to_owned(), entry.data().to_vec())
                }
            })
            .collect::<Vec<_>>();
        members.push((FOREIGN_MEMBER.to_owned(), foreign_data));
        let refs = members
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice()))
            .collect::<Vec<_>>();
        Ok(litchi_iwa_archive::package::to_bytes(
            refs,
            Limits::default(),
        )?)
    }

    pub(crate) fn metadata_hostiles(source: &[u8]) -> Result<Vec<Vec<u8>>, Box<dyn StdError>> {
        let mut cases = Vec::new();
        cases.push(rewrite_metadata(source, |metadata| {
            let component = metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?;
            component
                .object_uuid_map_entries
                .push(uuid_entry(PARENT_STYLE_ID));
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            let component = metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?;
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != FIRST_MODEL_IDENTIFIER);
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            metadata.versioned_components.push(tsp::ComponentInfo {
                identifier: 1,
                preferred_locator: "Document".to_owned(),
                locator: Some("Document".to_owned()),
                object_uuid_map_entries: vec![uuid_entry(FIRST_MODEL_IDENTIFIER)],
                ..tsp::ComponentInfo::default()
            });
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            let component = metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?;
            component.data_references.push(tsp::ComponentDataReference {
                data_identifier: PARENT_STYLE_ID,
                object_reference_list: Vec::new(),
            });
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            let component = metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?;
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != PARENT_STYLE_ID);
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            metadata.versioned_components.push(tsp::ComponentInfo {
                identifier: 1,
                preferred_locator: "Document".to_owned(),
                locator: Some("Document".to_owned()),
                object_uuid_map_entries: vec![uuid_entry(PARENT_STYLE_ID)],
                ..tsp::ComponentInfo::default()
            });
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            let component = metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?;
            component.ambiguous_object_identifiers.push(PARENT_STYLE_ID);
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            let component = metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?;
            component.data_references.push(tsp::ComponentDataReference {
                data_identifier: 77,
                object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                    object_identifier: PARENT_STYLE_ID,
                    count: 1,
                }],
            });
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            metadata.data_metadata_map = Some(reference(PARENT_STYLE_ID));
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            metadata.components.push(tsp::ComponentInfo {
                identifier: 2,
                preferred_locator: "Foreign".to_owned(),
                locator: Some("Foreign".to_owned()),
                external_references: vec![tsp::ComponentExternalReference {
                    component_identifier: 1,
                    object_identifier: Some(PARENT_STYLE_ID),
                    ..tsp::ComponentExternalReference::default()
                }],
                ..tsp::ComponentInfo::default()
            });
            Ok(())
        })?);
        cases.push(rewrite_metadata(source, |metadata| {
            metadata.components.clear();
            Ok(())
        })?);
        cases.push(append_metadata_unknown(source)?);
        Ok(cases)
    }

    pub(crate) fn direct_package() -> Result<Vec<u8>, Box<dyn StdError>> {
        let source = synthetic_package(["Revenue", "Costs"], None)?;
        metadata_package(&rewrite_model_style(
            &source,
            PARENT_STYLE_ID,
            PARENT_STYLE_ID,
            None,
        )?)
    }

    pub(crate) fn child_package() -> Result<Vec<u8>, Box<dyn StdError>> {
        let source = synthetic_package(["Revenue", "Costs"], None)?;
        metadata_package(&rewrite_model_style(
            &source,
            CHILD_STYLE_ID,
            PARENT_STYLE_ID,
            None,
        )?)
    }

    pub(crate) fn preset_package() -> Result<Vec<u8>, Box<dyn StdError>> {
        let source = synthetic_package(["Revenue", "Costs"], None)?;
        metadata_package(&rewrite_model_style(
            &source,
            0,
            PARENT_STYLE_ID,
            Some(PRESET_ID),
        )?)
    }

    pub(crate) fn default_package() -> Result<Vec<u8>, Box<dyn StdError>> {
        let source = synthetic_package(["Revenue", "Costs"], None)?;
        metadata_package(&rewrite_model_style(&source, 0, PARENT_STYLE_ID, None)?)
    }

    pub(crate) fn duplicate_name_package() -> Result<Vec<u8>, Box<dyn StdError>> {
        let source = synthetic_package(["Revenue", "Revenue"], None)?;
        metadata_package(&rewrite_model_style(
            &source,
            PARENT_STYLE_ID,
            PARENT_STYLE_ID,
            None,
        )?)
    }

    pub(crate) fn rewrite_archive(
        source: &[u8],
        mutate: impl FnOnce(&mut Archive) -> Result<(), Box<dyn StdError>>,
    ) -> Result<Vec<u8>, Box<dyn StdError>> {
        rewrite_document_archive(source, mutate)
    }

    pub(crate) fn malformed_model(source: &[u8], raw: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        rewrite_document_archive(source, |archive| {
            let model = archive
                .object_mut(FIRST_MODEL_IDENTIFIER)
                .ok_or("missing table model")?;
            let message_index = model
                .messages
                .iter()
                .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
                .ok_or("missing table model message")?;
            let mut data = model.messages[message_index].data.clone();
            data.extend_from_slice(raw);
            model.replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: TABLE_MODEL_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
    }

    pub(crate) fn remove_style(source: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        rewrite_document_archive(source, |archive| {
            archive
                .objects
                .retain(|object| object.archive_info.identifier != Some(PARENT_STYLE_ID));
            Ok(())
        })
    }

    pub(crate) fn corrupt_archive_info(source: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        rewrite_document_archive(source, |archive| {
            let model = archive
                .object_mut(FIRST_MODEL_IDENTIFIER)
                .ok_or("missing table model")?;
            model.archive_info.message_infos[0]
                .object_references
                .push(999_999);
            Ok(())
        })
    }

    pub(crate) fn duplicate_model_identifier(source: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        rewrite_document_archive(source, |archive| {
            let second = archive
                .object_mut(SECOND_MODEL_IDENTIFIER)
                .ok_or("missing second model")?;
            second.archive_info.identifier = Some(FIRST_MODEL_IDENTIFIER);
            Ok(())
        })
    }

    pub(crate) fn locked_source(source: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        rewrite_document_archive(source, |archive| {
            let info = archive
                .object_mut(FIRST_DRAWABLE_IDENTIFIER)
                .ok_or("missing table info")?;
            let message_index = info
                .messages
                .iter()
                .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
                .ok_or("missing table info message")?;
            let mut native =
                tst::TableInfoArchive::decode(info.messages[message_index].data.as_slice())?;
            native.super_.locked = Some(true);
            info.replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: TABLE_INFO_MESSAGE_TYPE,
                    data: native.encode_to_vec(),
                },
            )?;
            Ok(())
        })
    }

    pub(crate) fn add_cross_component_inbound(source: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
        let catalog = Catalog::from_bytes(source)?;
        let inbound = object(7_000, 7_001, vec![0x08, 0x01], &[PARENT_STYLE_ID])?;
        let component = SnappyStream::compress(
            &Archive {
                objects: vec![inbound],
            }
            .to_bytes()?,
        )?;
        let mut members = catalog
            .iter()
            .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
            .collect::<Vec<_>>();
        members.push(("Index/Foreign.iwa".to_owned(), component));
        let refs = members
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice()))
            .collect::<Vec<_>>();
        Ok(litchi_iwa_archive::package::to_bytes(
            refs,
            Limits::default(),
        )?)
    }
}

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const FIRST_MODEL_IDENTIFIER: u64 = 300;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const UNKNOWN_MODEL_FIELD: u32 = 99;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts package bytes");
        bytes
    }
}

fn expected_direct() -> Appearance {
    Appearance {
        row_banding: Banding::Disabled,
        row_sizing: RowSizing::Fixed,
        gridlines: Gridlines {
            body_horizontal: GridlineVisibility::Visible,
            header_columns_horizontal: GridlineVisibility::Hidden,
            body_vertical: GridlineVisibility::Visible,
            header_rows_vertical: GridlineVisibility::Visible,
            footer_rows_vertical: GridlineVisibility::Hidden,
        },
    }
}

fn expected_child() -> Appearance {
    Appearance {
        row_banding: Banding::Enabled,
        row_sizing: RowSizing::FitCellContents,
        gridlines: Gridlines {
            body_horizontal: GridlineVisibility::Hidden,
            header_columns_horizontal: GridlineVisibility::Visible,
            body_vertical: GridlineVisibility::Hidden,
            header_rows_vertical: GridlineVisibility::Hidden,
            footer_rows_vertical: GridlineVisibility::Visible,
        },
    }
}

fn member_bytes(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| format!("missing package member {name}"))?
        .data()
        .to_vec())
}

fn document_archive(package: &[u8]) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn model_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let archive = document_archive(package)?;
    Ok(archive
        .object(identifier)
        .ok_or("missing model object")?
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or("missing model message")?
        .data
        .clone())
}

fn style_payloads(package: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let archive = document_archive(package)?;
    Ok(archive
        .objects
        .iter()
        .flat_map(|object| {
            object
                .messages
                .iter()
                .filter(|message| message.type_ == 6_003)
                .map(|message| message.data.clone())
        })
        .collect())
}

fn assert_rejected(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        return Ok(());
    };
    let before = package.exact_bytes();
    assert!(package.body_table_appearance(0usize).is_err());
    assert!(package.edit_body_table_appearance(0usize).is_err());
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn changed_appearance() -> Appearance {
    Appearance {
        row_banding: Banding::Enabled,
        row_sizing: RowSizing::FitCellContents,
        gridlines: Gridlines {
            body_horizontal: GridlineVisibility::Hidden,
            header_columns_horizontal: GridlineVisibility::Hidden,
            body_vertical: GridlineVisibility::Visible,
            header_rows_vertical: GridlineVisibility::Hidden,
            footer_rows_vertical: GridlineVisibility::Hidden,
        },
    }
}

#[test]
fn selectors_and_direct_preset_default_resolution_are_typed() -> TestResult {
    let direct = Package::from_bytes(&body_fixture::direct_package()?)?;
    assert_eq!(
        direct.body_table_appearance(BodyTableSelector::index(0))?,
        direct.body_table_appearance(BodyTableSelector::name("Revenue"))?
    );
    assert_eq!(direct.body_table_appearance(0usize)?, expected_direct());
    assert!(
        direct
            .body_table_appearance(BodyTableSelector::name("Missing"))
            .is_err()
    );

    let ambiguous = Package::from_bytes(&body_fixture::duplicate_name_package()?)?;
    assert!(
        ambiguous
            .body_table_appearance(BodyTableSelector::name("Revenue"))
            .is_err()
    );
    assert_eq!(
        Package::from_bytes(&body_fixture::preset_package()?)?.body_table_appearance(0usize)?,
        expected_direct()
    );
    assert_eq!(
        Package::from_bytes(&body_fixture::default_package()?)?.body_table_appearance(0usize)?,
        Appearance::default()
    );
    Ok(())
}

#[test]
fn inherited_child_overrides_parent_and_full_chain_is_read() -> TestResult {
    let package = Package::from_bytes(&body_fixture::child_package()?)?;
    assert_eq!(package.body_table_appearance(0usize)?, expected_child());
    assert_eq!(package.body_table_appearance(1usize)?, expected_direct());
    Ok(())
}

#[test]
fn no_op_is_exact_and_apply_is_idempotence_checked() -> TestResult {
    let source = body_fixture::direct_package()?;
    let package = Package::from_bytes(&source)?;
    let before = package.body_table_appearance(0usize)?;
    let commit = package
        .edit_body_table_appearance(BodyTableSelector::name("Revenue"))?
        .set(before)
        .commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert_eq!(commit.package().exact_bytes(), source);
    assert_eq!(
        package
            .apply_body_table_appearance(commit.patch())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn exact_no_op_skips_changed_target_resolution_but_changes_fail_closed() -> TestResult {
    let source = body_fixture::direct_package()?;
    // This profile admits the staged target resolution but not the second
    // changed-path resolution, providing a deterministic no-op ordering gate.
    let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(75)?;
    let limits = Limits::default().with_archive_limits(archive_limits)?;
    let package = Package::from_bytes_with_limits(&source, limits)?;
    let edit = package.edit_body_table_appearance(0usize)?;
    let before = package.exact_bytes();
    let appearance = edit.appearance();
    let noop = edit.set(appearance).commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(
        noop.patch().source_fingerprint(),
        noop.patch().target_fingerprint()
    );
    assert_eq!(noop.package().exact_bytes(), before);
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());

    let changed_edit = package.edit_body_table_appearance(0usize)?;
    let changed = changed_edit.set(changed_appearance()).commit();
    assert!(matches!(changed, Err(Error::LimitExceeded { .. })));
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[cfg(feature = "internal-iwork-source")]
#[test]
fn apply_rejects_semantic_prepared_source_but_keeps_noop_exact() -> TestResult {
    use std::sync::Arc;

    use litchi_iwa_detect::PreparedSource;

    let source: Arc<[u8]> = body_fixture::direct_package()?.into();
    let exact_prepared = PreparedSource::from_shared_bytes(Arc::clone(&source))?
        .ok_or("fixture was not detected as Pages")?;
    let exact_package = Package::__from_prepared_source(exact_prepared)?;
    let semantic_prepared = PreparedSource::__from_shared_bytes_with_pages_metadata(
        Arc::clone(&source),
        litchi_iwa_detect::Limits::default(),
    )?
    .ok_or("fixture was not detected as Pages")?;
    let semantic_package = Package::__from_prepared_source(semantic_prepared)?;
    let before = semantic_package.exact_bytes();

    let noop = semantic_package
        .edit_body_table_appearance(0usize)?
        .set(semantic_package.body_table_appearance(0usize)?)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), before);
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(
        semantic_package
            .apply_body_table_appearance(noop.patch())?
            .package()
            .exact_bytes(),
        before
    );

    let changed = exact_package
        .edit_body_table_appearance(0usize)?
        .set(changed_appearance())
        .commit()?;
    assert!(!changed.patch().is_noop());
    assert!(matches!(
        semantic_package.apply_body_table_appearance(changed.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(semantic_package.exact_bytes(), before);
    Ok(())
}

#[test]
fn changed_appearance_cows_shared_style_preserves_unknowns_and_locality() -> TestResult {
    let source = body_fixture::direct_package()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_body_table_appearance(0usize)?
        .set(changed_appearance())
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_eq!(
        commit.package().body_table_appearance(0usize)?,
        changed_appearance()
    );
    assert_eq!(
        commit.package().body_table_appearance(1usize)?,
        expected_direct()
    );
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    assert_eq!(
        member_bytes(&source, "Data/sentinel.bin")?,
        b"untouched-sentinel"
    );
    assert_eq!(
        member_bytes(&target, "Data/sentinel.bin")?,
        b"untouched-sentinel"
    );
    assert_ne!(source, target);
    let model = model_payload(&target, FIRST_MODEL_IDENTIFIER)?;
    assert!(
        WireView::parse(&model)?
            .fields()
            .any(|field| field.number() == UNKNOWN_MODEL_FIELD)
    );
    assert!(style_payloads(&target)?.iter().any(|payload| {
        WireView::parse(payload)
            .map(|view| view.fields().any(|field| field.number() == 90))
            .unwrap_or(false)
    }));
    let restored = commit
        .package()
        .apply_body_table_appearance(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);
    Ok(())
}

#[test]
fn apply_conflict_and_reopen_are_exact_source_bound() -> TestResult {
    let source = body_fixture::direct_package()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_body_table_appearance(0usize)?
        .set(changed_appearance())
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_eq!(
        Package::from_bytes(&target)?.body_table_appearance(0usize)?,
        changed_appearance()
    );
    assert_eq!(
        package
            .apply_body_table_appearance(commit.patch())?
            .package()
            .exact_bytes(),
        target
    );
    assert!(matches!(
        commit.package().apply_body_table_appearance(commit.patch()),
        Err(Error::PatchConflict)
    ));
    let restored = commit
        .package()
        .apply_body_table_appearance(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);
    Ok(())
}

#[test]
fn locked_table_allows_exact_noop_but_refuses_change_atomically() -> TestResult {
    let source = body_fixture::locked_source(&body_fixture::direct_package()?)?;
    let package = Package::from_bytes(&source)?;
    let before = package.exact_bytes();
    let noop = package
        .edit_body_table_appearance(0usize)?
        .set(package.body_table_appearance(0usize)?)
        .commit()?;
    assert!(noop.patch().is_noop());
    let error = package
        .edit_body_table_appearance(0usize)?
        .set(changed_appearance())
        .commit()
        .expect_err("locked appearance must reject changes");
    assert!(matches!(
        error,
        Error::TableLocked { .. } | Error::UnsupportedDependency { .. }
    ));
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn malformed_unknown_duplicate_wrong_wire_and_noncanonical_fields_fail_closed() -> TestResult {
    let source = body_fixture::direct_package()?;
    let mut duplicate_known = Vec::new();
    append_length_delimited_field(
        &mut duplicate_known,
        3,
        &tsp::Reference {
            identifier: 1_000,
            ..tsp::Reference::default()
        }
        .encode_to_vec(),
    )?;
    let cases = [
        ("duplicate known", duplicate_known),
        ("wrong wire", vec![0x1d, 0, 0, 0, 0]),
        ("noncanonical reference", vec![0x1a, 0x03, 0x08, 0x80, 0x00]),
        ("unterminated group", vec![0xa3, 0x06, 0x08, 0x01]),
    ];
    for (name, raw) in cases {
        let malformed = body_fixture::malformed_model(&source, &raw)?;
        assert_rejected(&malformed).map_err(|error| format!("{name}: {error}"))?;
    }
    Ok(())
}

#[test]
fn missing_style_archive_info_duplicate_identity_and_foreign_inbound_are_atomic() -> TestResult {
    let source = body_fixture::direct_package()?;
    assert_rejected(&body_fixture::remove_style(&source)?)?;
    assert_rejected(&body_fixture::split_style_component(&source)?)?;
    assert_rejected(&body_fixture::corrupt_archive_info(&source)?)?;
    if let Ok(duplicate) = body_fixture::duplicate_model_identifier(&source) {
        assert_rejected(&duplicate)?;
    }
    let inbound = body_fixture::add_cross_component_inbound(&source)?;
    assert_rejected(&inbound)?;
    let duplicate_style_field = body_fixture::rewrite_archive(&source, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        let field = model.archive_info.message_infos[0]
            .field_infos
            .first()
            .cloned()
            .ok_or("missing style FieldInfo")?;
        model.archive_info.message_infos[0].field_infos.push(field);
        Ok(())
    })?;
    assert_rejected(&duplicate_style_field)?;
    let wrong_style_field_type = body_fixture::rewrite_archive(&source, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        model.archive_info.message_infos[0].field_infos[0].r#type =
            Some(litchi_iwa_core::FieldType::DataReference);
        Ok(())
    })?;
    assert_rejected(&wrong_style_field_type)?;
    let model_data_reference = body_fixture::rewrite_archive(&source, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        model.archive_info.message_infos[0]
            .data_references
            .push(1_000);
        Ok(())
    })?;
    assert_rejected(&model_data_reference)?;
    Ok(())
}

#[test]
fn metadata_authority_collisions_and_opaque_inbound_are_atomic() -> TestResult {
    let source = body_fixture::direct_package()?;
    assert_rejected(&body_fixture::without_metadata_member(&source)?)?;
    for hostile in body_fixture::metadata_hostiles(&source)? {
        assert_rejected(&hostile)?;
    }
    Ok(())
}

#[test]
fn missing_or_cyclic_ancestors_fail_even_when_child_is_fully_specified() -> TestResult {
    let source = body_fixture::child_package()?;
    let missing_parent = body_fixture::remove_style(&source)?;
    assert_rejected(&missing_parent)?;

    let cycle = body_fixture::rewrite_archive(&source, |archive| {
        let style = archive.object_mut(1_001).ok_or("missing child style")?;
        let message = style
            .messages
            .iter_mut()
            .find(|message| message.type_ == 6_003)
            .ok_or("missing child style message")?;
        let mut native = tst::TableStyleArchive::decode(message.data.as_slice())?;
        native.super_.parent = Some(tsp::Reference {
            identifier: 1_001,
            ..tsp::Reference::default()
        });
        message.data = native.encode_to_vec();
        Ok(())
    })?;
    assert_rejected(&cycle)?;
    Ok(())
}

#[test]
fn finite_ingress_and_semantic_limits_refuse_before_publication() -> TestResult {
    let source = body_fixture::direct_package()?;
    let mut refused = false;
    for fields in 1..=4_096 {
        let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(fields)?;
        let limits = Limits::default().with_archive_limits(archive_limits)?;
        let Ok(package) = Package::from_bytes_with_limits(&source, limits) else {
            continue;
        };
        let before = package.exact_bytes();
        let result = package
            .edit_body_table_appearance(0usize)
            .and_then(|edit| edit.set(changed_appearance()).commit());
        if matches!(result, Err(Error::LimitExceeded { .. })) {
            assert_eq!(package.exact_bytes(), before);
            refused = true;
            break;
        }
    }
    assert!(
        refused,
        "an ingress-compatible finite profile must reject eventually"
    );
    Ok(())
}

#[test]
fn public_transaction_values_are_archive_free_send_sync_and_redacted() -> TestResult {
    fn assert_traits<T: Send + Sync + std::fmt::Debug>() {}

    assert_traits::<Appearance>();
    assert_traits::<Banding>();
    assert_traits::<GridlineVisibility>();
    assert_traits::<Gridlines>();
    assert_traits::<RowSizing>();
    assert_traits::<litchi_pages::BodyTableAppearanceEdit<'static>>();
    assert_traits::<litchi_pages::BodyTableAppearancePatch>();
    assert_traits::<litchi_pages::BodyTableAppearanceCommit>();
    assert_traits::<litchi_pages::BodyTableAppearanceDiagnostics>();
    assert_traits::<litchi_pages::BodyTableAppearanceError>();
    assert_traits::<litchi_pages::BodyTableAppearanceLimitKind>();
    assert_traits::<litchi_pages::BodyTableAppearancePath>();

    let package = Package::from_bytes(&body_fixture::direct_package()?)?;
    let edit = package.edit_body_table_appearance(0usize)?;
    let debug = format!("{edit:?}");
    assert!(debug.contains("appearance"));
    assert!(!debug.contains("1000"));
    Ok(())
}
