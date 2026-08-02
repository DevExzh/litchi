//! Typed, copy-on-write appearance controls shared by native iWork tables.

mod wire;

use std::collections::HashSet;

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    next_object_identifier, set_package_last_object_identifier,
};
use crate::protobuf::{tsp, tss, tst};
use crate::shapes::insert_style_variation;
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};
use wire::{TableAppearanceOverrides, table_appearance_overrides};

const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const TABLE_STYLE_REFERENCE_FIELD: u32 = 3;
const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

/// Whether native iWork applies its alternating fill to body rows.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum TableRowBanding {
    /// Use the same fill for adjacent body rows.
    #[default]
    Disabled,
    /// Apply the table style's alternating-row fill.
    Enabled,
}

impl TableRowBanding {
    fn from_native(value: bool) -> Self {
        if value { Self::Enabled } else { Self::Disabled }
    }

    fn native(self) -> bool {
        matches!(self, Self::Enabled)
    }
}

/// Whether row heights automatically expand to fit their cell contents.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum TableRowSizing {
    /// Preserve explicit or style-provided row heights.
    #[default]
    Fixed,
    /// Automatically expand rows to fit their cell contents.
    FitCellContents,
}

impl TableRowSizing {
    fn from_native(value: bool) -> Self {
        if value {
            Self::FitCellContents
        } else {
            Self::Fixed
        }
    }

    fn native(self) -> bool {
        matches!(self, Self::FitCellContents)
    }
}

/// Whether one family of native table gridlines is drawn.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub enum TableGridlineVisibility {
    /// Do not draw this gridline family.
    Hidden,
    /// Draw this gridline family using the table style's strokes.
    #[default]
    Visible,
}

impl TableGridlineVisibility {
    fn from_native(value: bool) -> Self {
        if value { Self::Visible } else { Self::Hidden }
    }

    fn native(self) -> bool {
        matches!(self, Self::Visible)
    }
}

/// Gridline visibility for each native iWork table region.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub struct TableGridlines {
    /// Horizontal lines between body rows, excluding header-column lines.
    pub body_horizontal: TableGridlineVisibility,
    /// Horizontal lines between rows inside the header-column region.
    pub header_columns_horizontal: TableGridlineVisibility,
    /// Vertical lines between body columns, excluding header-row and footer lines.
    pub body_vertical: TableGridlineVisibility,
    /// Vertical lines between columns inside the header-row region.
    pub header_rows_vertical: TableGridlineVisibility,
    /// Vertical lines between columns inside the footer-row region.
    pub footer_rows_vertical: TableGridlineVisibility,
}

/// Effective appearance settings backed by a native table-style inheritance chain.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub struct TableAppearance {
    /// Alternating body-row fill behavior.
    pub row_banding: TableRowBanding,
    /// Automatic row-height behavior.
    pub row_sizing: TableRowSizing,
    /// Horizontal and vertical gridline visibility by table region.
    pub gridlines: TableGridlines,
}

pub(crate) fn table_appearance(
    package: &IWorkPackage,
    model_object_id: u64,
) -> Result<TableAppearance> {
    let (_, model) = decode_unique_any::<tst::TableModelArchive>(
        package,
        model_object_id,
        TABLE_MODEL_MESSAGE_TYPES,
        "table model",
    )?;
    let Some(style_id) = effective_table_style_id(package, &model)? else {
        return Ok(TableAppearance::default());
    };
    inherited_table_appearance(package, style_id)
}

pub(crate) fn set_table_appearance(
    package: &mut IWorkPackage,
    model_object_id: u64,
    appearance: TableAppearance,
) -> Result<()> {
    if table_appearance(package, model_object_id)? == appearance {
        return Ok(());
    }
    let (model_archive, model) = decode_unique_any::<tst::TableModelArchive>(
        package,
        model_object_id,
        TABLE_MODEL_MESSAGE_TYPES,
        "table model",
    )?;
    let parent_style_id = effective_table_style_id(package, &model)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table model {model_object_id} has neither a table style nor a preset"
        ))
    })?;
    let (style_archive, parent_style) = decode_unique::<tst::TableStyleArchive>(
        package,
        parent_style_id,
        TABLE_STYLE_MESSAGE_TYPE,
        "table style",
    )?;
    let stylesheet_id = parent_style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table style {parent_style_id} has no stylesheet"
            ))
        })?;
    if object_archive_name(package, stylesheet_id)? != style_archive {
        return Err(Error::InvalidFormat(format!(
            "iWork table style {parent_style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }

    let new_style_id = next_object_identifier(package)?;
    let new_style =
        table_style_variation(new_style_id, parent_style_id, stylesheet_id, appearance)?;
    let mut staged = package.clone();
    patch_table_style_reference(
        &mut staged,
        &model_archive,
        model_object_id,
        model.table_style.identifier,
        new_style_id,
    )?;
    insert_style_variation(
        &mut staged,
        &style_archive,
        stylesheet_id,
        parent_style_id,
        new_style_id,
        new_style,
    )?;
    if let Some(style_component) = component_identifier_for_entry(&staged, &style_archive)? {
        add_component_object_uuids(&mut staged, style_component, &[new_style_id])?;
        if let Some(model_component) = component_identifier_for_entry(&staged, &model_archive)?
            && model_component != style_component
        {
            add_component_external_reference(
                &mut staged,
                model_component,
                style_component,
                new_style_id,
            )?;
        }
    }
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    if table_appearance(&staged, model_object_id)? != appearance {
        return Err(Error::InvalidFormat(
            "iWork table appearance failed round-trip validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

fn effective_table_style_id(
    package: &IWorkPackage,
    model: &tst::TableModelArchive,
) -> Result<Option<u64>> {
    if model.table_style.identifier != 0 {
        return Ok(Some(model.table_style.identifier));
    }
    let Some(preset_id) = model
        .table_style_preset
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
    else {
        return Ok(None);
    };
    let (_, preset) = decode_unique::<tst::TableStylePresetArchive>(
        package,
        preset_id,
        TABLE_STYLE_PRESET_MESSAGE_TYPE,
        "table style preset",
    )?;
    let network_id = preset
        .style_network
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table style preset {preset_id} has no style network"
            ))
        })?;
    let (_, network) = decode_unique::<tst::TableStyleNetworkArchive>(
        package,
        network_id,
        TABLE_STYLE_NETWORK_MESSAGE_TYPE,
        "table style network",
    )?;
    if network.table_style.identifier == 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork table style network {network_id} has no table style"
        )));
    }
    Ok(Some(network.table_style.identifier))
}

fn inherited_table_appearance(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TableAppearance> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    let mut banded_rows = None;
    let mut auto_resize = None;
    let mut horizontal_gridlines = None;
    let mut header_column_gridlines = None;
    let mut vertical_gridlines = None;
    let mut header_row_gridlines = None;
    let mut footer_row_gridlines = None;
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(TableAppearance {
                row_banding: TableRowBanding::from_native(banded_rows.unwrap_or(false)),
                row_sizing: TableRowSizing::from_native(auto_resize.unwrap_or(false)),
                gridlines: TableGridlines {
                    body_horizontal: TableGridlineVisibility::from_native(
                        horizontal_gridlines.unwrap_or(true),
                    ),
                    header_columns_horizontal: TableGridlineVisibility::from_native(
                        header_column_gridlines.unwrap_or(true),
                    ),
                    body_vertical: TableGridlineVisibility::from_native(
                        vertical_gridlines.unwrap_or(true),
                    ),
                    header_rows_vertical: TableGridlineVisibility::from_native(
                        header_row_gridlines.unwrap_or(true),
                    ),
                    footer_rows_vertical: TableGridlineVisibility::from_native(
                        footer_row_gridlines.unwrap_or(true),
                    ),
                },
            });
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork table style inheritance cycles at {identifier}"
            )));
        }
        let (style, overrides) = table_style_with_overrides(package, identifier)?;
        banded_rows = banded_rows.or(overrides.banded_rows);
        auto_resize = auto_resize.or(overrides.auto_resize);
        horizontal_gridlines = horizontal_gridlines.or(overrides.horizontal_body_gridlines);
        header_column_gridlines =
            header_column_gridlines.or(overrides.horizontal_header_column_gridlines);
        vertical_gridlines = vertical_gridlines.or(overrides.vertical_body_gridlines);
        header_row_gridlines = header_row_gridlines.or(overrides.vertical_header_row_gridlines);
        footer_row_gridlines = footer_row_gridlines.or(overrides.vertical_footer_row_gridlines);
        if let (
            Some(banded_rows),
            Some(auto_resize),
            Some(horizontal_gridlines),
            Some(header_column_gridlines),
            Some(vertical_gridlines),
            Some(header_row_gridlines),
            Some(footer_row_gridlines),
        ) = (
            banded_rows,
            auto_resize,
            horizontal_gridlines,
            header_column_gridlines,
            vertical_gridlines,
            header_row_gridlines,
            footer_row_gridlines,
        ) {
            return Ok(TableAppearance {
                row_banding: TableRowBanding::from_native(banded_rows),
                row_sizing: TableRowSizing::from_native(auto_resize),
                gridlines: TableGridlines {
                    body_horizontal: TableGridlineVisibility::from_native(horizontal_gridlines),
                    header_columns_horizontal: TableGridlineVisibility::from_native(
                        header_column_gridlines,
                    ),
                    body_vertical: TableGridlineVisibility::from_native(vertical_gridlines),
                    header_rows_vertical: TableGridlineVisibility::from_native(
                        header_row_gridlines,
                    ),
                    footer_rows_vertical: TableGridlineVisibility::from_native(
                        footer_row_gridlines,
                    ),
                },
            });
        }
        style_id = style
            .super_
            .parent
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0);
    }
    Err(Error::InvalidFormat(format!(
        "iWork table style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
    )))
}

fn table_style_with_overrides(
    package: &IWorkPackage,
    identifier: u64,
) -> Result<(tst::TableStyleArchive, TableAppearanceOverrides)> {
    let archive_name = object_archive_name(package, identifier)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table style {identifier} is missing"))
    })?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == TABLE_STYLE_MESSAGE_TYPE);
    let Some(message) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "iWork table style {identifier} must have exactly one native payload"
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "iWork table style {identifier} must have exactly one native payload"
        )));
    }
    Ok((
        tst::TableStyleArchive::decode(message.data.as_slice())?,
        table_appearance_overrides(&message.data)?,
    ))
}

fn table_style_variation(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    appearance: TableAppearance,
) -> Result<ArchiveObject> {
    let data = tst::TableStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(7),
        table_properties: Some(tst::TableStylePropertiesArchive {
            banded_rows: Some(appearance.row_banding.native()),
            auto_resize: Some(appearance.row_sizing.native()),
            h_strokes_visible: Some(appearance.gridlines.body_horizontal.native()),
            v_strokes_visible: Some(appearance.gridlines.body_vertical.native()),
            table_hc_divider_visible: Some(appearance.gridlines.header_columns_horizontal.native()),
            table_hr_divider_visible: Some(appearance.gridlines.header_rows_vertical.native()),
            table_footer_divider_visible: Some(appearance.gridlines.footer_rows_vertical.native()),
            ..Default::default()
        }),
    }
    .encode_to_vec();
    tst::TableStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: TABLE_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references.push(parent_style_id);
    if stylesheet_id != parent_style_id {
        info.object_references.push(stylesheet_id);
    }
    Ok(object)
}

fn patch_table_style_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    model_object_id: u64,
    old_style_id: u64,
    new_style_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(model_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table model {model_object_id} is missing"))
        })?;
        let mut indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_))
            .map(|(index, _)| index);
        let Some(index) = indexes.next() else {
            return Err(Error::InvalidFormat(format!(
                "iWork table model {model_object_id} must have exactly one native payload"
            )));
        };
        if indexes.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork table model {model_object_id} must have exactly one native payload"
            )));
        }
        let message_type = object.messages[index].type_;
        let data = patch_length_delimited_field(
            &object.messages[index].data,
            TABLE_STYLE_REFERENCE_FIELD,
            true,
            Some(&reference(new_style_id).encode_to_vec()),
        )?;
        let decoded = tst::TableModelArchive::decode(data.as_slice())?;
        if decoded.table_style.identifier != new_style_id {
            return Err(Error::InvalidFormat(format!(
                "iWork table model {model_object_id} rejected style {new_style_id}"
            )));
        }
        object.replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[index];
        if old_style_id == 0 {
            info.object_references.push(new_style_id);
        } else {
            let reference = info
                .object_references
                .iter_mut()
                .find(|identifier| **identifier == old_style_id)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "iWork table model {model_object_id} metadata omits style {old_style_id}"
                    ))
                })?;
            *reference = new_style_id;
        }
        for field in &mut info.field_infos {
            if field.path.path.as_slice() == [TABLE_STYLE_REFERENCE_FIELD] {
                for identifier in &mut field.object_references {
                    if *identifier == old_style_id {
                        *identifier = new_style_id;
                    }
                }
            }
        }
        Ok(())
    })
}

fn decode_unique<T: Message + Default>(
    package: &IWorkPackage,
    identifier: u64,
    message_type: u32,
    context: &str,
) -> Result<(String, T)> {
    decode_unique_any(package, identifier, &[message_type], context)
}

fn decode_unique_any<T: Message + Default>(
    package: &IWorkPackage,
    identifier: u64,
    message_types: &[u32],
    context: &str,
) -> Result<(String, T)> {
    let archive_name = object_archive_name(package, identifier)?;
    let archive = package.archive(&archive_name)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork {context} {identifier} is missing")))?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message_types.contains(&message.type_));
    let Some(message) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "iWork {context} {identifier} must have exactly one native payload"
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "iWork {context} {identifier} must have exactly one native payload"
        )));
    }
    Ok((archive_name, T::decode(message.data.as_slice())?))
}

fn object_archive_name(package: &IWorkPackage, identifier: u64) -> Result<String> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        if package.archive(name)?.object(identifier).is_some()
            && found.replace(name.to_owned()).is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork object {identifier} occurs in multiple archives"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork object {identifier} is missing")))
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
