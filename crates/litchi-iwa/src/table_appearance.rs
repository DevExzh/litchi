//! Compatibility reads for native iWork table appearance metadata.

mod wire;

use std::collections::HashSet;

use prost::Message;

use crate::protobuf::tst;
use crate::{Error, IWorkPackage, Result};
use wire::{TableAppearanceOverrides, table_appearance_overrides};

pub use litchi_iwa_common::table::appearance::{
    Appearance, Banding, GridlineVisibility, Gridlines, RowSizing,
};

/// Facade-local name for the shared table appearance value.
pub type TableAppearance = Appearance;
/// Facade-local name for the shared alternating-row setting.
pub type TableRowBanding = Banding;
/// Facade-local name for the shared row-sizing setting.
pub type TableRowSizing = RowSizing;
/// Facade-local name for the shared gridline visibility setting.
pub type TableGridlineVisibility = GridlineVisibility;
/// Facade-local name for the shared per-region gridline settings.
pub type TableGridlines = Gridlines;

const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

fn banding_from_native(value: bool) -> TableRowBanding {
    if value {
        Banding::Enabled
    } else {
        Banding::Disabled
    }
}

fn row_sizing_from_native(value: bool) -> TableRowSizing {
    if value {
        RowSizing::FitCellContents
    } else {
        RowSizing::Fixed
    }
}

fn gridline_visibility_from_native(value: bool) -> TableGridlineVisibility {
    if value {
        GridlineVisibility::Visible
    } else {
        GridlineVisibility::Hidden
    }
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
                row_banding: banding_from_native(banded_rows.unwrap_or(false)),
                row_sizing: row_sizing_from_native(auto_resize.unwrap_or(false)),
                gridlines: TableGridlines {
                    body_horizontal: gridline_visibility_from_native(
                        horizontal_gridlines.unwrap_or(true),
                    ),
                    header_columns_horizontal: gridline_visibility_from_native(
                        header_column_gridlines.unwrap_or(true),
                    ),
                    body_vertical: gridline_visibility_from_native(
                        vertical_gridlines.unwrap_or(true),
                    ),
                    header_rows_vertical: gridline_visibility_from_native(
                        header_row_gridlines.unwrap_or(true),
                    ),
                    footer_rows_vertical: gridline_visibility_from_native(
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
                row_banding: banding_from_native(banded_rows),
                row_sizing: row_sizing_from_native(auto_resize),
                gridlines: TableGridlines {
                    body_horizontal: gridline_visibility_from_native(horizontal_gridlines),
                    header_columns_horizontal: gridline_visibility_from_native(
                        header_column_gridlines,
                    ),
                    body_vertical: gridline_visibility_from_native(vertical_gridlines),
                    header_rows_vertical: gridline_visibility_from_native(header_row_gridlines),
                    footer_rows_vertical: gridline_visibility_from_native(footer_row_gridlines),
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
