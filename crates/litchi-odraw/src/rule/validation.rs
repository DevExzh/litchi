//! Structural validation for `OfficeArt` solver-rule records.

use crate::{Error, Record, RecordKind, Result};

/// A validated fixed-layout rule kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Layout {
    Connector,
    Arc,
    Callout,
    Opaque,
}

/// Checks the record header and body extent required by `[MS-ODRAW]`.
pub(super) fn validate(record: &Record<'_>) -> Result<Layout> {
    match record.kind() {
        RecordKind::ConnectorRule => {
            validate_fixed(record, 1, 0, 24, "connector rule")?;
            Ok(Layout::Connector)
        },
        RecordKind::ArcRule => {
            validate_fixed(record, 0, 0, 8, "arc rule")?;
            Ok(Layout::Arc)
        },
        RecordKind::CalloutRule => {
            validate_fixed(record, 0, 0, 8, "callout rule")?;
            Ok(Layout::Callout)
        },
        RecordKind::Unknown(_) => Ok(Layout::Opaque),
        _ => Err(Error::MalformedShape {
            reason: "record is not an OfficeArt solver rule",
        }),
    }
}

fn validate_fixed(
    record: &Record<'_>,
    version: u8,
    instance: u16,
    length: u32,
    name: &'static str,
) -> Result<()> {
    if record.version() != version {
        return Err(Error::MalformedShape {
            reason: match name {
                "connector rule" => "connector rule version is invalid",
                "arc rule" => "arc rule version is invalid",
                "callout rule" => "callout rule version is invalid",
                _ => "OfficeArt solver rule version is invalid",
            },
        });
    }
    if record.instance() != instance {
        return Err(Error::MalformedShape {
            reason: match name {
                "connector rule" => "connector rule instance is invalid",
                "arc rule" => "arc rule instance is invalid",
                "callout rule" => "callout rule instance is invalid",
                _ => "OfficeArt solver rule instance is invalid",
            },
        });
    }
    if record.len() != length {
        return Err(Error::MalformedShape {
            reason: match name {
                "connector rule" => "connector rule length is invalid",
                "arc rule" => "arc rule length is invalid",
                "callout rule" => "callout rule length is invalid",
                _ => "OfficeArt solver rule length is invalid",
            },
        });
    }
    Ok(())
}
