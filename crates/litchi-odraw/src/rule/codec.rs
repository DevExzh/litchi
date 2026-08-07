//! `OfficeArt` solver-rule decoding and writing.

use std::io::{self, Write};

use super::model::{Arc, Callout, Connector, Opaque, Rule};
use super::validation::{Layout, validate};
use crate::{Error, Record, Result, write};

impl<'data> TryFrom<Record<'data>> for Rule<'data> {
    type Error = Error;

    fn try_from(record: Record<'data>) -> Result<Self> {
        from_record(record)
    }
}

impl Rule<'_> {
    /// Writes the complete `OfficeArt` record, including its eight-byte header.
    ///
    /// Known rules are emitted with their required fixed header.  Opaque
    /// records reuse every header field and body byte from the borrowed input.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying writer fails.
    pub fn write_to<W: Write>(&self, writer: &mut W) -> io::Result<()> {
        match self {
            Self::Connector(rule) => {
                write::raw_header(writer, 1, 0, 0xF012, 24)?;
                write_field(writer, rule.rule_id())?;
                write_field(writer, rule.start_shape_id())?;
                write_field(writer, rule.end_shape_id())?;
                write_field(writer, rule.connector_shape_id())?;
                write_field(writer, rule.start_connection_site())?;
                write_field(writer, rule.end_connection_site())
            },
            Self::Arc(rule) => {
                write::raw_header(writer, 0, 0, 0xF014, 8)?;
                write_field(writer, rule.rule_id())?;
                write_field(writer, rule.shape_id())
            },
            Self::Callout(rule) => {
                write::raw_header(writer, 0, 0, 0xF017, 8)?;
                write_field(writer, rule.rule_id())?;
                write_field(writer, rule.shape_id())
            },
            Self::Opaque(record) => {
                write::raw_header(
                    writer,
                    record.version(),
                    record.instance(),
                    record.raw_kind(),
                    record.len(),
                )?;
                writer.write_all(record.data())
            },
        }
    }
}

/// Parses exactly one `OfficeArt` solver-rule record.
///
/// # Errors
///
/// Returns `Error::TruncatedHeader` or `Error::TruncatedPayload` when the
/// record header or declared body is incomplete, `Error::ArithmeticOverflow`
/// when a declared extent cannot be represented, and `Error::TrailingData`
/// when bytes follow the first record.  Returns `Error::MalformedShape` under
/// the conditions documented on [`from_record`].
pub fn parse(data: &[u8]) -> Result<Rule<'_>> {
    let (record, consumed) = Record::parse(data, 0)?;
    if consumed != data.len() {
        return Err(Error::TrailingData { offset: consumed });
    }
    from_record(record)
}

/// Decodes a previously parsed `OfficeArt` record as a solver rule.
///
/// # Errors
///
/// Returns `Error::MalformedShape` when the record kind is not an `OfficeArt`
/// solver rule, when a known rule violates its fixed `[MS-ODRAW]` version,
/// instance, or body length, or when a fixed-width field is truncated.
/// Returns `Error::ArithmeticOverflow` when a field extent cannot be
/// represented.
pub fn from_record(record: Record<'_>) -> Result<Rule<'_>> {
    let layout = validate(&record)?;
    match layout {
        Layout::Connector => Ok(Rule::Connector(Connector::new(
            field(record.data(), 0, "connector rule identifier")?,
            field(record.data(), 4, "connector start shape identifier")?,
            field(record.data(), 8, "connector end shape identifier")?,
            field(record.data(), 12, "connector shape identifier")?,
            field(record.data(), 16, "connector start connection site")?,
            field(record.data(), 20, "connector end connection site")?,
        ))),
        Layout::Arc => Ok(Rule::Arc(Arc::new(
            field(record.data(), 0, "arc rule identifier")?,
            field(record.data(), 4, "arc shape identifier")?,
        ))),
        Layout::Callout => Ok(Rule::Callout(Callout::new(
            field(record.data(), 0, "callout rule identifier")?,
            field(record.data(), 4, "callout shape identifier")?,
        ))),
        Layout::Opaque => Ok(Rule::Opaque(Opaque { record })),
    }
}

fn field(data: &[u8], offset: usize, name: &'static str) -> Result<u32> {
    let end = offset.checked_add(4).ok_or(Error::ArithmeticOverflow {
        context: "solver-rule field extent",
    })?;
    let bytes = data.get(offset..end).ok_or(Error::MalformedShape {
        reason: match name {
            "connector rule identifier" | "arc rule identifier" | "callout rule identifier" => {
                "solver rule identifier is truncated"
            },
            "connector start shape identifier"
            | "connector end shape identifier"
            | "connector shape identifier"
            | "arc shape identifier"
            | "callout shape identifier" => "solver rule shape identifier is truncated",
            "connector start connection site" | "connector end connection site" => {
                "connector connection site is truncated"
            },
            _ => "solver rule field is truncated",
        },
    })?;
    Ok(u32::from_le_bytes(bytes.try_into().map_err(|_err| {
        Error::MalformedShape {
            reason: "solver rule field is not four bytes",
        }
    })?))
}

fn write_field<W: Write>(writer: &mut W, value: u32) -> io::Result<()> {
    writer.write_all(&value.to_le_bytes())
}
