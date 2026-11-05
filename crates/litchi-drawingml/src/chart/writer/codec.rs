//! XML/wire codec façade for `DrawingML` charts.
//!
//! Record-family serialization lives under [`super::semantic`]. This module
//! keeps the public writer contract small and stable while package owners can
//! still provide relationship identifiers through [`write_with_rels`].

use super::semantic::write_chart_space;
use super::validation::{invalid_chart_input, validate_chart_style};
use crate::chart::Limits;
use crate::chart::model::Chart;
use crate::{Error, Result};
use std::io::Write;

/// Write a chart to canonical `DrawingML` chart XML.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write<W: Write>(writer: &mut W, chart: &Chart, limits: &Limits) -> Result<()> {
    write_with_rels(writer, chart, limits, None, None)
}

/// Write chart XML using relationship identifiers assigned by its package owner.
///
/// The chart model owns relationship metadata, while DOCX, PPTX, XLSX, and
/// XLSB package writers remain responsible for allocating and validating OPC
/// relationships. This entry point only serializes the resulting wire form.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn write_with_rels<W: Write>(
    writer: &mut W,
    chart: &Chart,
    limits: &Limits,
    external_data_relationship_id: Option<&str>,
    user_shapes_relationship_id: Option<&str>,
) -> Result<()> {
    let limits = limits.validate()?;
    limits.validate_chart(chart)?;
    validate_structure(
        chart,
        limits,
        external_data_relationship_id,
        user_shapes_relationship_id,
    )?;
    let mut bounded = LimitedWriter::new(writer, limits.max_output_bytes);
    match write_chart_space(
        &mut bounded,
        chart,
        external_data_relationship_id,
        user_shapes_relationship_id,
    ) {
        Ok(()) => Ok(()),
        Err(_error) if bounded.exceeded => Err(Error::limit(
            crate::LimitKind::OutputBytes,
            bounded.attempted,
            limits.max_output_bytes,
        )),
        Err(error) => Err(Error::Io(error)),
    }
}

fn validate_structure(
    chart: &Chart,
    limits: Limits,
    external_data_relationship_id: Option<&str>,
    user_shapes_relationship_id: Option<&str>,
) -> Result<()> {
    validate_chart_style(chart.style).map_err(Error::Io)?;
    validate_relationship(
        limits,
        chart
            .external_data
            .as_ref()
            .map(|data| data.relationship_id.as_deref()),
        external_data_relationship_id,
        "chart external data",
    )?;
    validate_relationship(
        limits,
        chart
            .user_shapes
            .as_ref()
            .map(|data| data.relationship_id.as_deref()),
        user_shapes_relationship_id,
        "chart user shapes",
    )?;
    Ok(())
}

fn validate_relationship(
    limits: Limits,
    metadata: Option<Option<&str>>,
    override_id: Option<&str>,
    description: &str,
) -> Result<()> {
    if metadata.is_none() {
        if override_id.is_some() {
            return Err(Error::Io(invalid_chart_input(format!(
                "chart package supplied {description} without chart metadata"
            ))));
        }
        return Ok(());
    }
    let relationship_id = override_id.or(metadata.flatten());
    if relationship_id.is_none_or(str::is_empty) {
        return Err(Error::Io(invalid_chart_input(format!(
            "{description} have no relationship ID"
        ))));
    }
    if let Some(relationship_id) = relationship_id {
        if relationship_id.len() > limits.max_attribute_value_bytes {
            return Err(Error::limit(
                crate::LimitKind::AttributeValueBytes,
                relationship_id.len(),
                limits.max_attribute_value_bytes,
            ));
        }
    }
    Ok(())
}

struct LimitedWriter<'a, W: Write> {
    inner: &'a mut W,
    maximum: usize,
    written: usize,
    attempted: usize,
    exceeded: bool,
}

impl<'a, W: Write> LimitedWriter<'a, W> {
    fn new(inner: &'a mut W, maximum: usize) -> Self {
        Self {
            inner,
            maximum,
            written: 0,
            attempted: 0,
            exceeded: false,
        }
    }
}

impl<W: Write> Write for LimitedWriter<'_, W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.attempted = self.written.saturating_add(bytes.len());
        if bytes.len() > self.maximum.saturating_sub(self.written) {
            self.exceeded = true;
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "chart output exceeds the configured byte limit",
            ));
        }
        let written = self.inner.write(bytes)?;
        self.written = self.written.saturating_add(written);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}
