#![allow(
    clippy::map_err_ignore,
    reason = "legacy module confines normalization into the module's stable typed public error to this codec boundary"
)]

//! Typed XLSB chart-sheet authoring (MS-XLSB 2.1.7.7).

use crate::chart::Chart;
use crate::package::chartsheet::{
    ChartSheet, Color, ColorType, PageSetup, Protection, State, View,
};
use crate::package::error::{Error, Result};
use crate::raw::Writer;
use crate::raw::kind as rt;
use crate::sheet::StrongProtection;
use std::io::Write;

const MAX_CHART_SHEETS: usize = 65_536;
const MAX_VIEWS: usize = 256;
const MAX_STRING_UNITS: usize = 32_767;
const MAX_PROTECTION_BYTES: usize = 1 << 20;
const MAX_PRINTER_SETTINGS_BYTES: usize = 16 << 20;
const MAX_THEME_COLOR_INDEX: u8 = 0x0b;
const MAX_INDEXED_COLOR: u8 = 0x51;
const MAX_SPIN_COUNT: u32 = 10_000_000;
const MAX_COPIES: u32 = 32_767;
const MAX_PAPER_SIZE: u32 = i32::MAX as u32 - 1;

/// A chart sheet being authored in a new XLSB workbook.
#[derive(Debug, Clone)]
pub struct MutableChartSheet {
    metadata: ChartSheet,
    chart: Chart,
    printer_settings: Option<Vec<u8>>,
}

impl MutableChartSheet {
    /// Create a visible chart sheet with one default workbook view.
    pub fn new(name: impl Into<String>, chart: Chart) -> Self {
        let name = name.into();
        Self {
            metadata: ChartSheet {
                name,
                state: State::Visible,
                code_name: String::new(),
                published: false,
                tab_color: Color::automatic(),
                views: vec![View {
                    selected: false,
                    scale: 100,
                    workbook_view_index: 0,
                }],
                protection: None,
                strong_protection: None,
                page_setup: None,
                drawing_rel_id: None,
                legacy_drawing_rel_id: None,
                legacy_drawing_header_footer_rel_id: None,
            },
            chart,
            printer_settings: None,
        }
    }

    /// Workbook-visible sheet name.
    pub fn name(&self) -> &str {
        &self.metadata.name
    }

    /// Sheet-level metadata written to the binary chart-sheet stream.
    pub fn metadata(&self) -> &ChartSheet {
        &self.metadata
    }

    /// Mutably configure sheet metadata.
    ///
    /// Relationship identifiers are allocated by the writer. Setting a
    /// drawing or legacy-drawing identifier causes validation to fail.
    pub fn metadata_mut(&mut self) -> &mut ChartSheet {
        &mut self.metadata
    }

    /// The hosted DrawingML chart.
    pub fn chart(&self) -> &Chart {
        &self.chart
    }

    /// Mutably configure the hosted DrawingML chart.
    pub fn chart_mut(&mut self) -> &mut Chart {
        &mut self.chart
    }

    /// Attach an opaque, bounded printer-settings payload and its page setup.
    ///
    /// The relationship identifier in `page_setup` is ignored and replaced
    /// with an identifier allocated by the package writer.
    pub fn set_page_setup(
        &mut self,
        mut page_setup: PageSetup,
        printer_settings: Vec<u8>,
    ) -> Result<&mut Self> {
        if printer_settings.is_empty() || printer_settings.len() > MAX_PRINTER_SETTINGS_BYTES {
            return Err(Error::InvalidLength {
                expected: MAX_PRINTER_SETTINGS_BYTES,
                found: printer_settings.len(),
            });
        }
        page_setup.printer_settings_rel_id.clear();
        self.metadata.page_setup = Some(page_setup);
        self.printer_settings = Some(printer_settings);
        Ok(self)
    }

    /// Remove page setup and its printer-settings payload.
    pub fn clear_page_setup(&mut self) -> bool {
        let changed = self.metadata.page_setup.take().is_some();
        self.printer_settings = None;
        changed
    }

    pub(crate) fn printer_settings(&self) -> Option<&[u8]> {
        self.printer_settings.as_deref()
    }

    pub(crate) fn validate(&self) -> Result<()> {
        validate_name(self.name())?;
        if self.metadata.drawing_rel_id.is_some()
            || self.metadata.legacy_drawing_rel_id.is_some()
            || self.metadata.legacy_drawing_header_footer_rel_id.is_some()
        {
            return Err(Error::UnsupportedFeature(
                "chart-sheet relationship IDs are allocated by the XLSB writer; legacy VML authoring is unsupported"
                    .to_string(),
            ));
        }
        validate_string(&self.metadata.code_name, "chart-sheet code name")?;
        validate_color(self.metadata.tab_color)?;
        if self.metadata.views.is_empty() || self.metadata.views.len() > MAX_VIEWS {
            return Err(Error::InvalidLength {
                expected: MAX_VIEWS,
                found: self.metadata.views.len(),
            });
        }
        for view in &self.metadata.views {
            if view.scale != 0 && !(10..=400).contains(&view.scale) {
                return Err(Error::InvalidFormula(format!(
                    "chart-sheet zoom {} is outside 10..=400 or zero",
                    view.scale
                )));
            }
            if view.workbook_view_index != 0 {
                return Err(Error::UnsupportedFeature(format!(
                    "chart-sheet workbook view index {} has no authored workbook view",
                    view.workbook_view_index
                )));
            }
        }
        if let Some(strong) = &self.metadata.strong_protection {
            validate_strong_protection(strong)?;
            if self.metadata.protection.is_none() {
                return Err(Error::InvalidFormula(
                    "strong chart-sheet protection requires classic protection flags".to_string(),
                ));
            }
        }
        match (&self.metadata.page_setup, &self.printer_settings) {
            (Some(setup), Some(_)) => validate_page_setup(setup)?,
            (None, None) => {},
            _ => {
                return Err(Error::InvalidFormula(
                    "chart-sheet page setup and printer settings must be supplied together"
                        .to_string(),
                ));
            },
        }
        crate::package::drawing_write::validate_chart(&self.chart)
    }
}

pub(crate) fn max_chart_sheets() -> usize {
    MAX_CHART_SHEETS
}

pub(crate) fn write_chart_sheet(
    sheet: &MutableChartSheet,
    drawing_rel_id: &str,
    printer_rel_id: Option<&str>,
) -> Result<Vec<u8>> {
    sheet.validate()?;
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    writer.write_record(rt::BEGIN_SHEET, &[])?;
    write_properties(&mut writer, &sheet.metadata)?;
    write_views(&mut writer, &sheet.metadata.views)?;
    write_protection(
        &mut writer,
        sheet.metadata.protection,
        sheet.metadata.strong_protection.as_ref(),
    )?;
    if let Some(setup) = &sheet.metadata.page_setup {
        write_page_setup(
            &mut writer,
            setup,
            printer_rel_id.ok_or_else(|| {
                Error::InvalidFormula(
                    "chart-sheet page setup has no printer relationship".to_string(),
                )
            })?,
        )?;
    }
    write_rel_id(&mut writer, rt::DRAWING, drawing_rel_id)?;
    writer.write_record(rt::END_SHEET, &[])?;
    Ok(output)
}

fn validate_name(name: &str) -> Result<()> {
    let units = name.encode_utf16().count();
    if units == 0
        || units > 31
        || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
        || name.starts_with('\'')
        || name.ends_with('\'')
    {
        return Err(Error::InvalidFormula(format!(
            "chart-sheet name {name:?} does not follow BrtBundleSh grammar"
        )));
    }
    Ok(())
}

fn validate_string(value: &str, context: &str) -> Result<()> {
    if value.encode_utf16().count() > MAX_STRING_UNITS || value.contains('\0') {
        return Err(Error::InvalidFormula(format!(
            "{context} is too long or contains NUL"
        )));
    }
    Ok(())
}

fn validate_color(color: Color) -> Result<()> {
    if color.color_type == ColorType::Rgb && !color.valid_rgb {
        return Err(Error::InvalidFormula(
            "direct chart-sheet tab color is not marked valid".to_string(),
        ));
    }
    if color.color_type == ColorType::Theme && color.index > MAX_THEME_COLOR_INDEX {
        return Err(Error::InvalidFormula(format!(
            "chart-sheet theme color index {} exceeds {MAX_THEME_COLOR_INDEX}",
            color.index
        )));
    }
    if color.color_type == ColorType::Indexed && color.index > MAX_INDEXED_COLOR {
        return Err(Error::InvalidFormula(format!(
            "chart-sheet indexed color {} exceeds {MAX_INDEXED_COLOR}",
            color.index
        )));
    }
    Ok(())
}

fn validate_strong_protection(value: &StrongProtection) -> Result<()> {
    if value.spin_count > MAX_SPIN_COUNT
        || value.hash.is_empty()
        || value.hash.len() > MAX_PROTECTION_BYTES
        || value.salt.len() > MAX_PROTECTION_BYTES
        || value.algorithm.is_empty()
    {
        return Err(Error::InvalidFormula(
            "invalid or oversized chart-sheet strong-protection metadata".to_string(),
        ));
    }
    validate_string(&value.algorithm, "strong-protection algorithm")
}

fn validate_page_setup(value: &PageSetup) -> Result<()> {
    if value.paper_size > MAX_PAPER_SIZE || (119..256).contains(&value.paper_size) {
        return Err(Error::InvalidFormula(format!(
            "chart-sheet paper size {} is invalid or reserved",
            value.paper_size
        )));
    }
    if value.copies == 0 || value.copies > MAX_COPIES {
        return Err(Error::InvalidFormula(format!(
            "chart-sheet print copies {} is outside 1..={MAX_COPIES}",
            value.copies
        )));
    }
    Ok(())
}

fn write_properties<W: Write>(writer: &mut Writer<W>, value: &ChartSheet) -> Result<()> {
    let mut data = Vec::new();
    data.extend_from_slice(&(u16::from(value.published)).to_le_bytes());
    let color_type = match value.tab_color.color_type {
        ColorType::Automatic => 0,
        ColorType::Indexed => 1,
        ColorType::Rgb => 2,
        ColorType::Theme => 3,
    };
    data.push(u8::from(value.tab_color.valid_rgb) | (color_type << 1));
    data.push(value.tab_color.index);
    data.extend_from_slice(&value.tab_color.tint.to_le_bytes());
    data.extend_from_slice(&value.tab_color.rgba);
    Writer::new(&mut data).write_wide_string(&value.code_name)?;
    Ok(writer.write_record(rt::CS_PROP, &data)?)
}

fn write_views<W: Write>(writer: &mut Writer<W>, views: &[View]) -> Result<()> {
    writer.write_record(rt::BEGIN_CS_VIEWS, &[])?;
    for view in views {
        let mut data = Vec::with_capacity(10);
        data.extend_from_slice(&u16::from(view.selected).to_le_bytes());
        data.extend_from_slice(&view.scale.to_le_bytes());
        data.extend_from_slice(&view.workbook_view_index.to_le_bytes());
        writer.write_record(rt::BEGIN_CS_VIEW, &data)?;
        writer.write_record(rt::END_CS_VIEW, &[])?;
    }
    Ok(writer.write_record(rt::END_CS_VIEWS, &[])?)
}

fn write_protection<W: Write>(
    writer: &mut Writer<W>,
    classic: Option<Protection>,
    strong: Option<&StrongProtection>,
) -> Result<()> {
    let Some(classic) = classic else {
        return Ok(());
    };
    if let Some(strong) = strong {
        let mut data = Vec::new();
        data.extend_from_slice(&strong.spin_count.to_le_bytes());
        data.extend_from_slice(&u32::from(classic.locked).to_le_bytes());
        data.extend_from_slice(&u32::from(classic.objects).to_le_bytes());
        write_blob(&mut data, &strong.hash)?;
        write_blob(&mut data, &strong.salt)?;
        Writer::new(&mut data).write_wide_string(&strong.algorithm)?;
        writer.write_record(rt::CS_PROTECTION_ISO, &data)?;
    }
    let mut data = Vec::with_capacity(10);
    data.extend_from_slice(
        &if strong.is_some() {
            0
        } else {
            classic.password_verifier
        }
        .to_le_bytes(),
    );
    data.extend_from_slice(&u32::from(classic.locked).to_le_bytes());
    data.extend_from_slice(&u32::from(classic.objects).to_le_bytes());
    Ok(writer.write_record(rt::CS_PROTECTION, &data)?)
}

fn write_page_setup<W: Write>(
    writer: &mut Writer<W>,
    setup: &PageSetup,
    printer_rel_id: &str,
) -> Result<()> {
    let mut data = Vec::new();
    data.extend_from_slice(&setup.paper_size.to_le_bytes());
    data.extend_from_slice(&setup.horizontal_resolution.to_le_bytes());
    data.extend_from_slice(&setup.vertical_resolution.to_le_bytes());
    data.extend_from_slice(&setup.copies.to_le_bytes());
    data.extend_from_slice(&setup.page_start.to_le_bytes());
    let flags = u16::from(setup.landscape)
        | (u16::from(setup.black_and_white) << 2)
        | (u16::from(setup.use_default_orientation) << 3)
        | (u16::from(setup.use_page_start) << 4)
        | (u16::from(setup.draft) << 5);
    data.extend_from_slice(&flags.to_le_bytes());
    Writer::new(&mut data).write_wide_string(printer_rel_id)?;
    Ok(writer.write_record(rt::CS_PAGE_SETUP, &data)?)
}

fn write_rel_id<W: Write>(
    writer: &mut Writer<W>,
    record_type: crate::raw::Kind,
    rel_id: &str,
) -> Result<()> {
    let mut data = Vec::new();
    Writer::new(&mut data).write_wide_string(rel_id)?;
    Ok(writer.write_record(record_type, &data)?)
}

fn write_blob(data: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    data.extend_from_slice(
        &u32::try_from(value.len())
            .map_err(|_| Error::InvalidLength {
                expected: u32::MAX as usize,
                found: value.len(),
            })?
            .to_le_bytes(),
    );
    data.extend_from_slice(value);
    Ok(())
}
