#![allow(
    clippy::expect_used,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check to this codec boundary"
)]

//! Workbook writer state and public configuration API.

use crate::calc::Props;
use crate::named_ranges::{Definition, validate_name};
use crate::package::error::{Error, Result};
use crate::package::formula::excel_name_eq;
use crate::writer::{
    MutableChartSheet, MutableSharedStringsWriter, MutableWorksheet, StylesWriter,
};
#[cfg(feature = "vba-inspection")]
use std::sync::Arc;

const MAX_AUTHORED_EXTERNAL_LINKS: usize = 4_096;

/// XLSB workbook writer
///
/// Creates complete XLSB workbook files with support for:
/// - Multiple worksheets
/// - Shared strings
/// - Styles (fonts, fills, borders, number formats)
/// - Workbook properties (date system, etc.)
///
/// # Example
///
/// ```rust,no_run
/// use litchi_xlsb::writer::{WorkbookWriter, MutableWorksheet};
/// use std::fs::File;
///
/// let mut workbook = WorkbookWriter::new();
///
/// let mut sheet = MutableWorksheet::new("Sheet1");
/// sheet.set_cell(0, 0, "Hello");
/// sheet.set_cell(0, 1, 42.0);
///
/// workbook.add_worksheet(sheet);
///
/// let file = File::create("output.xlsb")?;
/// workbook.save(file)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct WorkbookWriter {
    pub(super) worksheets: Vec<MutableWorksheet>,
    pub(super) chart_sheets: Vec<MutableChartSheet>,
    pub(super) sheet_order: Vec<SheetSlot>,
    pub(super) named_ranges: Vec<Definition>,
    pub(super) shared_strings: MutableSharedStringsWriter,
    pub(super) styles: StylesWriter,
    pub(super) calc: Props,
    pub(super) is_1904: bool,
    pub(super) connections: Option<crate::package::connections::Connections>,
    pub(super) external_links: Vec<crate::external_link::Link>,
    pub(super) pivot_caches: Vec<AuthoredPivotCache>,
    pub(super) xml_maps: Option<crate::xml_maps::XmlMapInfo>,
    pub(super) data_model: Option<crate::data_model::Model>,
    /// Optional typed DrawingML theme; `None` uses the writer's Office
    /// template as an authoring policy.
    pub(super) theme: Option<crate::theme::Theme>,
    /// Optional DrawingML 2012 applied-theme family metadata.
    pub(super) theme_family: Option<crate::theme::Family>,
    #[cfg(feature = "vba-inspection")]
    pub(super) vba: Option<Arc<Vec<u8>>>,
}

#[derive(Debug, Clone, Copy)]
pub(super) enum SheetSlot {
    Worksheet(usize),
    ChartSheet(usize),
}

pub(super) struct AuthoredPivotCache {
    pub(super) id: u32,
    pub(super) version_created: u8,
    pub(super) bytes: Vec<u8>,
}

impl WorkbookWriter {
    /// Create a new XLSB workbook writer
    pub fn new() -> Self {
        WorkbookWriter {
            worksheets: Vec::new(),
            chart_sheets: Vec::new(),
            sheet_order: Vec::new(),
            named_ranges: Vec::new(),
            shared_strings: MutableSharedStringsWriter::new(),
            styles: StylesWriter::new(),
            calc: Props::default(),
            is_1904: false,
            connections: None,
            external_links: Vec::new(),
            pivot_caches: Vec::new(),
            xml_maps: None,
            data_model: None,
            theme: None,
            theme_family: None,
            #[cfg(feature = "vba-inspection")]
            vba: None,
        }
    }

    /// Set the date system (1900 or 1904)
    ///
    /// # Arguments
    ///
    /// * `is_1904` - `true` for 1904 date system (Mac), `false` for 1900 (Windows, default)
    pub fn set_date_system(&mut self, is_1904: bool) {
        self.is_1904 = is_1904;
    }

    /// Set the typed DrawingML Theme used by the generated workbook.
    ///
    /// The Theme is validated and encoded only after the complete typed model
    /// is supplied; failed validation leaves the writer unchanged.
    pub fn set_theme(&mut self, theme: crate::theme::Theme) -> Result<&mut Self> {
        litchi_drawingml::theme::codec::encode_part(&theme.name, &theme.colors, &theme.fonts)?;
        self.theme = Some(theme);
        Ok(self)
    }

    /// Set the optional DrawingML 2012 applied-theme family metadata.
    pub fn set_theme_family(&mut self, family: crate::theme::Family) -> Result<&mut Self> {
        self.theme_family = Some(family);
        Ok(self)
    }

    /// Validated workbook formula calculation policy written to `BrtCalcProp`.
    pub fn calc(&self) -> &Props {
        &self.calc
    }

    /// Mutably configure the policy through its safe setters.
    pub fn calc_mut(&mut self) -> &mut Props {
        &mut self.calc
    }

    /// Replace the policy by moving in an already validated value.
    pub fn put_calc(&mut self, props: Props) -> &mut Self {
        self.calc = props;
        self
    }

    /// Attach a cache-free, inert MS-OVBA project to generated workbooks.
    #[cfg(feature = "vba-inspection")]
    pub fn set_vba(&mut self, project: litchi_vba::build::Project) -> Result<&mut Self> {
        self.set_vba_with(project, &litchi_vba::Limits::default())
    }

    /// Attach a cache-free project with explicit resource limits.
    #[cfg(feature = "vba-inspection")]
    pub fn set_vba_with(
        &mut self,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> Result<&mut Self> {
        let payload = project.finish(limits)?;
        Ok(self.put_vba(payload))
    }

    /// Attach a prevalidated `vbaProject.bin` payload without executing it.
    #[cfg(feature = "vba-inspection")]
    pub fn put_vba(&mut self, payload: litchi_vba::Payload) -> &mut Self {
        self.vba = Some(Arc::new(payload.into_bytes()));
        self
    }

    /// Remove the project scheduled for insertion into generated workbooks.
    #[cfg(feature = "vba-inspection")]
    pub fn clear_vba(&mut self) -> bool {
        self.vba.take().is_some()
    }

    /// Add a worksheet to the workbook
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi_xlsb::writer::{WorkbookWriter, MutableWorksheet};
    ///
    /// let mut workbook = WorkbookWriter::new();
    /// let sheet = MutableWorksheet::new("Sheet1");
    /// workbook.add_worksheet(sheet);
    /// ```
    pub fn add_worksheet(&mut self, worksheet: MutableWorksheet) {
        self.worksheets.push(worksheet);
        self.sheet_order
            .push(SheetSlot::Worksheet(self.worksheets.len() - 1));
    }

    /// Add a chart sheet in the current workbook sheet order.
    ///
    /// The chart is stored inertly in standard DrawingML parts. Package
    /// relationships are allocated at save time and never followed or fetched.
    pub fn add_chart_sheet(
        &mut self,
        chart_sheet: MutableChartSheet,
    ) -> Result<&mut MutableChartSheet> {
        chart_sheet.validate()?;
        if self.chart_sheets.len() >= crate::writer::chartsheet::max_chart_sheets() {
            return Err(Error::InvalidFormula(
                "XLSB chart-sheet count limit exceeded".to_string(),
            ));
        }
        if self
            .sheet_order
            .iter()
            .any(|slot| excel_name_eq(self.sheet_name(*slot), chart_sheet.name()))
        {
            return Err(Error::InvalidFormula(format!(
                "duplicate sheet name {:?}",
                chart_sheet.name()
            )));
        }
        self.chart_sheets.push(chart_sheet);
        self.sheet_order
            .push(SheetSlot::ChartSheet(self.chart_sheets.len() - 1));
        Ok(self
            .chart_sheets
            .last_mut()
            .expect("chart sheet was just added"))
    }

    /// Add a named range (defined name) to the workbook.
    pub fn add_named_range(&mut self, named_range: Definition) {
        self.named_ranges.push(named_range);
    }

    /// Attach an External Data Connections part (MS-XLSB 2.1.7.24) to the
    /// workbook.
    ///
    /// Connection identifiers and names must be unique and non-empty;
    /// strings, commands, URLs, and credential metadata are stored verbatim
    /// and are never resolved, contacted, refreshed, or executed.
    pub fn set_connections(
        &mut self,
        connections: crate::package::connections::Connections,
    ) -> Result<()> {
        crate::package::connections::write::validate_connections(&connections)?;
        if let Some(model) = self.data_model.as_ref() {
            crate::data_model::validate_definition_connections(
                &model.definition,
                Some(&connections),
            )?;
        }
        self.connections = Some(connections);
        Ok(())
    }

    /// The attached External Data Connections part, when set.
    pub fn connections(&self) -> Option<&crate::package::connections::Connections> {
        self.connections.as_ref()
    }

    /// Add inert Workbook, DDE, or OLE external-link metadata.
    ///
    /// External targets are stored but never followed, contacted, refreshed,
    /// instantiated, evaluated, or executed.
    pub fn add_external_link(&mut self, link: crate::external_link::Link) -> Result<&mut Self> {
        link.validate()?;
        if self.external_links.len() >= MAX_AUTHORED_EXTERNAL_LINKS {
            return Err(Error::InvalidFormula(
                "XLSB external-link count exceeds the safety limit".to_string(),
            ));
        }
        self.external_links.push(link);
        Ok(self)
    }

    /// External links scheduled for authoring, in workbook support-link order.
    pub fn external_links(&self) -> &[crate::external_link::Link] {
        &self.external_links
    }

    /// Replace the workbook's inert Custom XML Maps catalog.
    ///
    /// Schemas, XPath strings, and data-binding metadata are validated and
    /// stored only. No referenced resource is resolved, opened, or evaluated.
    pub fn set_xml_maps(&mut self, value: crate::xml_maps::XmlMapInfo) -> Result<&mut Self> {
        crate::xml_maps::validate_catalog(&value, crate::xml_maps::XmlMapLimits::DEFAULT)?;
        crate::xml_maps::serialize_xml_map_info(
            &value,
            crate::xml_maps::XmlMapConformance::Transitional,
        )?;
        self.xml_maps = Some(value);
        Ok(self)
    }

    /// Borrow the Custom XML Maps catalog scheduled for authoring.
    pub fn xml_maps(&self) -> Option<&crate::xml_maps::XmlMapInfo> {
        self.xml_maps.as_ref()
    }

    /// Compatibility alias for [`Self::xml_maps`].
    pub fn xml_map_info(&self) -> Option<&crate::xml_maps::XmlMapInfo> {
        self.xml_maps()
    }

    /// Remove and return the scheduled Custom XML Maps catalog.
    pub fn clear_xml_maps(&mut self) -> Option<crate::xml_maps::XmlMapInfo> {
        self.xml_maps.take()
    }

    /// Compatibility alias for [`Self::clear_xml_maps`].
    pub fn clear_xml_map_info(&mut self) -> Option<crate::xml_maps::XmlMapInfo> {
        self.clear_xml_maps()
    }

    /// Attach a complete inert XLSB Data Model pair for a new workbook.
    ///
    /// The workbook records are validated and emitted by the BIFF12 writer;
    /// the model part remains opaque and is never refreshed or evaluated.
    pub fn set_data_model(&mut self, value: crate::data_model::Model) -> Result<&mut Self> {
        crate::data_model::serialize_workbook_records(&value.definition)?;
        crate::data_model::validate_model_payload_and_groupings(&value)?;
        if let Some(connections) = self.connections.as_ref() {
            crate::data_model::validate_definition_connections(
                &value.definition,
                Some(connections),
            )?;
        }
        self.data_model = Some(value);
        Ok(self)
    }

    /// Borrow the Data Model scheduled for authoring.
    pub fn data_model(&self) -> Option<&crate::data_model::Model> {
        self.data_model.as_ref()
    }

    /// Remove and return the Data Model scheduled for authoring.
    pub fn clear_data_model(&mut self) -> Option<crate::data_model::Model> {
        self.data_model.take()
    }

    /// Attach a PivotCache definition (MS-XLSB 2.1.7.38) to the workbook.
    ///
    /// The definition is serialized immediately, so model content the
    /// serializer cannot represent losslessly is rejected here rather than
    /// at save time. Returns the allocated workbook PivotCache identifier
    /// (`idSx`), unique per writer. Cache contents are stored verbatim and
    /// are never refreshed, contacted, or evaluated.
    pub fn add_pivot_cache(
        &mut self,
        definition: &crate::package::pivot::PivotCacheDefinition,
    ) -> Result<u32> {
        let bytes = crate::package::pivot::write::write_pivot_cache_definition(definition)?;
        let cache_id = u32::try_from(self.pivot_caches.len())
            .ok()
            .and_then(|next| next.checked_add(1))
            .ok_or_else(|| Error::InvalidFormula("PivotCache identifier overflow".to_string()))?;
        self.pivot_caches.push(AuthoredPivotCache {
            id: cache_id,
            version_created: definition.version_created,
            bytes,
        });
        Ok(cache_id)
    }

    /// Get a mutable reference to a worksheet by index
    pub fn get_worksheet_mut(&mut self, index: usize) -> Option<&mut MutableWorksheet> {
        self.worksheets.get_mut(index)
    }

    /// Get the number of worksheets
    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Number of chart sheets in the workbook.
    pub fn chart_sheet_count(&self) -> usize {
        self.chart_sheets.len()
    }

    /// Mutable chart-sheet access by chart-sheet insertion index.
    pub fn get_chart_sheet_mut(&mut self, index: usize) -> Option<&mut MutableChartSheet> {
        self.chart_sheets.get_mut(index)
    }

    pub(super) fn sheet_name(&self, slot: SheetSlot) -> &str {
        match slot {
            SheetSlot::Worksheet(index) => self.worksheets[index].name(),
            SheetSlot::ChartSheet(index) => self.chart_sheets[index].name(),
        }
    }

    /// Get a reference to the styles writer
    pub fn styles(&self) -> &StylesWriter {
        &self.styles
    }

    /// Get a mutable reference to the styles writer
    pub fn styles_mut(&mut self) -> &mut StylesWriter {
        &mut self.styles
    }
}

impl WorkbookWriter {
    pub(super) fn validate_formula_metadata(&self) -> Result<()> {
        if self.sheet_order.len() > usize::from(u16::MAX) - 2 {
            return Err(Error::InvalidFormula(format!(
                "{} sheets exceed the XLSB extern-sheet index limit",
                self.sheet_order.len()
            )));
        }
        for (index, slot) in self.sheet_order.iter().copied().enumerate() {
            let name = self.sheet_name(slot);
            let name_len = name.encode_utf16().count();
            if name_len == 0
                || name_len > 31
                || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
                || name.starts_with('\'')
                || name.ends_with('\'')
            {
                return Err(Error::InvalidFormula(format!(
                    "sheet name {name:?} does not follow BrtBundleSh grammar"
                )));
            }
            if self.sheet_order[..index]
                .iter()
                .any(|existing| excel_name_eq(self.sheet_name(*existing), name))
            {
                return Err(Error::InvalidFormula(format!(
                    "duplicate sheet name {name:?}"
                )));
            }
        }
        for chart_sheet in &self.chart_sheets {
            chart_sheet.validate()?;
        }
        for (index, named_range) in self.named_ranges.iter().enumerate() {
            if named_range.function {
                return Err(Error::UnsupportedFeature(format!(
                    "macro defined name {} cannot be emitted",
                    named_range.name
                )));
            }
            validate_name(&named_range.name)?;
            if named_range.formula.is_none() {
                return Err(Error::InvalidFormula(format!(
                    "defined name {} has no formula",
                    named_range.name
                )));
            }
            if named_range.sheet_id.is_some_and(|sheet_id| {
                usize::try_from(sheet_id)
                    .ok()
                    .is_none_or(|sheet_id| sheet_id >= self.sheet_order.len())
            }) {
                return Err(Error::InvalidFormula(format!(
                    "defined name {} has invalid sheet scope {:?}",
                    named_range.name, named_range.sheet_id
                )));
            }
            if self.named_ranges[..index].iter().any(|existing| {
                existing.sheet_id == named_range.sheet_id
                    && excel_name_eq(&existing.name, &named_range.name)
            }) {
                return Err(Error::InvalidFormula(format!(
                    "duplicate defined name {:?} in scope {:?}",
                    named_range.name, named_range.sheet_id
                )));
            }
        }
        Ok(())
    }
}

impl Default for WorkbookWriter {
    fn default() -> Self {
        Self::new()
    }
}
