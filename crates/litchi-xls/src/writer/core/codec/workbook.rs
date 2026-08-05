use super::super::*;
use crate::error::{Error, Result};

impl Writer {
    /// Set the date system (1900 vs 1904)
    ///
    /// # Arguments
    ///
    /// * `use_1904` - True to use 1904 date system (Mac), false for 1900 (Windows, default)
    pub fn set_1904_dates(&mut self, use_1904: bool) {
        self.use_1904_dates = use_1904;
    }

    pub fn set_workbook_environment(&mut self, options: WorkbookEnvironmentOptions) -> Result<()> {
        if options.refresh_external_data_on_load && !options.template {
            return Err(Error::InvalidData(
                "RefreshAll requires a template workbook".to_string(),
            ));
        }
        if (options.envelope_visible || options.envelope_initialized) && !options.has_envelope {
            return Err(Error::InvalidData(
                "envelope state flags require has_envelope".to_string(),
            ));
        }
        if !(1..=981).contains(&options.default_country_code)
            || !(1..=981).contains(&options.current_country_code)
        {
            return Err(Error::InvalidData(
                "country codes must be 1..=981".to_string(),
            ));
        }
        self.environment_options = options;
        Ok(())
    }

    /// Set the workbook extension flags emitted as a `BookExt` record
    /// (MS-XLS 2.4.23); `None` emits no record.
    pub fn set_book_ext(&mut self, book_ext: Option<crate::BookExt>) {
        self.book_ext = book_ext;
    }

    /// Append a real-time data (RTD) topic emitted as a `RealTimeData`
    /// record (MS-XLS 2.4.214) in the workbook globals.
    ///
    /// When the topic shares a prefix with the previously added topic, set
    /// [`crate::RealTimeData::common_prefix_len`] and store only the
    /// trailing sub-strings in `topic_segments`, matching the on-disk prefix
    /// compression.
    pub fn add_real_time_data(&mut self, topic: crate::RealTimeData) -> Result<()> {
        if let Some(cell) = topic
            .cells
            .iter()
            .find(|cell| usize::from(cell.sheet_index) >= self.worksheets.len())
        {
            return Err(Error::WorksheetNotFound(format!(
                "Sheet {}",
                cell.sheet_index
            )));
        }
        self.real_time_data.push(topic);
        Ok(())
    }

    /// Append a Web page published from the workbook globals, emitted as a
    /// `WebPub` record (MS-XLS 2.4.344).
    pub fn add_web_publication(&mut self, publication: crate::WebPub) -> Result<()> {
        publication.validate_for_write()?;
        self.web_publications.push(publication);
        Ok(())
    }

    /// Append a Web page published from a worksheet, emitted as a `WebPub`
    /// record (MS-XLS 2.4.344) in that sheet's substream.
    pub fn add_sheet_web_publication(
        &mut self,
        sheet: usize,
        publication: crate::WebPub,
    ) -> Result<()> {
        publication.validate_for_write()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.web_publications.push(publication);
        Ok(())
    }

    /// Set a worksheet's default phonetic format and visible phonetic ranges
    /// (PHONETICINFO, MS-XLS 2.4.192); `None` emits no record.
    pub fn set_phonetic_info(
        &mut self,
        sheet: usize,
        phonetic_info: Option<crate::PhoneticInfo>,
    ) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.phonetic_info = phonetic_info;
        Ok(())
    }

    /// Set the document theme emitted as a `Theme` record (MS-XLS 2.4.326);
    /// `None` emits no record. Large custom theme contents are chunked into
    /// ContinueFrt12 records automatically.
    pub fn set_theme(&mut self, theme: Option<crate::Theme>) {
        self.theme = theme;
    }

    /// Set the MDX (OLAP cube) metadata emitted as the workbook globals
    /// `METADATA` production (MS-XLS 2.1); `None` emits no records. Oversized
    /// record payloads are chunked into ContinueFrt12 records automatically.
    pub fn set_mdx_metadata(&mut self, metadata: Option<crate::MdxMetadata>) {
        self.mdx_metadata = metadata;
    }

    /// Set the `XFExt` formatting property extensions (MS-XLS 2.4.355)
    /// emitted after the XF table. Each extension's `xf_index` is validated
    /// against the written XF record count when the workbook is saved.
    pub fn set_xf_extensions(&mut self, xf_extensions: Vec<crate::XfExt>) {
        self.xf_extensions = xf_extensions;
    }

    /// Set the `StyleExt` cell-style extensions (MS-XLS 2.4.270) emitted
    /// after the built-in STYLE records.
    pub fn set_style_extensions(&mut self, style_extensions: Vec<crate::StyleExt>) {
        self.style_extensions = style_extensions;
    }

    pub fn set_workbook_window(&mut self, options: WorkbookWindowOptions) -> Result<()> {
        options.validate_intrinsic()?;
        self.workbook_window_options = options;
        self.synchronize_workbook_window_selection();
        Ok(())
    }

    pub(super) fn synchronize_workbook_window_selection(&mut self) {
        let sheet_count = self.worksheets.len();
        let selected_count = usize::from(self.workbook_window_options.selected_sheet_count);
        let active = usize::from(self.workbook_window_options.active_sheet_index);
        if selected_count == 0 || selected_count > sheet_count || active >= sheet_count {
            return;
        }
        let first_selected = active.min(sheet_count - selected_count);
        let selected_range = first_selected..first_selected + selected_count;
        for (index, worksheet) in self.worksheets.iter_mut().enumerate() {
            worksheet.view.select(selected_range.contains(&index));
        }
    }

    pub fn set_function_groups(&mut self, options: FunctionGroupOptions) -> Result<()> {
        options.validate()?;
        self.function_group_options = options;
        Ok(())
    }

    pub fn add_external_workbook_link(
        &mut self,
        options: ExternalWorkbookOptions,
    ) -> Result<usize> {
        options.validate()?;
        if self.external_workbooks.len()
            + self.dde_or_ole_links.len()
            + usize::from(!self.add_in_functions.is_empty())
            >= 1024
        {
            return Err(Error::InvalidData(
                "external supporting-book count exceeds resource bound".to_string(),
            ));
        }
        let index = self.external_workbooks.len();
        self.external_workbooks.push(options);
        self.external_names.push(Vec::new());
        Ok(index)
    }

    fn external_name_count(&self) -> usize {
        self.external_names.iter().map(Vec::len).sum::<usize>()
            + self.add_in_functions.len()
            + self
                .dde_or_ole_links
                .iter()
                .map(|link| link.items.len())
                .sum::<usize>()
    }

    pub fn add_external_defined_name(
        &mut self,
        external_workbook: usize,
        options: ExternalDefinedNameOptions,
    ) -> Result<usize> {
        let book = self
            .external_workbooks
            .get(external_workbook)
            .ok_or_else(|| {
                Error::InvalidData("external workbook index is out of range".to_string())
            })?;
        options.validate(book.sheets.len())?;
        if self.external_name_count() >= 4096 {
            return Err(Error::InvalidData(
                "external name count exceeds resource bound".to_string(),
            ));
        }
        let names = &mut self.external_names[external_workbook];
        let index = names.len();
        names.push(options);
        Ok(index)
    }

    pub fn add_add_in_function(&mut self, options: AddInFunctionOptions) -> Result<usize> {
        options.validate()?;
        if self.add_in_functions.is_empty()
            && self.external_workbooks.len() + self.dde_or_ole_links.len() >= 1024
        {
            return Err(Error::InvalidData(
                "supporting-book count exceeds resource bound".to_string(),
            ));
        }
        if self.external_name_count() >= 4096 {
            return Err(Error::InvalidData(
                "add-in function count exceeds resource bound".to_string(),
            ));
        }
        let index = self.add_in_functions.len();
        self.add_in_functions.push(options);
        Ok(index)
    }

    pub fn add_dde_or_ole_link(&mut self, options: DdeOrOleLinkOptions) -> Result<usize> {
        options.validate()?;
        if self.external_workbooks.len()
            + self.dde_or_ole_links.len()
            + usize::from(!self.add_in_functions.is_empty())
            >= 1024
        {
            return Err(Error::InvalidData(
                "supporting-book count exceeds resource bound".to_string(),
            ));
        }
        if self
            .external_name_count()
            .checked_add(options.items.len())
            .is_none_or(|count| count > 4096)
        {
            return Err(Error::InvalidData(
                "external name count exceeds resource bound".to_string(),
            ));
        }
        let index = self.dde_or_ole_links.len();
        self.dde_or_ole_links.push(options);
        Ok(index)
    }

    pub fn set_calculation_settings(&mut self, settings: CalculationSettings) -> Result<()> {
        if !(1..=32_767).contains(&settings.maximum_iterations) {
            return Err(Error::InvalidData(
                "maximum calculation iterations must be 1..=32767".to_string(),
            ));
        }
        if !settings.iteration_delta.is_finite() || settings.iteration_delta < 0.0 {
            return Err(Error::InvalidData(
                "calculation iteration delta must be finite and non-negative".to_string(),
            ));
        }
        self.calculation_settings = settings;
        Ok(())
    }

    pub fn set_recalculation_pending(&mut self, sheet: usize, pending: bool) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.formulas_pending_recalculation = pending;
        Ok(())
    }

    pub fn set_scenario_manager(
        &mut self,
        sheet: usize,
        manager: crate::ScenarioManager,
    ) -> Result<()> {
        manager.validate_for_write()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.scenario_manager = Some(manager);
        Ok(())
    }

    pub fn clear_scenario_manager(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.scenario_manager = None;
        Ok(())
    }

    /// Configure an inert BIFF8 data-consolidation directory for a worksheet.
    pub fn set_consolidation(
        &mut self,
        sheet: usize,
        consolidation: crate::Consolidation,
    ) -> Result<()> {
        consolidation.validate_for_write()?;
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.consolidation = Some(consolidation);
        Ok(())
    }

    pub fn clear_consolidation(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.consolidation = None;
        Ok(())
    }

    /// Configure a complete inert VBA project with safe default limits.
    pub fn set_vba(
        &mut self,
        workbook_code_name: &str,
        project: litchi_vba::build::Project,
    ) -> Result<()> {
        self.set_vba_with(workbook_code_name, project, &litchi_vba::Limits::default())
    }

    /// Configure a complete inert VBA project using explicit resource limits.
    ///
    /// Module source is serialized but never compiled, interpreted, or run.
    /// Validation and serialization finish before the writer state is changed.
    pub fn set_vba_with(
        &mut self,
        workbook_code_name: &str,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> Result<()> {
        crate::vba::validate_code_name(workbook_code_name)?;
        let payload = project.finish(limits)?;
        self.put_vba(workbook_code_name, payload)
    }

    /// Configure an already validated and serialized inert VBA project.
    ///
    /// Import standalone CFB bytes through [`litchi_vba::Payload::read`] first.
    pub fn put_vba(
        &mut self,
        workbook_code_name: &str,
        payload: litchi_vba::Payload,
    ) -> Result<()> {
        crate::vba::validate_code_name(workbook_code_name)?;
        self.vba_metadata = Some(VbaWriteMetadata {
            workbook_code_name: workbook_code_name.to_string(),
            project: payload,
        });
        Ok(())
    }

    /// Remove the configured project and all worksheet VBA code names.
    pub fn clear_vba(&mut self) {
        self.vba_metadata = None;
        for worksheet in &mut self.worksheets {
            worksheet.vba_code_name = None;
        }
    }

    /// Whether a complete VBA project is configured for output.
    pub fn has_vba(&self) -> bool {
        self.vba_metadata.is_some()
    }
}
