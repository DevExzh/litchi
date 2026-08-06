use super::super::super::formatting::FormattingManager;
use super::super::*;
use crate::EncryptionProfile;
use crate::encryption::{WriterEncryption, validate_writer_encryption};
use crate::error::{Error, Result};
use std::collections::HashMap;
use zeroize::Zeroizing;

impl Writer {
    /// Create a new XLS writer
    pub fn new() -> Self {
        Self {
            worksheets: Vec::new(),
            shared_strings: Vec::new(),
            string_map: HashMap::new(),
            defined_names: Vec::new(),
            defined_name_records: Vec::new(),
            sst_total: 0,
            fmt: FormattingManager::new(),
            workbook_protection: None,
            file_sharing: None,
            use_1904_dates: false,
            calculation_settings: CalculationSettings::default(),
            vba_metadata: None,
            environment_options: WorkbookEnvironmentOptions::default(),
            workbook_window_options: WorkbookWindowOptions::default(),
            function_group_options: FunctionGroupOptions::default(),
            external_workbooks: Vec::new(),
            external_names: Vec::new(),
            add_in_functions: Vec::new(),
            dde_or_ole_links: Vec::new(),
            custom_table_styles: None,
            book_ext: None,
            theme: None,
            mdx_metadata: None,
            real_time_data: Vec::new(),
            web_publications: Vec::new(),
            xf_extensions: Vec::new(),
            style_extensions: Vec::new(),
            toolbar: None,
            encryption: None,
        }
    }

    /// Configure the inert Office Toolbars (`XCB`) stream for the next write.
    ///
    /// The toolbar graph is serialized as metadata only. Controls, macros,
    /// ActiveX payloads, and UI commands are never activated.
    pub fn set_toolbar(&mut self, toolbar: crate::Wrapper<'_>) -> Result<()> {
        let toolbar = toolbar.into_owned();
        toolbar.validate()?;
        self.toolbar = Some(toolbar);
        Ok(())
    }

    /// Remove the optional Office Toolbars (`XCB`) stream from future writes.
    pub fn clear_toolbar(&mut self) {
        self.toolbar = None;
    }

    /// Return the configured inert Office Toolbars metadata, if any.
    pub fn toolbar(&self) -> Option<&crate::Wrapper<'static>> {
        self.toolbar.as_ref()
    }

    /// Configure BIFF8 password-to-open encryption for subsequent writes.
    ///
    /// Validation is atomic: an invalid password or profile leaves the current
    /// encryption configuration unchanged.
    pub fn set_password(
        &mut self,
        password: impl Into<String>,
        profile: EncryptionProfile,
    ) -> Result<()> {
        let password = password.into();
        validate_writer_encryption(&password, profile)?;
        self.encryption = Some(WriterEncryption {
            password: Zeroizing::new(password),
            profile,
        });
        Ok(())
    }

    /// Remove password-to-open encryption from subsequent writes.
    pub fn clear_password(&mut self) {
        self.encryption = None;
    }

    /// Return the configured password-to-open encryption profile.
    pub fn encryption_profile(&self) -> Option<EncryptionProfile> {
        self.encryption.as_ref().map(|value| value.profile)
    }

    /// Add a new worksheet
    ///
    /// # Arguments
    ///
    /// * `name` - Worksheet name (max 31 characters)
    ///
    /// # Returns
    ///
    /// * `Result<usize, Error>` - Worksheet index or error
    pub fn add_worksheet(&mut self, name: &str) -> Result<usize> {
        // Validate worksheet name
        if name.is_empty() || name.len() > 31 {
            return Err(Error::InvalidData(
                "Worksheet name must be 1-31 characters".to_string(),
            ));
        }

        // Check for duplicate names
        if self.worksheets.iter().any(|ws| ws.name == name) {
            return Err(Error::InvalidData(format!(
                "Worksheet '{}' already exists",
                name
            )));
        }

        let index = self.worksheets.len();
        self.worksheets
            .push(WritableWorksheet::new(name.to_string()));
        self.synchronize_workbook_window_selection();
        Ok(index)
    }
}
