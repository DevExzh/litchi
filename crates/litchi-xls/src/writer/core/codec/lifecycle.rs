use super::super::super::formatting::FormattingManager;
use super::super::{
    CalculationSettings, FunctionGroupOptions, WorkbookEnvironmentOptions, WorkbookWindowOptions,
    WritableWorksheet, Writer,
};
use crate::encryption::{WriterEncryption, validate_writer_encryption};
use crate::error::{Error, Result};
use crate::records::{SheetNameFault, sheet_name_fault};
use crate::writer::string_limits::WORKSHEET_NAME_UNITS;
use crate::{EncryptionProfile, WeakEncryptionPolicy};
use zeroize::Zeroizing;

impl Writer {
    /// Create a new XLS writer
    #[must_use]
    pub fn new() -> Self {
        Self {
            worksheets: Vec::new(),
            defined_names: Vec::new(),
            defined_name_records: Vec::new(),
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
            xml_map: None,
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
    /// `ActiveX` payloads, and UI commands are never activated.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
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
    #[must_use]
    pub fn toolbar(&self) -> Option<&crate::Wrapper<'static>> {
        self.toolbar.as_ref()
    }

    /// Configure BIFF8 password-to-open encryption for subsequent writes.
    ///
    /// Validation is atomic: an invalid password or profile leaves the current
    /// encryption configuration unchanged.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn set_password(
        &mut self,
        password: impl Into<String>,
        profile: EncryptionProfile,
    ) -> Result<()> {
        if matches!(profile, EncryptionProfile::XorObfuscation) {
            return Err(Error::WeakEncryptionRequiresExplicitPolicy);
        }
        self.configure_password(password.into(), profile)
    }

    /// Configure legacy BIFF8 XOR obfuscation for subsequent writes.
    ///
    /// This is an explicit compatibility downgrade, not encryption suitable
    /// for protecting data. The policy value makes that decision auditable at
    /// each authoring call site. Reading existing XOR-obfuscated files does
    /// not require this policy.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn set_xor_obfuscation_password(
        &mut self,
        password: impl Into<String>,
        _policy: WeakEncryptionPolicy,
    ) -> Result<()> {
        self.configure_password(password.into(), EncryptionProfile::XorObfuscation)
    }

    fn configure_password(&mut self, password: String, profile: EncryptionProfile) -> Result<()> {
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
    #[must_use]
    pub fn encryption_profile(&self) -> Option<EncryptionProfile> {
        self.encryption.as_ref().map(|value| value.profile)
    }

    /// Add a new worksheet
    ///
    /// # Arguments
    ///
    /// * `name` - Worksheet name: 1 through 31 UTF-16 code units, the
    ///   characters BIFF8 counts
    ///
    /// # Returns
    ///
    /// * `Result<usize, Error>` - Worksheet index or error
    /// # Errors
    ///
    /// Returns [`Error::StringTooLong`] for a name longer than 31 UTF-16 code
    /// units; [`Error::InvalidData`] for an empty name, a name containing a
    /// character BIFF8 forbids in worksheet names (NUL, U+0003 or one of
    /// `: \ * ? / [ ]`) or beginning or ending with an apostrophe ([MS-XLS]
    /// 2.4.28), and a duplicate name. A refused name adds nothing.
    pub fn add_worksheet(&mut self, name: &str) -> Result<usize> {
        // Validate worksheet name: the rule litchi's reader enforces.
        match sheet_name_fault(name) {
            None => {},
            Some(SheetNameFault::Empty) => {
                return Err(Error::InvalidData(
                    "Worksheet name must not be empty".to_string(),
                ));
            },
            Some(SheetNameFault::TooLong { utf16_units }) => {
                return Err(Error::StringTooLong {
                    field: "worksheet name",
                    utf16_units,
                    limit: WORKSHEET_NAME_UNITS,
                });
            },
            Some(SheetNameFault::ForbiddenCharacter(character)) => {
                return Err(Error::InvalidData(format!(
                    "Worksheet name {name:?} contains {character:?}, which BIFF8 forbids in worksheet names"
                )));
            },
            Some(SheetNameFault::EdgeApostrophe) => {
                return Err(Error::InvalidData(format!(
                    "Worksheet name {name:?} begins or ends with an apostrophe"
                )));
            },
        }

        // Check for duplicate names
        let normalized_name = name.to_lowercase();
        if self
            .worksheets
            .iter()
            .any(|ws| ws.name.to_lowercase() == normalized_name)
        {
            return Err(Error::InvalidData(format!(
                "Worksheet '{name}' already exists"
            )));
        }

        let index = self.worksheets.len();
        self.worksheets
            .push(WritableWorksheet::new(name.to_string()));
        self.synchronize_workbook_window_selection();
        Ok(index)
    }
}
