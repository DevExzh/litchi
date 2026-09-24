//! Bounded, inert ODF database connection and driver settings.
//!
//! These values describe the XML configuration carried by an ODB package.  A
//! settings value is deliberately independent of any database driver: reading
//! it never resolves a driver, asks for a password, opens a socket, or runs a
//! command.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event, attributes::Attribute},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};
#[cfg(test)]
use std::sync::{
    Mutex,
    atomic::{AtomicUsize, Ordering},
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DATABASE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_SETTINGS: usize = 65_536;
const MAX_SETTING_VALUES: usize = 65_536;
const MAX_TABLE_SETTINGS: usize = 65_536;
const MAX_FILTER_PATTERNS: usize = 65_536;
const MAX_REFERENCE_BYTES: usize = 128;
const MAX_DEPTH: usize = 256;

#[cfg(test)]
static PARSE_COUNT: AtomicUsize = AtomicUsize::new(0);
#[cfg(test)]
static PARSE_TEST_LOCK: Mutex<()> = Mutex::new(());

/// A login declaration from `db:connection-data`.
///
/// Password values are intentionally not modeled.  `is-password-required` is
/// only an inert producer hint; this crate never obtains or tests credentials.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct LoginSettings {
    user_name: Option<String>,
    use_system_user: Option<bool>,
    password_required: Option<bool>,
    login_timeout: Option<u64>,
}

impl LoginSettings {
    /// Creates an empty login declaration.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            user_name: None,
            use_system_user: None,
            password_required: None,
            login_timeout: None,
        }
    }

    /// Sets the inert user-name hint.
    #[must_use]
    pub fn with_user_name(mut self, value: impl Into<String>) -> Self {
        self.user_name = Some(value.into());
        self.use_system_user = None;
        self
    }

    /// Sets the producer's system-user flag.
    #[must_use]
    pub fn with_use_system_user(mut self, value: Option<bool>) -> Self {
        self.use_system_user = value;
        if value.is_some() {
            self.user_name = None;
        }
        self
    }

    /// Sets whether a host would request a password.  No password is stored.
    #[must_use]
    pub const fn with_password_required(mut self, value: Option<bool>) -> Self {
        self.password_required = value;
        self
    }

    /// Sets the positive, inert login timeout in seconds.
    ///
    /// The typed projection is bounded to `u64`; larger schema-valid
    /// `positiveInteger` values are refused rather than truncated.
    #[must_use]
    pub const fn with_login_timeout(mut self, value: Option<u64>) -> Self {
        self.login_timeout = value;
        self
    }

    #[must_use]
    pub fn user_name(&self) -> Option<&str> {
        self.user_name.as_deref()
    }

    #[must_use]
    pub const fn use_system_user(&self) -> Option<bool> {
        self.use_system_user
    }

    #[must_use]
    pub const fn password_required(&self) -> Option<bool> {
        self.password_required
    }

    #[must_use]
    pub const fn login_timeout(&self) -> Option<u64> {
        self.login_timeout
    }
}

/// The bounded inert auto-increment statements under driver settings.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct AutoIncrementSettings {
    additional_column_statement: Option<String>,
    row_retrieving_statement: Option<String>,
}

impl AutoIncrementSettings {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            additional_column_statement: None,
            row_retrieving_statement: None,
        }
    }

    #[must_use]
    pub fn with_additional_column_statement(mut self, value: impl Into<String>) -> Self {
        self.additional_column_statement = Some(value.into());
        self
    }

    #[must_use]
    pub fn with_row_retrieving_statement(mut self, value: impl Into<String>) -> Self {
        self.row_retrieving_statement = Some(value.into());
        self
    }

    #[must_use]
    pub fn additional_column_statement(&self) -> Option<&str> {
        self.additional_column_statement.as_deref()
    }

    #[must_use]
    pub fn row_retrieving_statement(&self) -> Option<&str> {
        self.row_retrieving_statement.as_deref()
    }
}

/// Delimiter hints used by a driver or one table setting.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct DelimiterSettings {
    field: Option<String>,
    string: Option<String>,
    decimal: Option<String>,
    thousand: Option<String>,
}

impl DelimiterSettings {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            field: None,
            string: None,
            decimal: None,
            thousand: None,
        }
    }

    #[must_use]
    pub fn with_field(mut self, value: impl Into<String>) -> Self {
        self.field = Some(value.into());
        self
    }

    #[must_use]
    pub fn with_string(mut self, value: impl Into<String>) -> Self {
        self.string = Some(value.into());
        self
    }

    #[must_use]
    pub fn with_decimal(mut self, value: impl Into<String>) -> Self {
        self.decimal = Some(value.into());
        self
    }

    #[must_use]
    pub fn with_thousand(mut self, value: impl Into<String>) -> Self {
        self.thousand = Some(value.into());
        self
    }

    #[must_use]
    pub fn field(&self) -> Option<&str> {
        self.field.as_deref()
    }

    #[must_use]
    pub fn string(&self) -> Option<&str> {
        self.string.as_deref()
    }

    #[must_use]
    pub fn decimal(&self) -> Option<&str> {
        self.decimal.as_deref()
    }

    #[must_use]
    pub fn thousand(&self) -> Option<&str> {
        self.thousand.as_deref()
    }
}

/// An inert character-set declaration.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CharacterSetSettings {
    encoding: Option<String>,
}

impl CharacterSetSettings {
    #[must_use]
    pub const fn new() -> Self {
        Self { encoding: None }
    }

    #[must_use]
    pub fn with_encoding(mut self, value: impl Into<String>) -> Self {
        self.encoding = Some(value.into());
        self
    }

    #[must_use]
    pub fn encoding(&self) -> Option<&str> {
        self.encoding.as_deref()
    }

    pub(crate) fn validate(&self) -> Result<()> {
        // ODF 1.4 textEncoding: [A-Za-z][A-Za-z0-9._\-]*.
        if let Some(value) = self.encoding() {
            let mut bytes = value.bytes();
            if !bytes.next().is_some_and(|byte| byte.is_ascii_alphabetic())
                || !bytes
                    .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_' | b'-'))
            {
                return Err(Error::InvalidFormat(
                    "invalid database character-set encoding".into(),
                ));
            }
        }
        Ok(())
    }
}

/// One table-scoped driver setting.  ODF intentionally does not attach a
/// table name to this element; the order and multiplicity are retained.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TableSetting {
    first_row_header_line: Option<bool>,
    show_deleted: Option<bool>,
    delimiter: Option<DelimiterSettings>,
    character_set: Option<CharacterSetSettings>,
}

impl TableSetting {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            first_row_header_line: None,
            show_deleted: None,
            delimiter: None,
            character_set: None,
        }
    }

    #[must_use]
    pub const fn with_first_row_header_line(mut self, value: Option<bool>) -> Self {
        self.first_row_header_line = value;
        self
    }

    #[must_use]
    pub const fn with_show_deleted(mut self, value: Option<bool>) -> Self {
        self.show_deleted = value;
        self
    }

    #[must_use]
    pub fn with_delimiter(mut self, value: Option<DelimiterSettings>) -> Self {
        self.delimiter = value;
        self
    }

    #[must_use]
    pub fn with_character_set(mut self, value: Option<CharacterSetSettings>) -> Self {
        self.character_set = value;
        self
    }

    #[must_use]
    pub const fn first_row_header_line(&self) -> Option<bool> {
        self.first_row_header_line
    }

    #[must_use]
    pub const fn show_deleted(&self) -> Option<bool> {
        self.show_deleted
    }

    #[must_use]
    pub fn delimiter(&self) -> Option<&DelimiterSettings> {
        self.delimiter.as_ref()
    }

    #[must_use]
    pub fn character_set(&self) -> Option<&CharacterSetSettings> {
        self.character_set.as_ref()
    }
}

/// A bounded driver-settings subtree.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct DriverSettings {
    show_deleted: Option<bool>,
    system_driver_settings: Option<String>,
    base_dn: Option<String>,
    first_row_header_line: Option<bool>,
    parameter_name_substitution: Option<bool>,
    auto_increment: Option<AutoIncrementSettings>,
    delimiter: Option<DelimiterSettings>,
    character_set: Option<CharacterSetSettings>,
    table_settings: Vec<TableSetting>,
}

impl DriverSettings {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            show_deleted: None,
            system_driver_settings: None,
            base_dn: None,
            first_row_header_line: None,
            parameter_name_substitution: None,
            auto_increment: None,
            delimiter: None,
            character_set: None,
            table_settings: Vec::new(),
        }
    }

    #[must_use]
    pub const fn with_show_deleted(mut self, value: Option<bool>) -> Self {
        self.show_deleted = value;
        self
    }

    #[must_use]
    pub fn with_system_driver_settings(mut self, value: impl Into<String>) -> Self {
        self.system_driver_settings = Some(value.into());
        self
    }

    #[must_use]
    pub fn with_base_dn(mut self, value: impl Into<String>) -> Self {
        self.base_dn = Some(value.into());
        self
    }

    #[must_use]
    pub const fn with_first_row_header_line(mut self, value: Option<bool>) -> Self {
        self.first_row_header_line = value;
        self
    }

    #[must_use]
    pub const fn with_parameter_name_substitution(mut self, value: Option<bool>) -> Self {
        self.parameter_name_substitution = value;
        self
    }

    #[must_use]
    pub fn with_auto_increment(mut self, value: Option<AutoIncrementSettings>) -> Self {
        self.auto_increment = value;
        self
    }

    #[must_use]
    pub fn with_delimiter(mut self, value: Option<DelimiterSettings>) -> Self {
        self.delimiter = value;
        self
    }

    #[must_use]
    pub fn with_character_set(mut self, value: Option<CharacterSetSettings>) -> Self {
        self.character_set = value;
        self
    }

    #[must_use]
    pub fn with_table_settings(mut self, value: Vec<TableSetting>) -> Self {
        self.table_settings = value;
        self
    }

    #[must_use]
    pub const fn show_deleted(&self) -> Option<bool> {
        self.show_deleted
    }

    #[must_use]
    pub fn system_driver_settings(&self) -> Option<&str> {
        self.system_driver_settings.as_deref()
    }

    #[must_use]
    pub fn base_dn(&self) -> Option<&str> {
        self.base_dn.as_deref()
    }

    #[must_use]
    pub const fn first_row_header_line(&self) -> Option<bool> {
        self.first_row_header_line
    }

    #[must_use]
    pub const fn parameter_name_substitution(&self) -> Option<bool> {
        self.parameter_name_substitution
    }

    #[must_use]
    pub fn auto_increment(&self) -> Option<&AutoIncrementSettings> {
        self.auto_increment.as_ref()
    }

    #[must_use]
    pub fn delimiter(&self) -> Option<&DelimiterSettings> {
        self.delimiter.as_ref()
    }

    #[must_use]
    pub fn character_set(&self) -> Option<&CharacterSetSettings> {
        self.character_set.as_ref()
    }

    #[must_use]
    pub fn table_settings(&self) -> &[TableSetting] {
        &self.table_settings
    }
}

/// The comparison mode vocabulary from ODF §12.15.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum BooleanComparisonMode {
    EqualInteger,
    IsBoolean,
    EqualBoolean,
    EqualUseOnlyZero,
}

impl BooleanComparisonMode {
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::EqualInteger => "equal-integer",
            Self::IsBoolean => "is-boolean",
            Self::EqualBoolean => "equal-boolean",
            Self::EqualUseOnlyZero => "equal-use-only-zero",
        }
    }
}

/// A table include/exclude filter.  Patterns are inert producer strings.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TableFilter {
    include: Vec<String>,
    exclude: Vec<String>,
}

impl TableFilter {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            include: Vec::new(),
            exclude: Vec::new(),
        }
    }

    #[must_use]
    pub fn with_include(mut self, values: Vec<String>) -> Self {
        self.include = values;
        self
    }

    #[must_use]
    pub fn with_exclude(mut self, values: Vec<String>) -> Self {
        self.exclude = values;
        self
    }

    #[must_use]
    pub fn include(&self) -> &[String] {
        &self.include
    }

    #[must_use]
    pub fn exclude(&self) -> &[String] {
        &self.exclude
    }
}

/// A table-type filter from application connection settings.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TableTypeFilter {
    table_types: Vec<String>,
}

impl TableTypeFilter {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            table_types: Vec::new(),
        }
    }

    #[must_use]
    pub fn with_table_types(mut self, values: Vec<String>) -> Self {
        self.table_types = values;
        self
    }

    #[must_use]
    pub fn table_types(&self) -> &[String] {
        &self.table_types
    }
}

/// The six ODF data-source setting type tokens.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum DataSourceSettingType {
    Boolean,
    Short,
    Int,
    Long,
    Double,
    String,
}

impl DataSourceSettingType {
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Boolean => "boolean",
            Self::Short => "short",
            Self::Int => "int",
            Self::Long => "long",
            Self::Double => "double",
            Self::String => "string",
        }
    }
}

/// One typed-but-inert data-source property.  Values remain lexical because
/// this crate does not own a driver's conversion rules.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DataSourceSetting {
    name: String,
    setting_type: DataSourceSettingType,
    is_list: Option<bool>,
    values: Vec<String>,
}

impl DataSourceSetting {
    #[must_use]
    pub fn new(name: impl Into<String>, setting_type: DataSourceSettingType) -> Self {
        Self {
            name: name.into(),
            setting_type,
            is_list: None,
            values: Vec::new(),
        }
    }

    #[must_use]
    pub const fn with_is_list(mut self, value: Option<bool>) -> Self {
        self.is_list = value;
        self
    }

    #[must_use]
    pub fn with_values(mut self, values: Vec<String>) -> Self {
        self.values = values;
        self
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub const fn setting_type(&self) -> DataSourceSettingType {
        self.setting_type
    }

    #[must_use]
    pub const fn is_list(&self) -> Option<bool> {
        self.is_list
    }

    #[must_use]
    pub fn values(&self) -> &[String] {
        &self.values
    }
}

/// Application-level ODF connection settings, with inert table filters.
///
/// `max-row-count` is projected to a bounded `i64`; values outside that
/// representation are refused rather than truncated.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ApplicationConnectionSettings {
    table_name_length_limited: Option<bool>,
    enable_sql92_check: Option<bool>,
    append_table_alias_name: Option<bool>,
    ignore_driver_privileges: Option<bool>,
    boolean_comparison_mode: Option<BooleanComparisonMode>,
    use_catalog: Option<bool>,
    max_row_count: Option<i64>,
    suppress_version_columns: Option<bool>,
    table_filter: Option<TableFilter>,
    table_type_filter: Option<TableTypeFilter>,
    data_source_settings: Vec<DataSourceSetting>,
}

impl ApplicationConnectionSettings {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            table_name_length_limited: None,
            enable_sql92_check: None,
            append_table_alias_name: None,
            ignore_driver_privileges: None,
            boolean_comparison_mode: None,
            use_catalog: None,
            max_row_count: None,
            suppress_version_columns: None,
            table_filter: None,
            table_type_filter: None,
            data_source_settings: Vec::new(),
        }
    }

    #[must_use]
    pub const fn with_table_name_length_limited(mut self, value: Option<bool>) -> Self {
        self.table_name_length_limited = value;
        self
    }

    #[must_use]
    pub const fn with_enable_sql92_check(mut self, value: Option<bool>) -> Self {
        self.enable_sql92_check = value;
        self
    }

    #[must_use]
    pub const fn with_append_table_alias_name(mut self, value: Option<bool>) -> Self {
        self.append_table_alias_name = value;
        self
    }

    #[must_use]
    pub const fn with_ignore_driver_privileges(mut self, value: Option<bool>) -> Self {
        self.ignore_driver_privileges = value;
        self
    }

    #[must_use]
    pub const fn with_boolean_comparison_mode(
        mut self,
        value: Option<BooleanComparisonMode>,
    ) -> Self {
        self.boolean_comparison_mode = value;
        self
    }

    #[must_use]
    pub const fn with_use_catalog(mut self, value: Option<bool>) -> Self {
        self.use_catalog = value;
        self
    }

    /// Sets the bounded inert maximum row count.
    #[must_use]
    pub const fn with_max_row_count(mut self, value: Option<i64>) -> Self {
        self.max_row_count = value;
        self
    }

    #[must_use]
    pub const fn with_suppress_version_columns(mut self, value: Option<bool>) -> Self {
        self.suppress_version_columns = value;
        self
    }

    #[must_use]
    pub fn with_table_filter(mut self, value: Option<TableFilter>) -> Self {
        self.table_filter = value;
        self
    }

    #[must_use]
    pub fn with_table_type_filter(mut self, value: Option<TableTypeFilter>) -> Self {
        self.table_type_filter = value;
        self
    }

    #[must_use]
    pub fn with_data_source_settings(mut self, value: Vec<DataSourceSetting>) -> Self {
        self.data_source_settings = value;
        self
    }

    #[must_use]
    pub const fn table_name_length_limited(&self) -> Option<bool> {
        self.table_name_length_limited
    }

    #[must_use]
    pub const fn enable_sql92_check(&self) -> Option<bool> {
        self.enable_sql92_check
    }

    #[must_use]
    pub const fn append_table_alias_name(&self) -> Option<bool> {
        self.append_table_alias_name
    }

    #[must_use]
    pub const fn ignore_driver_privileges(&self) -> Option<bool> {
        self.ignore_driver_privileges
    }

    #[must_use]
    pub const fn boolean_comparison_mode(&self) -> Option<BooleanComparisonMode> {
        self.boolean_comparison_mode
    }

    #[must_use]
    pub const fn use_catalog(&self) -> Option<bool> {
        self.use_catalog
    }

    #[must_use]
    pub const fn max_row_count(&self) -> Option<i64> {
        self.max_row_count
    }

    #[must_use]
    pub const fn suppress_version_columns(&self) -> Option<bool> {
        self.suppress_version_columns
    }

    #[must_use]
    pub fn table_filter(&self) -> Option<&TableFilter> {
        self.table_filter.as_ref()
    }

    #[must_use]
    pub fn table_type_filter(&self) -> Option<&TableTypeFilter> {
        self.table_type_filter.as_ref()
    }

    #[must_use]
    pub fn data_source_settings(&self) -> &[DataSourceSetting] {
        &self.data_source_settings
    }
}

/// All modeled settings below one ODF database data-source.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct DatabaseSettings {
    login: Option<LoginSettings>,
    driver: Option<DriverSettings>,
    application_connection: Option<ApplicationConnectionSettings>,
}

impl DatabaseSettings {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            login: None,
            driver: None,
            application_connection: None,
        }
    }

    #[must_use]
    pub fn with_login(mut self, value: Option<LoginSettings>) -> Self {
        self.login = value;
        self
    }

    #[must_use]
    pub fn with_driver(mut self, value: Option<DriverSettings>) -> Self {
        self.driver = value;
        self
    }

    #[must_use]
    pub fn with_application_connection(
        mut self,
        value: Option<ApplicationConnectionSettings>,
    ) -> Self {
        self.application_connection = value;
        self
    }

    #[must_use]
    pub fn login(&self) -> Option<&LoginSettings> {
        self.login.as_ref()
    }

    #[must_use]
    pub fn driver(&self) -> Option<&DriverSettings> {
        self.driver.as_ref()
    }

    #[must_use]
    pub fn application_connection(&self) -> Option<&ApplicationConnectionSettings> {
        self.application_connection.as_ref()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.login.is_none() && self.driver.is_none() && self.application_connection.is_none()
    }
}

/// Short ergonomic alias for [`DatabaseSettings`].
pub type Settings = DatabaseSettings;

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Database,
    Xlink,
    Other,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Element {
    Document,
    Body,
    Database,
    DataSource,
    ConnectionData,
    DatabaseDescription,
    FileBasedDatabase,
    ServerDatabase,
    ConnectionResource,
    Login,
    Driver,
    AutoIncrement,
    Delimiter,
    CharacterSet,
    TableSettings,
    TableSetting,
    Application,
    TableFilter,
    IncludeFilter,
    ExcludeFilter,
    Pattern,
    TableTypeFilter,
    TableType,
    DataSourceSettings,
    DataSourceSetting,
    DataSourceSettingValue,
    Other,
}

struct Frame {
    element: Element,
    text: String,
    child_count: usize,
    has_child: bool,
    has_opaque_content: bool,
    last_child_order: u8,
    connection_target_seen: bool,
    database_target_seen: bool,
    connection_data_seen: bool,
    login: Option<LoginSettings>,
    driver: Option<DriverSettings>,
    auto_increment: Option<AutoIncrementSettings>,
    delimiter: Option<DelimiterSettings>,
    character_set: Option<CharacterSetSettings>,
    table_setting: Option<TableSetting>,
    application: Option<ApplicationConnectionSettings>,
    table_filter: Option<TableFilter>,
    table_type_filter: Option<TableTypeFilter>,
    data_source_setting: Option<DataSourceSetting>,
    data_source_settings: Vec<DataSourceSetting>,
    data_source_settings_container_seen: bool,
}

impl Frame {
    fn new(element: Element) -> Self {
        Self {
            element,
            text: String::new(),
            child_count: 0,
            has_child: false,
            has_opaque_content: false,
            last_child_order: 0,
            connection_target_seen: false,
            database_target_seen: false,
            connection_data_seen: false,
            login: None,
            driver: None,
            auto_increment: None,
            delimiter: None,
            character_set: None,
            table_setting: None,
            application: None,
            table_filter: None,
            table_type_filter: None,
            data_source_setting: None,
            data_source_settings: Vec::new(),
            data_source_settings_container_seen: false,
        }
    }
}

/// Parses the known ODF settings vocabulary under the unique data source.
pub(crate) fn parse(
    source: &str,
    max_attribute_bytes: usize,
    max_settings: usize,
    max_setting_values: usize,
    max_table_settings: usize,
    max_filter_patterns: usize,
) -> Result<DatabaseSettings> {
    #[cfg(test)]
    PARSE_COUNT.fetch_add(1, Ordering::Relaxed);

    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut settings = DatabaseSettings::new();
    let mut data_source_seen = false;
    let mut setting_count = 0usize;
    let mut value_count = 0usize;
    let mut table_setting_count = 0usize;
    let mut pattern_count = 0usize;

    loop {
        let (resolved, raw_event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(&format!("invalid ODB settings XML: {error}")))?;
        let namespace = namespace_kind(&resolved);
        let is_eof = matches!(raw_event, Event::Eof);
        match raw_event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("ODB settings XML nesting exceeds the limit"));
                }
                let kind = classify(stack.last().map(|frame| frame.element), namespace, &element);
                validate_reserved_element(
                    stack.last().map(|frame| frame.element),
                    namespace,
                    element.name().as_ref(),
                    element.local_name().as_ref(),
                )?;
                validate_database_child(
                    stack.last().map(|frame| frame.element),
                    namespace,
                    element.local_name().as_ref(),
                )?;
                validate_child_order(&mut stack, namespace, element.local_name().as_ref())?;
                if let Some(parent) = stack.last_mut() {
                    parent.has_child = true;
                }
                if kind == Element::Other {
                    if let Some(parent) = stack.last_mut() {
                        parent.has_opaque_content = true;
                    }
                }
                let mut frame = Frame::new(kind);
                parse_attributes(&reader, &element, kind, &mut frame, max_attribute_bytes)?;
                if kind == Element::DataSource {
                    if data_source_seen {
                        return Err(invalid("ODB settings has multiple data sources"));
                    }
                    data_source_seen = true;
                }
                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "ODB settings element stack",
                    source,
                })?;
                stack.push(frame);
            },
            Event::Empty(element) => {
                let kind = classify(stack.last().map(|frame| frame.element), namespace, &element);
                validate_reserved_element(
                    stack.last().map(|frame| frame.element),
                    namespace,
                    element.name().as_ref(),
                    element.local_name().as_ref(),
                )?;
                validate_database_child(
                    stack.last().map(|frame| frame.element),
                    namespace,
                    element.local_name().as_ref(),
                )?;
                validate_child_order(&mut stack, namespace, element.local_name().as_ref())?;
                if let Some(parent) = stack.last_mut() {
                    parent.has_child = true;
                }
                if kind == Element::Other {
                    if let Some(parent) = stack.last_mut() {
                        parent.has_opaque_content = true;
                    }
                }
                let mut frame = Frame::new(kind);
                parse_attributes(&reader, &element, kind, &mut frame, max_attribute_bytes)?;
                if kind == Element::DataSource {
                    if data_source_seen {
                        return Err(invalid("ODB settings has multiple data sources"));
                    }
                    data_source_seen = true;
                }
                finalize(
                    frame,
                    &mut stack,
                    &mut settings,
                    &mut setting_count,
                    &mut value_count,
                    &mut table_setting_count,
                    &mut pattern_count,
                    max_attribute_bytes,
                    max_settings,
                    max_setting_values,
                    max_table_settings,
                    max_filter_patterns,
                )?;
            },
            Event::Text(text) => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| invalid(&format!("invalid ODB settings text: {error}")))?;
                handle_text(&mut stack, value.as_ref(), max_attribute_bytes)?;
            },
            Event::CData(text) => {
                let value = text
                    .decode()
                    .map_err(|error| invalid(&format!("invalid ODB settings CDATA: {error}")))?;
                handle_text(&mut stack, value.as_ref(), max_attribute_bytes)?;
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("ODB settings has an unmatched closing tag"))?;
                finalize(
                    frame,
                    &mut stack,
                    &mut settings,
                    &mut setting_count,
                    &mut value_count,
                    &mut table_setting_count,
                    &mut pattern_count,
                    max_attribute_bytes,
                    max_settings,
                    max_setting_values,
                    max_table_settings,
                    max_filter_patterns,
                )?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not permitted in ODB settings")),
            Event::GeneralRef(reference) => {
                if stack
                    .last()
                    .is_some_and(|frame| frame.element != Element::Other)
                {
                    let decoded = decode_reference(&reference)?;
                    handle_text(&mut stack, &decoded, max_attribute_bytes)?;
                }
            },
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {
                if let Some(frame) = stack.last_mut() {
                    frame.has_opaque_content = true;
                }
            },
            Event::Eof => {},
        }
        buffer.clear();
        if is_eof {
            break;
        }
    }
    if !stack.is_empty() {
        return Err(invalid("ODB settings XML is incomplete"));
    }
    if !data_source_seen {
        return Err(invalid("ODB settings has no data source"));
    }
    Ok(settings)
}

#[cfg(test)]
pub(crate) fn reset_parse_count() {
    PARSE_COUNT.store(0, Ordering::Relaxed);
}

#[cfg(test)]
pub(crate) fn parse_count() -> usize {
    PARSE_COUNT.load(Ordering::Relaxed)
}

#[cfg(test)]
pub(crate) fn parse_test_lock() -> &'static Mutex<()> {
    &PARSE_TEST_LOCK
}

fn classify(
    parent: Option<Element>,
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
) -> Element {
    let local = element.local_name();
    if namespace == NamespaceKind::Other {
        return Element::Other;
    }
    if namespace == NamespaceKind::Database {
        return match (parent, local.as_ref()) {
            (Some(Element::Database), b"data-source") => Element::DataSource,
            (Some(Element::DataSource), b"connection-data") => Element::ConnectionData,
            (Some(Element::ConnectionData), b"database-description") => {
                Element::DatabaseDescription
            },
            (Some(Element::DatabaseDescription), b"file-based-database") => {
                Element::FileBasedDatabase
            },
            (Some(Element::DatabaseDescription), b"server-database") => Element::ServerDatabase,
            (Some(Element::ConnectionData), b"connection-resource") => Element::ConnectionResource,
            (Some(Element::ConnectionData), b"login") => Element::Login,
            (Some(Element::DataSource), b"driver-settings") => Element::Driver,
            (Some(Element::Driver), b"auto-increment") => Element::AutoIncrement,
            (Some(Element::Driver | Element::TableSetting), b"delimiter") => Element::Delimiter,
            (Some(Element::Driver | Element::TableSetting), b"character-set") => {
                Element::CharacterSet
            },
            (Some(Element::Driver), b"table-settings") => Element::TableSettings,
            (Some(Element::TableSettings), b"table-setting") => Element::TableSetting,
            (Some(Element::DataSource), b"application-connection-settings") => Element::Application,
            (Some(Element::Application), b"table-filter") => Element::TableFilter,
            (Some(Element::TableFilter), b"table-include-filter") => Element::IncludeFilter,
            (Some(Element::TableFilter), b"table-exclude-filter") => Element::ExcludeFilter,
            (Some(Element::IncludeFilter | Element::ExcludeFilter), b"table-filter-pattern") => {
                Element::Pattern
            },
            (Some(Element::Application), b"table-type-filter") => Element::TableTypeFilter,
            (Some(Element::TableTypeFilter), b"table-type") => Element::TableType,
            (Some(Element::Application), b"data-source-settings") => Element::DataSourceSettings,
            (Some(Element::DataSourceSettings), b"data-source-setting") => {
                Element::DataSourceSetting
            },
            (Some(Element::DataSourceSetting), b"data-source-setting-value") => {
                Element::DataSourceSettingValue
            },
            _ => Element::Other,
        };
    }
    if namespace != NamespaceKind::Office {
        return Element::Other;
    }
    match (parent, local.as_ref()) {
        (None, b"document-content") => Element::Document,
        (Some(Element::Document), b"body") => Element::Body,
        (Some(Element::Body), b"database") => Element::Database,
        _ => Element::Other,
    }
}

fn validate_database_child(
    parent: Option<Element>,
    namespace: NamespaceKind,
    local: &[u8],
) -> Result<()> {
    if namespace != NamespaceKind::Database {
        return Ok(());
    }
    let reject = match parent {
        Some(Element::DataSource) => !matches!(
            local,
            b"connection-data" | b"driver-settings" | b"application-connection-settings"
        ),
        Some(Element::ConnectionData) => !matches!(
            local,
            b"connection-resource" | b"database-description" | b"login"
        ),
        Some(Element::DatabaseDescription) => {
            !matches!(local, b"file-based-database" | b"server-database")
        },
        Some(
            Element::ConnectionResource | Element::FileBasedDatabase | Element::ServerDatabase,
        ) => true,
        Some(Element::Login)
        | Some(Element::AutoIncrement)
        | Some(Element::Delimiter)
        | Some(Element::CharacterSet)
        | Some(Element::Pattern)
        | Some(Element::TableType)
        | Some(Element::DataSourceSettingValue) => true,
        Some(Element::Driver) => !matches!(
            local,
            b"auto-increment"
                | b"character-set"
                | b"delimiter"
                | b"font-charset"
                | b"table-settings"
        ),
        Some(Element::TableSettings) => local != b"table-setting",
        Some(Element::TableSetting) => !matches!(local, b"delimiter" | b"character-set"),
        Some(Element::Application) => !matches!(
            local,
            b"table-filter" | b"table-type-filter" | b"data-source-settings"
        ),
        Some(Element::TableFilter) => {
            !matches!(local, b"table-include-filter" | b"table-exclude-filter")
        },
        Some(Element::IncludeFilter | Element::ExcludeFilter) => local != b"table-filter-pattern",
        Some(Element::TableTypeFilter) => local != b"table-type",
        Some(Element::DataSourceSettings) => local != b"data-source-setting",
        Some(Element::DataSourceSetting) => local != b"data-source-setting-value",
        _ => false,
    };
    if reject {
        return Err(invalid(
            "ODB settings contains an invalid nested database element",
        ));
    }
    Ok(())
}

fn validate_reserved_element(
    parent: Option<Element>,
    namespace: NamespaceKind,
    raw_name: &[u8],
    local: &[u8],
) -> Result<()> {
    if namespace == NamespaceKind::Other && reserved_prefix(raw_name) {
        return Err(invalid("ODB reserved namespace prefix is not bound"));
    }

    if matches!(
        parent,
        Some(Element::ConnectionResource | Element::FileBasedDatabase | Element::ServerDatabase)
    ) {
        return Err(invalid("ODB connection target must have empty content"));
    }

    match namespace {
        NamespaceKind::Office => {
            let allowed = match parent {
                None => local == b"document-content",
                Some(Element::Document) => matches!(
                    local,
                    b"automatic-styles" | b"body" | b"font-face-decls" | b"scripts"
                ),
                Some(Element::Body) => local == b"database",
                _ => false,
            };
            if !allowed {
                return Err(invalid("ODB settings contains an invalid office element"));
            }
        },
        NamespaceKind::Xlink => {
            return Err(invalid("ODB settings contains an invalid xlink element"));
        },
        NamespaceKind::Database if !is_known_database_element(local) => {
            return Err(invalid("ODB settings contains an unknown database element"));
        },
        NamespaceKind::Database | NamespaceKind::Other => {},
    }
    Ok(())
}

fn is_known_database_element(local: &[u8]) -> bool {
    matches!(
        local,
        b"application-connection-settings"
            | b"auto-increment"
            | b"character-set"
            | b"column"
            | b"column-definition"
            | b"column-definitions"
            | b"columns"
            | b"component"
            | b"component-collection"
            | b"connection-data"
            | b"connection-resource"
            | b"data-source"
            | b"data-source-setting"
            | b"data-source-setting-value"
            | b"data-source-settings"
            | b"database-description"
            | b"delimiter"
            | b"driver-settings"
            | b"file-based-database"
            | b"font-charset"
            | b"filter-statement"
            | b"forms"
            | b"index"
            | b"index-column"
            | b"index-columns"
            | b"indices"
            | b"key"
            | b"key-column"
            | b"key-columns"
            | b"keys"
            | b"login"
            | b"order-statement"
            | b"queries"
            | b"query"
            | b"query-collection"
            | b"reports"
            | b"schema-definition"
            | b"server-database"
            | b"table-definition"
            | b"table-definitions"
            | b"table-exclude-filter"
            | b"table-filter"
            | b"table-filter-pattern"
            | b"table-include-filter"
            | b"table-representation"
            | b"table-representations"
            | b"table-setting"
            | b"table-settings"
            | b"table-type"
            | b"table-type-filter"
            | b"update-table"
    )
}

fn validate_child_order(stack: &mut [Frame], namespace: NamespaceKind, local: &[u8]) -> Result<()> {
    if namespace != NamespaceKind::Database {
        return Ok(());
    }
    let Some(parent) = stack.last_mut() else {
        return Ok(());
    };
    if parent.element == Element::ConnectionData {
        if matches!(local, b"connection-resource" | b"database-description") {
            if parent.connection_target_seen {
                return Err(invalid(
                    "ODB connection-data has multiple connection targets",
                ));
            }
            if parent.last_child_order > 1 {
                return Err(invalid("ODB connection-data child order is invalid"));
            }
            parent.connection_target_seen = true;
            parent.last_child_order = 1;
            return Ok(());
        }
        if local == b"login" {
            if parent.last_child_order > 1 {
                return Err(invalid("ODB connection-data child order is invalid"));
            }
            parent.last_child_order = 2;
        }
        return Ok(());
    }
    if parent.element == Element::DatabaseDescription
        && matches!(local, b"file-based-database" | b"server-database")
    {
        if parent.database_target_seen {
            return Err(invalid(
                "ODB database-description has multiple database targets",
            ));
        }
        parent.database_target_seen = true;
        return Ok(());
    }
    let (order, repeatable) = match (parent.element, local) {
        (Element::DataSource, b"connection-data") => (1, false),
        (Element::DataSource, b"driver-settings") => (2, false),
        (Element::DataSource, b"application-connection-settings") => (3, false),
        (Element::Driver, b"auto-increment") => (1, false),
        (Element::Driver, b"delimiter") => (2, false),
        (Element::Driver, b"character-set") => (3, false),
        (Element::Driver, b"table-settings") => (4, false),
        (Element::TableSettings, b"table-setting") => (1, true),
        (Element::TableSetting, b"delimiter") => (1, false),
        (Element::TableSetting, b"character-set") => (2, false),
        (Element::Application, b"table-filter") => (1, false),
        (Element::Application, b"table-type-filter") => (2, false),
        (Element::Application, b"data-source-settings") => (3, false),
        (Element::TableFilter, b"table-include-filter") => (1, false),
        (Element::TableFilter, b"table-exclude-filter") => (2, false),
        (Element::IncludeFilter | Element::ExcludeFilter, b"table-filter-pattern") => (1, true),
        (Element::TableTypeFilter, b"table-type") => (1, true),
        (Element::DataSourceSettings, b"data-source-setting") => (1, true),
        (Element::DataSourceSetting, b"data-source-setting-value") => (1, true),
        _ => return Ok(()),
    };
    if order < parent.last_child_order || (order == parent.last_child_order && !repeatable) {
        return Err(invalid("ODB settings child order is invalid"));
    }
    if parent.element == Element::DataSource && local == b"connection-data" {
        if parent.connection_data_seen {
            return Err(invalid(
                "ODB data source has multiple connection-data declarations",
            ));
        }
        parent.connection_data_seen = true;
    }
    parent.last_child_order = order;
    Ok(())
}

fn handle_text(stack: &mut [Frame], text: &str, max: usize) -> Result<()> {
    let Some(frame) = stack.last() else {
        return Ok(());
    };
    if matches!(
        frame.element,
        Element::ConnectionResource | Element::FileBasedDatabase | Element::ServerDatabase
    ) {
        return Err(invalid("ODB connection target must have empty content"));
    }
    if matches!(
        frame.element,
        Element::Pattern | Element::TableType | Element::DataSourceSettingValue
    ) {
        return append_text(stack, text, max);
    }
    if frame.element != Element::Other && !is_xml_whitespace(text) {
        return Err(invalid(
            "ODB settings contains text outside a string element",
        ));
    }
    Ok(())
}

fn is_xml_whitespace(value: &str) -> bool {
    value
        .chars()
        .all(|character| matches!(character, ' ' | '\t' | '\n' | '\r'))
}

fn namespace_kind(value: &ResolveResult<'_>) -> NamespaceKind {
    if matches!(value, ResolveResult::Bound(Namespace(uri)) if *uri == DATABASE_NAMESPACE) {
        NamespaceKind::Database
    } else if matches!(value, ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE) {
        NamespaceKind::Office
    } else if matches!(value, ResolveResult::Bound(Namespace(uri)) if *uri == XLINK_NAMESPACE) {
        NamespaceKind::Xlink
    } else {
        NamespaceKind::Other
    }
}

fn parse_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: Element,
    frame: &mut Frame,
    max_attribute_bytes: usize,
) -> Result<()> {
    frame.has_opaque_content = has_opaque_attributes(reader, element, kind)?;
    match kind {
        Element::Login => {
            let user_name = db_attr(reader, element, b"user-name", max_attribute_bytes)?;
            let use_system_user =
                db_attr(reader, element, b"use-system-user", max_attribute_bytes)?
                    .map(|value| parse_bool(&value))
                    .transpose()?;
            if user_name.is_some() && use_system_user.is_some() {
                return Err(invalid(
                    "ODB login cannot declare both user-name and use-system-user",
                ));
            }
            frame.login = Some(LoginSettings {
                user_name,
                use_system_user,
                password_required: db_attr(
                    reader,
                    element,
                    b"is-password-required",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                login_timeout: db_attr(reader, element, b"login-timeout", max_attribute_bytes)?
                    .map(|value| parse_positive_u64(&value, "login-timeout"))
                    .transpose()?,
            });
        },
        Element::Driver => {
            frame.driver = Some(DriverSettings {
                show_deleted: db_attr(reader, element, b"show-deleted", max_attribute_bytes)?
                    .map(|value| parse_bool(&value))
                    .transpose()?,
                system_driver_settings: db_attr(
                    reader,
                    element,
                    b"system-driver-settings",
                    max_attribute_bytes,
                )?,
                base_dn: db_attr(reader, element, b"base-dn", max_attribute_bytes)?,
                first_row_header_line: db_attr(
                    reader,
                    element,
                    b"is-first-row-header-line",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                parameter_name_substitution: db_attr(
                    reader,
                    element,
                    b"parameter-name-substitution",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                auto_increment: None,
                delimiter: None,
                character_set: None,
                table_settings: Vec::new(),
            });
        },
        Element::AutoIncrement => {
            frame.auto_increment = Some(AutoIncrementSettings {
                additional_column_statement: db_attr(
                    reader,
                    element,
                    b"additional-column-statement",
                    max_attribute_bytes,
                )?,
                row_retrieving_statement: db_attr(
                    reader,
                    element,
                    b"row-retrieving-statement",
                    max_attribute_bytes,
                )?,
            });
        },
        Element::Delimiter => {
            frame.delimiter = Some(DelimiterSettings {
                field: db_attr(reader, element, b"field", max_attribute_bytes)?,
                string: db_attr(reader, element, b"string", max_attribute_bytes)?,
                decimal: db_attr(reader, element, b"decimal", max_attribute_bytes)?,
                thousand: db_attr(reader, element, b"thousand", max_attribute_bytes)?,
            });
        },
        Element::CharacterSet => {
            let value = CharacterSetSettings {
                encoding: db_attr(reader, element, b"encoding", max_attribute_bytes)?,
            };
            value.validate()?;
            frame.character_set = Some(value);
        },
        Element::TableSetting => {
            frame.table_setting = Some(TableSetting {
                first_row_header_line: db_attr(
                    reader,
                    element,
                    b"is-first-row-header-line",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                show_deleted: db_attr(reader, element, b"show-deleted", max_attribute_bytes)?
                    .map(|value| parse_bool(&value))
                    .transpose()?,
                delimiter: None,
                character_set: None,
            });
        },
        Element::Application => {
            frame.application = Some(ApplicationConnectionSettings {
                table_name_length_limited: db_attr(
                    reader,
                    element,
                    b"is-table-name-length-limited",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                enable_sql92_check: db_attr(
                    reader,
                    element,
                    b"enable-sql92-check",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                append_table_alias_name: db_attr(
                    reader,
                    element,
                    b"append-table-alias-name",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                ignore_driver_privileges: db_attr(
                    reader,
                    element,
                    b"ignore-driver-privileges",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                boolean_comparison_mode: db_attr(
                    reader,
                    element,
                    b"boolean-comparison-mode",
                    max_attribute_bytes,
                )?
                .map(|value| parse_boolean_comparison_mode(&value))
                .transpose()?,
                use_catalog: db_attr(reader, element, b"use-catalog", max_attribute_bytes)?
                    .map(|value| parse_bool(&value))
                    .transpose()?,
                max_row_count: db_attr(reader, element, b"max-row-count", max_attribute_bytes)?
                    .map(|value| {
                        collapse_xml_whitespace(&value)?
                            .parse::<i64>()
                            .map_err(|_| invalid("invalid ODB max-row-count"))
                    })
                    .transpose()?,
                suppress_version_columns: db_attr(
                    reader,
                    element,
                    b"suppress-version-columns",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                table_filter: None,
                table_type_filter: None,
                data_source_settings: Vec::new(),
            });
        },
        Element::TableFilter => frame.table_filter = Some(TableFilter::new()),
        Element::IncludeFilter | Element::ExcludeFilter => {},
        Element::TableTypeFilter => frame.table_type_filter = Some(TableTypeFilter::new()),
        Element::DataSourceSetting => {
            let name = db_attr(
                reader,
                element,
                b"data-source-setting-name",
                max_attribute_bytes,
            )?
            .ok_or_else(|| invalid("ODB data-source-setting is missing a name"))?;
            let setting_type = db_attr(
                reader,
                element,
                b"data-source-setting-type",
                max_attribute_bytes,
            )?
            .ok_or_else(|| invalid("ODB data-source-setting is missing a type"))
            .and_then(|value| parse_data_source_setting_type(&value))?;
            frame.data_source_setting = Some(DataSourceSetting {
                name,
                setting_type,
                is_list: db_attr(
                    reader,
                    element,
                    b"data-source-setting-is-list",
                    max_attribute_bytes,
                )?
                .map(|value| parse_bool(&value))
                .transpose()?,
                values: Vec::new(),
            });
        },
        Element::ConnectionResource => {
            let link_type = xlink_attr(reader, element, b"type", max_attribute_bytes)?
                .ok_or_else(|| invalid("ODB connection-resource is missing xlink:type"))?;
            if collapse_xml_whitespace(&link_type)? != "simple" {
                return Err(invalid("ODB connection-resource xlink:type must be simple"));
            }
            let _href = xlink_attr(reader, element, b"href", max_attribute_bytes)?
                .ok_or_else(|| invalid("ODB connection-resource is missing xlink:href"))?;
            if let Some(show) = xlink_attr(reader, element, b"show", max_attribute_bytes)?
                && collapse_xml_whitespace(&show)? != "none"
            {
                return Err(invalid("ODB connection-resource xlink:show must be none"));
            }
            if let Some(actuate) = xlink_attr(reader, element, b"actuate", max_attribute_bytes)?
                && collapse_xml_whitespace(&actuate)? != "onRequest"
            {
                return Err(invalid(
                    "ODB connection-resource xlink:actuate must be onRequest",
                ));
            }
        },
        Element::FileBasedDatabase => {
            let link_type = xlink_attr(reader, element, b"type", max_attribute_bytes)?
                .ok_or_else(|| invalid("ODB file-based-database is missing xlink:type"))?;
            if collapse_xml_whitespace(&link_type)? != "simple" {
                return Err(invalid("ODB file-based-database xlink:type must be simple"));
            }
            let _href = xlink_attr(reader, element, b"href", max_attribute_bytes)?
                .ok_or_else(|| invalid("ODB file-based-database is missing xlink:href"))?;
            let _media_type = db_attr(reader, element, b"media-type", max_attribute_bytes)?
                .ok_or_else(|| invalid("ODB file-based-database is missing db:media-type"))?;
        },
        Element::ServerDatabase => {
            let database_type = db_attr(reader, element, b"type", max_attribute_bytes)?
                .ok_or_else(|| invalid("ODB server-database is missing db:type"))?;
            parse_namespaced_token(reader, &database_type, "server-database type")?;
            let hostname = db_attr(reader, element, b"hostname", max_attribute_bytes)?;
            let port = db_attr(reader, element, b"port", max_attribute_bytes)?
                .map(|value| parse_positive_u64(&value, "server-database port"))
                .transpose()?;
            let local_socket = db_attr(reader, element, b"local-socket", max_attribute_bytes)?;
            if hostname.is_some() && local_socket.is_some() {
                return Err(invalid(
                    "ODB server-database cannot declare both hostname and local-socket",
                ));
            }
            if port.is_some() && hostname.is_none() {
                return Err(invalid("ODB server-database port requires a hostname"));
            }
        },
        Element::Document
        | Element::Body
        | Element::Database
        | Element::DataSource
        | Element::ConnectionData
        | Element::DatabaseDescription
        | Element::Pattern
        | Element::TableType
        | Element::TableSettings
        | Element::DataSourceSettings
        | Element::DataSourceSettingValue
        | Element::Other => {},
    }
    Ok(())
}

fn has_opaque_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: Element,
) -> Result<bool> {
    let mut opaque = false;
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(&format!("invalid ODB settings attribute: {error}")))?;
        let raw_name = attribute.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        match namespace {
            ResolveResult::Bound(Namespace(uri)) if uri == DATABASE_NAMESPACE => {
                if !(known_attribute(kind, name.as_ref())
                    || kind == Element::Other && known_catalog_database_attribute(name.as_ref()))
                {
                    return Err(invalid(
                        "ODB settings contains an unknown database attribute",
                    ));
                }
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE => {
                if !known_xlink_attribute(kind, name.as_ref()) {
                    if kind != Element::Other || !known_catalog_xlink_attribute(name.as_ref()) {
                        return Err(invalid("ODB settings contains an unknown xlink attribute"));
                    }
                }
            },
            ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE => {
                if !known_office_attribute(kind, name.as_ref()) {
                    return Err(invalid("ODB settings contains an unknown office attribute"));
                }
            },
            ResolveResult::Unknown(_) | ResolveResult::Unbound if reserved_prefix(raw_name) => {
                return Err(invalid("ODB reserved namespace attribute is not bound"));
            },
            _ => opaque = true,
        }
    }
    Ok(opaque)
}

fn reserved_prefix(raw_name: &[u8]) -> bool {
    raw_name
        .iter()
        .position(|byte| *byte == b':')
        .is_some_and(|index| matches!(&raw_name[..index], b"db" | b"office" | b"xlink"))
}

fn known_attribute(kind: Element, name: &[u8]) -> bool {
    match kind {
        Element::Login => matches!(
            name,
            b"user-name" | b"use-system-user" | b"is-password-required" | b"login-timeout"
        ),
        Element::Driver => matches!(
            name,
            b"show-deleted"
                | b"system-driver-settings"
                | b"base-dn"
                | b"is-first-row-header-line"
                | b"parameter-name-substitution"
        ),
        Element::AutoIncrement => {
            matches!(
                name,
                b"additional-column-statement" | b"row-retrieving-statement"
            )
        },
        Element::Delimiter => {
            matches!(name, b"field" | b"string" | b"decimal" | b"thousand")
        },
        Element::CharacterSet => name == b"encoding",
        Element::FileBasedDatabase => matches!(name, b"media-type" | b"extension"),
        Element::ServerDatabase => {
            matches!(
                name,
                b"type" | b"hostname" | b"port" | b"local-socket" | b"database-name"
            )
        },
        Element::TableSetting => matches!(name, b"is-first-row-header-line" | b"show-deleted"),
        Element::Application => matches!(
            name,
            b"is-table-name-length-limited"
                | b"enable-sql92-check"
                | b"append-table-alias-name"
                | b"ignore-driver-privileges"
                | b"boolean-comparison-mode"
                | b"use-catalog"
                | b"max-row-count"
                | b"suppress-version-columns"
        ),
        Element::DataSourceSetting => matches!(
            name,
            b"data-source-setting-name"
                | b"data-source-setting-type"
                | b"data-source-setting-is-list"
        ),
        _ => false,
    }
}

fn known_catalog_database_attribute(name: &[u8]) -> bool {
    matches!(
        name,
        b"additional-column-statement"
            | b"append-table-alias-name"
            | b"apply-command"
            | b"as-template"
            | b"base-dn"
            | b"boolean-comparison-mode"
            | b"catalog-name"
            | b"command"
            | b"data-source-setting-is-list"
            | b"data-source-setting-name"
            | b"data-source-setting-type"
            | b"data-type"
            | b"database-name"
            | b"decimal"
            | b"default-cell-style-name"
            | b"default-value"
            | b"default-row-style-name"
            | b"delete-rule"
            | b"description"
            | b"enable-sql92-check"
            | b"encoding"
            | b"escape-processing"
            | b"extension"
            | b"field"
            | b"hostname"
            | b"ignore-driver-privileges"
            | b"is-ascending"
            | b"is-autoincrement"
            | b"is-clustered"
            | b"is-empty-allowed"
            | b"is-first-row-header-line"
            | b"is-nullable"
            | b"is-password-required"
            | b"is-table-name-length-limited"
            | b"is-unique"
            | b"local-socket"
            | b"login-timeout"
            | b"max-row-count"
            | b"media-type"
            | b"name"
            | b"parameter-name-substitution"
            | b"port"
            | b"precision"
            | b"referenced-table-name"
            | b"related-column-name"
            | b"row-retrieving-statement"
            | b"scale"
            | b"schema-name"
            | b"show-deleted"
            | b"string"
            | b"style-name"
            | b"suppress-version-columns"
            | b"system-driver-settings"
            | b"thousand"
            | b"title"
            | b"type"
            | b"type-name"
            | b"update-rule"
            | b"use-catalog"
            | b"use-system-user"
            | b"user-name"
            | b"visible"
    )
}

fn known_catalog_xlink_attribute(name: &[u8]) -> bool {
    matches!(name, b"actuate" | b"href" | b"show" | b"title" | b"type")
}

fn known_office_attribute(kind: Element, name: &[u8]) -> bool {
    kind == Element::Document && name == b"version"
}

fn known_xlink_attribute(kind: Element, name: &[u8]) -> bool {
    matches!(
        (kind, name),
        (
            Element::ConnectionResource,
            b"type" | b"href" | b"show" | b"actuate"
        ) | (Element::FileBasedDatabase, b"type" | b"href")
    )
}

fn finalize(
    frame: Frame,
    stack: &mut [Frame],
    settings: &mut DatabaseSettings,
    setting_count: &mut usize,
    value_count: &mut usize,
    table_setting_count: &mut usize,
    pattern_count: &mut usize,
    max_attribute_bytes: usize,
    max_settings: usize,
    max_setting_values: usize,
    max_table_settings: usize,
    max_filter_patterns: usize,
) -> Result<()> {
    match frame.element {
        Element::DataSource if !frame.connection_data_seen => {
            return Err(invalid(
                "ODB data source requires exactly one connection-data declaration",
            ));
        },
        Element::ConnectionData if !frame.connection_target_seen => {
            return Err(invalid(
                "ODB connection-data requires exactly one connection target",
            ));
        },
        Element::DatabaseDescription if !frame.database_target_seen => {
            return Err(invalid(
                "ODB database-description requires exactly one file or server database",
            ));
        },
        Element::ConnectionResource if frame.has_child => {
            return Err(invalid("ODB connection-resource must have empty content"));
        },
        Element::FileBasedDatabase | Element::ServerDatabase if frame.has_child => {
            return Err(invalid("ODB database target must have empty content"));
        },
        _ => {},
    }
    if frame.element == Element::Pattern {
        *pattern_count = pattern_count
            .checked_add(1)
            .ok_or_else(|| invalid("ODB filter pattern count overflow"))?;
        if *pattern_count > max_filter_patterns.min(MAX_FILTER_PATTERNS) {
            return Err(invalid("ODB filter pattern count exceeds the limit"));
        }
        if frame.text.len() > max_attribute_bytes {
            return Err(invalid("ODB filter pattern exceeds the byte limit"));
        }
        let parent_index = stack
            .len()
            .checked_sub(1)
            .ok_or_else(|| invalid("ODB table pattern has no filter owner"))?;
        let include = stack[parent_index].element == Element::IncludeFilter;
        let filter_index = stack[..parent_index]
            .iter()
            .rposition(|value| value.element == Element::TableFilter)
            .ok_or_else(|| invalid("ODB table pattern has no table-filter owner"))?;
        let filter = stack[filter_index]
            .table_filter
            .as_mut()
            .ok_or_else(|| invalid("ODB table pattern filter state is missing"))?;
        if include {
            filter
                .include
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODB table include-filter patterns",
                    source,
                })?;
            filter.include.push(frame.text);
        } else if stack[parent_index].element == Element::ExcludeFilter {
            filter
                .exclude
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODB table exclude-filter patterns",
                    source,
                })?;
            filter.exclude.push(frame.text);
        } else {
            return Err(invalid("ODB table pattern has an invalid filter owner"));
        }
        stack[parent_index].child_count = stack[parent_index]
            .child_count
            .checked_add(1)
            .ok_or_else(|| invalid("ODB filter pattern count overflow"))?;
        return Ok(());
    }
    let Some(parent) = stack.last_mut() else {
        match frame.element {
            Element::DataSource => {},
            Element::Other => {},
            _ => return Ok(()),
        }
        return Ok(());
    };
    match frame.element {
        Element::Login => {
            if settings
                .login
                .replace(
                    frame
                        .login
                        .ok_or_else(|| invalid("ODB login state is missing"))?,
                )
                .is_some()
            {
                return Err(invalid("ODB data source has multiple login declarations"));
            }
        },
        Element::Driver => {
            if settings
                .driver
                .replace(
                    frame
                        .driver
                        .ok_or_else(|| invalid("ODB driver state is missing"))?,
                )
                .is_some()
            {
                return Err(invalid(
                    "ODB data source has multiple driver-settings declarations",
                ));
            }
        },
        Element::Application => {
            if settings
                .application_connection
                .replace(
                    frame
                        .application
                        .ok_or_else(|| invalid("ODB application settings state is missing"))?,
                )
                .is_some()
            {
                return Err(invalid(
                    "ODB data source has multiple application settings declarations",
                ));
            }
        },
        Element::AutoIncrement => {
            let value = frame
                .auto_increment
                .ok_or_else(|| invalid("ODB auto-increment state is missing"))?;
            let driver = parent
                .driver
                .as_mut()
                .ok_or_else(|| invalid("ODB auto-increment has no driver owner"))?;
            if driver.auto_increment.replace(value).is_some() {
                return Err(invalid(
                    "ODB driver has multiple auto-increment declarations",
                ));
            }
        },
        Element::Delimiter => {
            let value = frame
                .delimiter
                .ok_or_else(|| invalid("ODB delimiter state is missing"))?;
            if parent.element == Element::Driver {
                let driver = parent
                    .driver
                    .as_mut()
                    .ok_or_else(|| invalid("ODB driver state is missing"))?;
                if driver.delimiter.replace(value).is_some() {
                    return Err(invalid("ODB driver has multiple delimiter declarations"));
                }
            } else if parent.element == Element::TableSetting {
                let table_setting = parent
                    .table_setting
                    .as_mut()
                    .ok_or_else(|| invalid("ODB table setting state is missing"))?;
                if table_setting.delimiter.replace(value).is_some() {
                    return Err(invalid(
                        "ODB table setting has multiple delimiter declarations",
                    ));
                }
            }
        },
        Element::CharacterSet => {
            let value = frame
                .character_set
                .ok_or_else(|| invalid("ODB character-set state is missing"))?;
            if parent.element == Element::Driver {
                let driver = parent
                    .driver
                    .as_mut()
                    .ok_or_else(|| invalid("ODB driver state is missing"))?;
                if driver.character_set.replace(value).is_some() {
                    return Err(invalid(
                        "ODB driver has multiple character-set declarations",
                    ));
                }
            } else if parent.element == Element::TableSetting {
                let table_setting = parent
                    .table_setting
                    .as_mut()
                    .ok_or_else(|| invalid("ODB table setting state is missing"))?;
                if table_setting.character_set.replace(value).is_some() {
                    return Err(invalid(
                        "ODB table setting has multiple character-set declarations",
                    ));
                }
            }
        },
        Element::TableSetting => {
            *table_setting_count = table_setting_count
                .checked_add(1)
                .ok_or_else(|| invalid("ODB table-setting count overflow"))?;
            if *table_setting_count > max_table_settings.min(MAX_TABLE_SETTINGS) {
                return Err(invalid("ODB table-setting count exceeds the limit"));
            }
            let value = frame
                .table_setting
                .ok_or_else(|| invalid("ODB table setting state is missing"))?;
            let driver = stack
                .iter_mut()
                .rev()
                .find_map(|frame| frame.driver.as_mut())
                .ok_or_else(|| invalid("ODB table setting has no driver owner"))?;
            driver
                .table_settings
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODB driver table-settings",
                    source,
                })?;
            driver.table_settings.push(value);
        },
        Element::TableFilter => {
            let value = frame
                .table_filter
                .ok_or_else(|| invalid("ODB table filter state is missing"))?;
            let application = parent
                .application
                .as_mut()
                .ok_or_else(|| invalid("ODB table filter has no application owner"))?;
            if application.table_filter.replace(value).is_some() {
                return Err(invalid(
                    "ODB application settings has multiple table-filter declarations",
                ));
            }
        },
        Element::IncludeFilter | Element::ExcludeFilter => {
            if frame.child_count == 0 {
                return Err(invalid(
                    "ODB table include/exclude filter requires a pattern",
                ));
            }
        },
        Element::Pattern => {},
        Element::TableTypeFilter => {
            let value = frame
                .table_type_filter
                .ok_or_else(|| invalid("ODB table-type-filter state is missing"))?;
            let application = parent
                .application
                .as_mut()
                .ok_or_else(|| invalid("ODB table-type-filter has no application owner"))?;
            if application.table_type_filter.replace(value).is_some() {
                return Err(invalid(
                    "ODB application settings has multiple table-type-filter declarations",
                ));
            }
        },
        Element::TableType => {
            *pattern_count = pattern_count
                .checked_add(1)
                .ok_or_else(|| invalid("ODB table type count overflow"))?;
            if *pattern_count > max_filter_patterns.min(MAX_FILTER_PATTERNS) {
                return Err(invalid("ODB table type count exceeds the limit"));
            }
            if frame.text.len() > max_attribute_bytes {
                return Err(invalid("ODB table type exceeds the byte limit"));
            }
            parent
                .table_type_filter
                .as_mut()
                .ok_or_else(|| invalid("ODB table type has no filter owner"))?
                .table_types
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODB table-type filters",
                    source,
                })?;
            parent
                .table_type_filter
                .as_mut()
                .ok_or_else(|| invalid("ODB table type has no filter owner"))?
                .table_types
                .push(frame.text);
            parent
                .child_count
                .checked_add(1)
                .ok_or_else(|| invalid("ODB table type count overflow"))
                .map(|count| parent.child_count = count)?;
        },
        Element::DataSourceSettings => {
            if frame.child_count == 0 {
                return Err(invalid(
                    "ODB data-source-settings requires a data-source setting",
                ));
            }
            let application = parent
                .application
                .as_mut()
                .ok_or_else(|| invalid("ODB data-source-settings has no application owner"))?;
            if parent.data_source_settings_container_seen {
                return Err(invalid(
                    "ODB application settings has multiple data-source-settings declarations",
                ));
            }
            parent.data_source_settings_container_seen = true;
            application.data_source_settings = frame.data_source_settings;
        },
        Element::DataSourceSetting => {
            *setting_count = setting_count
                .checked_add(1)
                .ok_or_else(|| invalid("ODB data-source setting count overflow"))?;
            if *setting_count > max_settings.min(MAX_SETTINGS) {
                return Err(invalid("ODB data-source setting count exceeds the limit"));
            }
            let value = frame
                .data_source_setting
                .ok_or_else(|| invalid("ODB data-source setting state is missing"))?;
            if value.values.is_empty() {
                return Err(invalid(
                    "ODB data-source setting requires at least one value",
                ));
            }
            if parent.element != Element::DataSourceSettings {
                return Err(invalid("ODB data-source setting has no settings owner"));
            }
            parent.child_count = parent
                .child_count
                .checked_add(1)
                .ok_or_else(|| invalid("ODB data-source setting count overflow"))?;
            parent
                .data_source_settings
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODB data-source settings",
                    source,
                })?;
            parent.data_source_settings.push(value);
        },
        Element::DataSourceSettingValue => {
            *value_count = value_count
                .checked_add(1)
                .ok_or_else(|| invalid("ODB data-source setting value count overflow"))?;
            if *value_count > max_setting_values.min(MAX_SETTING_VALUES) {
                return Err(invalid(
                    "ODB data-source setting value count exceeds the limit",
                ));
            }
            if frame.text.len() > max_attribute_bytes {
                return Err(invalid(
                    "ODB data-source setting value exceeds the byte limit",
                ));
            }
            let setting = parent
                .data_source_setting
                .as_mut()
                .ok_or_else(|| invalid("ODB setting value has no setting owner"))?;
            setting
                .values
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "ODB data-source setting values",
                    source,
                })?;
            setting.values.push(frame.text);
        },
        Element::Document
        | Element::Body
        | Element::Database
        | Element::DataSource
        | Element::ConnectionData
        | Element::DatabaseDescription
        | Element::FileBasedDatabase
        | Element::ServerDatabase
        | Element::ConnectionResource
        | Element::TableSettings
        | Element::Other => {},
    }
    Ok(())
}

fn append_text(stack: &mut [Frame], text: &str, max: usize) -> Result<()> {
    let frame = stack
        .last_mut()
        .ok_or_else(|| invalid("ODB settings text has no element owner"))?;
    let next = frame
        .text
        .len()
        .checked_add(text.len())
        .ok_or_else(|| invalid("ODB settings text length overflow"))?;
    if next > max.min(MAX_TEXT_BYTES) {
        return Err(invalid("ODB settings text exceeds the byte limit"));
    }
    frame
        .text
        .try_reserve(text.len())
        .map_err(|source| Error::Allocation {
            resource: "ODB settings text",
            source,
        })?;
    frame.text.push_str(text);
    Ok(())
}

fn decode_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if reference.as_ref().len() > MAX_REFERENCE_BYTES {
        return Err(invalid(
            "ODB settings entity reference exceeds the byte limit",
        ));
    }
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        invalid(&format!(
            "invalid ODB settings character reference: {error}"
        ))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| invalid(&format!("invalid ODB settings entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_owned()),
        "lt" => Ok("<".to_owned()),
        "gt" => Ok(">".to_owned()),
        "quot" => Ok("\"".to_owned()),
        "apos" => Ok("'".to_owned()),
        _ => Err(invalid("unsupported ODB settings entity")),
    }
}

fn db_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
    max: usize,
) -> Result<Option<String>> {
    let mut found = None;
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(&format!("invalid ODB settings attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == DATABASE_NAMESPACE)
            && name.as_ref() == local
        {
            if attribute.value.len() > max {
                return Err(invalid("ODB settings attribute exceeds the byte limit"));
            }
            let value = decode_attribute_value(reader, &attribute)?;
            if value.len() > max {
                return Err(invalid(
                    "ODB decoded settings attribute exceeds the byte limit",
                ));
            }
            if found.replace(value).is_some() {
                return Err(invalid("duplicate ODB settings attribute"));
            }
        }
    }
    Ok(found)
}

fn xlink_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
    max: usize,
) -> Result<Option<String>> {
    let mut found = None;
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(&format!("invalid ODB settings attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE)
            && name.as_ref() == local
        {
            if attribute.value.len() > max {
                return Err(invalid("ODB settings attribute exceeds the byte limit"));
            }
            let value = decode_attribute_value(reader, &attribute)?;
            if value.len() > max {
                return Err(invalid(
                    "ODB decoded settings attribute exceeds the byte limit",
                ));
            }
            if found.replace(value).is_some() {
                return Err(invalid("duplicate ODB settings attribute"));
            }
        }
    }
    Ok(found)
}

fn decode_attribute_value(reader: &NsReader<&[u8]>, attribute: &Attribute<'_>) -> Result<String> {
    let decoded = reader
        .decoder()
        .decode(attribute.value.as_ref())
        .map_err(|error| invalid(&format!("invalid ODB settings attribute value: {error}")))?;
    quick_xml::escape::unescape(decoded.as_ref())
        .map(|value| value.into_owned())
        .map_err(|error| invalid(&format!("invalid ODB settings attribute value: {error}")))
}

fn parse_bool(value: &str) -> Result<bool> {
    let value = collapse_xml_whitespace(value)?;
    match value.as_str() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("invalid ODB settings boolean")),
    }
}

fn parse_positive_u64(value: &str, kind: &str) -> Result<u64> {
    let value = collapse_xml_whitespace(value)?
        .parse::<u64>()
        .map_err(|_| invalid(&format!("invalid ODB {kind}")))?;
    if value == 0 {
        return Err(invalid(&format!("ODB {kind} must be positive")));
    }
    Ok(value)
}

fn parse_namespaced_token(reader: &NsReader<&[u8]>, value: &str, kind: &str) -> Result<()> {
    let value = collapse_xml_whitespace(value)?;
    let mut parts = value.split(':');
    let prefix = parts.next().unwrap_or_default();
    let local = parts.next().unwrap_or_default();
    if !is_ncname(prefix) || !is_ncname(local) || parts.next().is_some() {
        return Err(invalid(&format!("invalid ODB {kind}")));
    }
    match reader
        .resolver()
        .resolve_prefix(QName(value.as_bytes()).prefix(), false)
    {
        ResolveResult::Bound(_) => Ok(()),
        ResolveResult::Unknown(_) | ResolveResult::Unbound => {
            Err(invalid(&format!("ODB {kind} prefix is not bound")))
        },
    }
}

fn is_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !(first == '_' || first.is_alphabetic()) {
        return false;
    }
    chars.all(|character| {
        character == '_'
            || character == '-'
            || character == '.'
            || character.is_alphanumeric()
            || character == '\u{b7}'
            || matches!(character as u32, 0x0300..=0x036f | 0x203f..=0x2040)
    })
}

fn parse_boolean_comparison_mode(value: &str) -> Result<BooleanComparisonMode> {
    let value = collapse_xml_whitespace(value)?;
    match value.as_str() {
        "equal-integer" => Ok(BooleanComparisonMode::EqualInteger),
        "is-boolean" => Ok(BooleanComparisonMode::IsBoolean),
        "equal-boolean" => Ok(BooleanComparisonMode::EqualBoolean),
        "equal-use-only-zero" => Ok(BooleanComparisonMode::EqualUseOnlyZero),
        _ => Err(invalid("invalid ODB boolean-comparison-mode")),
    }
}

fn parse_data_source_setting_type(value: &str) -> Result<DataSourceSettingType> {
    let value = collapse_xml_whitespace(value)?;
    match value.as_str() {
        "boolean" => Ok(DataSourceSettingType::Boolean),
        "short" => Ok(DataSourceSettingType::Short),
        "int" => Ok(DataSourceSettingType::Int),
        "long" => Ok(DataSourceSettingType::Long),
        "double" => Ok(DataSourceSettingType::Double),
        "string" => Ok(DataSourceSettingType::String),
        _ => Err(invalid("invalid ODB data-source-setting type")),
    }
}

fn collapse_xml_whitespace(value: &str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "ODB settings token",
            source,
        })?;
    let mut pending_space = false;
    for character in value.chars() {
        if matches!(character, ' ' | '\t' | '\n' | '\r') {
            if !output.is_empty() {
                pending_space = true;
            }
        } else {
            if pending_space {
                output.push(' ');
                pending_space = false;
            }
            output.push(character);
        }
    }
    Ok(output)
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}
