//! Typed inert driver and application connection settings for ODF databases.

use super::connection_data::connection_data_from_root;
use super::document::{
    DatabaseContent, DatabaseDocument, DatabaseElement, DatabaseElementKind,
    parse_database_content, validate_database_root,
};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const DB: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_VALUE: usize = 1024 * 1024;
const MAX_AGGREGATE: usize = 16 * 1024 * 1024;
const MAX_ITEMS: usize = 4096;
const MAX_VALUES: usize = 65_536;
const MAX_INTEGER_DIGITS: usize = 4096;

/// Canonical arbitrary-width XML Schema `integer` metadata.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfDatabaseInteger(String);

impl OdfDatabaseInteger {
    pub fn new(value: &str) -> Result<Self> {
        let value = value.trim_matches(|ch| matches!(ch, ' ' | '\t' | '\r' | '\n'));
        let (negative, digits) = match value.as_bytes().first() {
            Some(b'+') => (false, &value[1..]),
            Some(b'-') => (true, &value[1..]),
            _ => (false, value),
        };
        if digits.is_empty()
            || digits.len() > MAX_INTEGER_DIGITS
            || !digits.bytes().all(|byte| byte.is_ascii_digit())
        {
            return Err(Error::InvalidFormat(
                "database integer has an invalid lexical value".to_string(),
            ));
        }
        let canonical = digits.trim_start_matches('0');
        if canonical.is_empty() {
            return Ok(Self("0".to_string()));
        }
        Ok(Self(if negative {
            format!("-{canonical}")
        } else {
            canonical.to_string()
        }))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseAutoIncrement {
    pub additional_column_statement: Option<String>,
    pub row_retrieving_statement: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseDelimiter {
    pub field: Option<String>,
    pub string: Option<String>,
    pub decimal: Option<String>,
    pub thousand: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseCharacterSet {
    pub encoding: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseTableSetting {
    pub is_first_row_header_line: Option<bool>,
    pub show_deleted: Option<bool>,
    pub delimiter: Option<OdfDatabaseDelimiter>,
    pub character_set: Option<OdfDatabaseCharacterSet>,
}

/// Complete `db:driver-settings` metadata.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseDriverSettings {
    pub show_deleted: Option<bool>,
    pub system_driver_settings: Option<String>,
    pub base_dn: Option<String>,
    pub is_first_row_header_line: Option<bool>,
    pub parameter_name_substitution: Option<bool>,
    pub auto_increment: Option<OdfDatabaseAutoIncrement>,
    pub delimiter: Option<OdfDatabaseDelimiter>,
    pub character_set: Option<OdfDatabaseCharacterSet>,
    /// `Some(empty)` preserves a present empty `db:table-settings` container.
    pub table_settings: Option<Vec<OdfDatabaseTableSetting>>,
}

impl OdfDatabaseDriverSettings {
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_driver(self)?;
        let mut xml = format!("<db:driver-settings xmlns:db=\"{DB}\"");
        attr_bool(&mut xml, "show-deleted", self.show_deleted);
        attr_opt(
            &mut xml,
            "system-driver-settings",
            self.system_driver_settings.as_deref(),
        );
        attr_opt(&mut xml, "base-dn", self.base_dn.as_deref());
        attr_bool(
            &mut xml,
            "is-first-row-header-line",
            self.is_first_row_header_line,
        );
        attr_bool(
            &mut xml,
            "parameter-name-substitution",
            self.parameter_name_substitution,
        );
        if self.auto_increment.is_none()
            && self.delimiter.is_none()
            && self.character_set.is_none()
            && self.table_settings.is_none()
        {
            xml.push_str("/>");
            return Ok(xml);
        }
        xml.push('>');
        if let Some(value) = &self.auto_increment {
            xml.push_str("<db:auto-increment");
            attr_opt(
                &mut xml,
                "additional-column-statement",
                value.additional_column_statement.as_deref(),
            );
            attr_opt(
                &mut xml,
                "row-retrieving-statement",
                value.row_retrieving_statement.as_deref(),
            );
            xml.push_str("/>");
        }
        if let Some(value) = &self.delimiter {
            push_delimiter(&mut xml, value);
        }
        if let Some(value) = &self.character_set {
            push_character_set(&mut xml, value);
        }
        if let Some(settings) = &self.table_settings {
            xml.push_str("<db:table-settings>");
            for setting in settings {
                xml.push_str("<db:table-setting");
                attr_bool(
                    &mut xml,
                    "is-first-row-header-line",
                    setting.is_first_row_header_line,
                );
                attr_bool(&mut xml, "show-deleted", setting.show_deleted);
                if setting.delimiter.is_none() && setting.character_set.is_none() {
                    xml.push_str("/>");
                } else {
                    xml.push('>');
                    if let Some(value) = &setting.delimiter {
                        push_delimiter(&mut xml, value);
                    }
                    if let Some(value) = &setting.character_set {
                        push_character_set(&mut xml, value);
                    }
                    xml.push_str("</db:table-setting>");
                }
            }
            xml.push_str("</db:table-settings>");
        }
        xml.push_str("</db:driver-settings>");
        Ok(xml)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfDatabaseBooleanComparisonMode {
    EqualInteger,
    IsBoolean,
    EqualBoolean,
    EqualUseOnlyZero,
}

impl OdfDatabaseBooleanComparisonMode {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "equal-integer" => Ok(Self::EqualInteger),
            "is-boolean" => Ok(Self::IsBoolean),
            "equal-boolean" => Ok(Self::EqualBoolean),
            "equal-use-only-zero" => Ok(Self::EqualUseOnlyZero),
            _ => Err(Error::InvalidFormat(
                "invalid database boolean-comparison-mode".to_string(),
            )),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::EqualInteger => "equal-integer",
            Self::IsBoolean => "is-boolean",
            Self::EqualBoolean => "equal-boolean",
            Self::EqualUseOnlyZero => "equal-use-only-zero",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseTableFilter {
    pub include: Option<Vec<String>>,
    pub exclude: Option<Vec<String>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfDatabaseSettingType {
    Boolean,
    Short,
    Int,
    Long,
    Double,
    String,
}

impl OdfDatabaseSettingType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "boolean" => Ok(Self::Boolean),
            "short" => Ok(Self::Short),
            "int" => Ok(Self::Int),
            "long" => Ok(Self::Long),
            "double" => Ok(Self::Double),
            "string" => Ok(Self::String),
            _ => Err(Error::InvalidFormat(
                "invalid data-source-setting type".to_string(),
            )),
        }
    }

    fn as_str(self) -> &'static str {
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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseDataSourceSetting {
    pub is_list: Option<bool>,
    pub name: String,
    pub value_type: OdfDatabaseSettingType,
    pub values: Vec<String>,
}

/// Complete `db:application-connection-settings` metadata.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseApplicationConnectionSettings {
    pub is_table_name_length_limited: Option<bool>,
    pub enable_sql92_check: Option<bool>,
    pub append_table_alias_name: Option<bool>,
    pub ignore_driver_privileges: Option<bool>,
    pub boolean_comparison_mode: Option<OdfDatabaseBooleanComparisonMode>,
    pub use_catalog: Option<bool>,
    pub max_row_count: Option<OdfDatabaseInteger>,
    pub suppress_version_columns: Option<bool>,
    pub table_filter: Option<OdfDatabaseTableFilter>,
    /// `Some(empty)` preserves an empty `db:table-type-filter`.
    pub table_types: Option<Vec<String>>,
    /// The RNG requires one or more settings whenever this container is present.
    pub data_source_settings: Option<Vec<OdfDatabaseDataSourceSetting>>,
}

impl OdfDatabaseApplicationConnectionSettings {
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_application(self)?;
        let mut xml = format!("<db:application-connection-settings xmlns:db=\"{DB}\"");
        attr_bool(
            &mut xml,
            "is-table-name-length-limited",
            self.is_table_name_length_limited,
        );
        attr_bool(&mut xml, "enable-sql92-check", self.enable_sql92_check);
        attr_bool(
            &mut xml,
            "append-table-alias-name",
            self.append_table_alias_name,
        );
        attr_bool(
            &mut xml,
            "ignore-driver-privileges",
            self.ignore_driver_privileges,
        );
        if let Some(value) = self.boolean_comparison_mode {
            attr(&mut xml, "boolean-comparison-mode", value.as_str());
        }
        attr_bool(&mut xml, "use-catalog", self.use_catalog);
        if let Some(value) = &self.max_row_count {
            attr(&mut xml, "max-row-count", value.as_str());
        }
        attr_bool(
            &mut xml,
            "suppress-version-columns",
            self.suppress_version_columns,
        );
        if self.table_filter.is_none()
            && self.table_types.is_none()
            && self.data_source_settings.is_none()
        {
            xml.push_str("/>");
            return Ok(xml);
        }
        xml.push('>');
        if let Some(filter) = &self.table_filter {
            xml.push_str("<db:table-filter>");
            push_patterns(&mut xml, "table-include-filter", filter.include.as_deref());
            push_patterns(&mut xml, "table-exclude-filter", filter.exclude.as_deref());
            xml.push_str("</db:table-filter>");
        }
        if let Some(types) = &self.table_types {
            xml.push_str("<db:table-type-filter>");
            for value in types {
                xml.push_str("<db:table-type>");
                text(&mut xml, value);
                xml.push_str("</db:table-type>");
            }
            xml.push_str("</db:table-type-filter>");
        }
        if let Some(settings) = &self.data_source_settings {
            xml.push_str("<db:data-source-settings>");
            for setting in settings {
                xml.push_str("<db:data-source-setting");
                attr_bool(&mut xml, "data-source-setting-is-list", setting.is_list);
                attr(&mut xml, "data-source-setting-name", &setting.name);
                attr(
                    &mut xml,
                    "data-source-setting-type",
                    setting.value_type.as_str(),
                );
                xml.push('>');
                for value in &setting.values {
                    xml.push_str("<db:data-source-setting-value>");
                    text(&mut xml, value);
                    xml.push_str("</db:data-source-setting-value>");
                }
                xml.push_str("</db:data-source-setting>");
            }
            xml.push_str("</db:data-source-settings>");
        }
        xml.push_str("</db:application-connection-settings>");
        Ok(xml)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseTrailingSettings {
    pub driver: Option<OdfDatabaseDriverSettings>,
    pub application: Option<OdfDatabaseApplicationConnectionSettings>,
}

impl DatabaseDocument {
    pub fn driver_settings(&self) -> Result<Option<OdfDatabaseDriverSettings>> {
        Ok(trailing_settings_from_root(self.database())?.driver)
    }

    pub fn application_connection_settings(
        &self,
    ) -> Result<Option<OdfDatabaseApplicationConnectionSettings>> {
        Ok(trailing_settings_from_root(self.database())?.application)
    }

    pub fn trailing_connection_settings(&self) -> Result<OdfDatabaseTrailingSettings> {
        trailing_settings_from_root(self.database())
    }
}

pub fn parse_database_trailing_settings_xml(xml: &str) -> Result<OdfDatabaseTrailingSettings> {
    preflight(xml)?;
    let root = parse_database_content(xml)?;
    validate_database_root(&root)?;
    trailing_settings_from_root(&root)
}

pub fn set_database_driver_settings_xml(
    xml: &str,
    value: Option<&OdfDatabaseDriverSettings>,
) -> Result<String> {
    mutate_settings(
        xml,
        SettingsTarget::Driver,
        value.map(|item| item.to_xml_fragment()).transpose()?,
    )
}

pub fn set_database_application_connection_settings_xml(
    xml: &str,
    value: Option<&OdfDatabaseApplicationConnectionSettings>,
) -> Result<String> {
    mutate_settings(
        xml,
        SettingsTarget::Application,
        value.map(|item| item.to_xml_fragment()).transpose()?,
    )
}

fn trailing_settings_from_root(root: &DatabaseElement) -> Result<OdfDatabaseTrailingSettings> {
    connection_data_from_root(root)?;
    let source = root
        .children()
        .find(|child| child.kind() == DatabaseElementKind::DataSource)
        .ok_or_else(|| Error::InvalidFormat("database has no data source".to_string()))?;
    let children = children(source)?;
    let driver = children
        .iter()
        .find(|child| child.kind() == DatabaseElementKind::DriverSettings)
        .map(|element| parse_driver(element))
        .transpose()?;
    let application = children
        .iter()
        .find(|child| child.kind() == DatabaseElementKind::ApplicationConnectionSettings)
        .map(|element| parse_application(element))
        .transpose()?;
    Ok(OdfDatabaseTrailingSettings {
        driver,
        application,
    })
}

fn parse_driver(element: &DatabaseElement) -> Result<OdfDatabaseDriverSettings> {
    allow_attrs(
        element,
        &[
            "show-deleted",
            "system-driver-settings",
            "base-dn",
            "is-first-row-header-line",
            "parameter-name-substitution",
        ],
    )?;
    let direct = children(element)?;
    let mut phase = 0u8;
    let mut auto_increment = None;
    let mut delimiter = None;
    let mut character_set = None;
    let mut table_settings = None;
    for child in direct {
        match child.kind() {
            DatabaseElementKind::AutoIncrement if phase == 0 => {
                phase = 1;
                auto_increment = Some(parse_auto_increment(child)?);
            },
            DatabaseElementKind::Delimiter if phase <= 1 => {
                phase = 2;
                if delimiter.is_some() {
                    return duplicate("db:delimiter");
                }
                delimiter = Some(parse_delimiter(child)?);
            },
            DatabaseElementKind::CharacterSet | DatabaseElementKind::FontCharset if phase <= 2 => {
                phase = 3;
                if character_set.is_some() {
                    return duplicate("db:character-set");
                }
                character_set = Some(parse_character_set(child)?);
            },
            DatabaseElementKind::TableSettings if phase <= 3 => {
                phase = 4;
                if table_settings.is_some() {
                    return duplicate("db:table-settings");
                }
                table_settings = Some(parse_table_settings(child)?);
            },
            _ => return order_error("db:driver-settings"),
        }
    }
    let value = OdfDatabaseDriverSettings {
        show_deleted: bool_attr(element, "show-deleted")?,
        system_driver_settings: opt_attr(element, "system-driver-settings"),
        base_dn: opt_attr(element, "base-dn"),
        is_first_row_header_line: bool_attr(element, "is-first-row-header-line")?,
        parameter_name_substitution: bool_attr(element, "parameter-name-substitution")?,
        auto_increment,
        delimiter,
        character_set,
        table_settings,
    };
    validate_driver(&value)?;
    Ok(value)
}

fn parse_auto_increment(element: &DatabaseElement) -> Result<OdfDatabaseAutoIncrement> {
    empty(element)?;
    allow_attrs(
        element,
        &["additional-column-statement", "row-retrieving-statement"],
    )?;
    Ok(OdfDatabaseAutoIncrement {
        additional_column_statement: opt_attr(element, "additional-column-statement"),
        row_retrieving_statement: opt_attr(element, "row-retrieving-statement"),
    })
}

fn parse_delimiter(element: &DatabaseElement) -> Result<OdfDatabaseDelimiter> {
    empty(element)?;
    allow_attrs(element, &["field", "string", "decimal", "thousand"])?;
    Ok(OdfDatabaseDelimiter {
        field: opt_attr(element, "field"),
        string: opt_attr(element, "string"),
        decimal: opt_attr(element, "decimal"),
        thousand: opt_attr(element, "thousand"),
    })
}

fn parse_character_set(element: &DatabaseElement) -> Result<OdfDatabaseCharacterSet> {
    empty(element)?;
    allow_attrs(element, &["encoding"])?;
    Ok(OdfDatabaseCharacterSet {
        encoding: opt_attr(element, "encoding"),
    })
}

fn parse_table_settings(element: &DatabaseElement) -> Result<Vec<OdfDatabaseTableSetting>> {
    no_attrs(element)?;
    let direct = children(element)?;
    bounded(direct.len(), MAX_ITEMS, "table settings")?;
    direct
        .into_iter()
        .map(|child| {
            if child.kind() != DatabaseElementKind::TableSetting {
                return Err(Error::InvalidFormat(
                    "db:table-settings accepts only db:table-setting".to_string(),
                ));
            }
            allow_attrs(child, &["is-first-row-header-line", "show-deleted"])?;
            let nested = children(child)?;
            let mut delimiter = None;
            let mut character_set = None;
            for nested_child in nested {
                match nested_child.kind() {
                    DatabaseElementKind::Delimiter
                        if delimiter.is_none() && character_set.is_none() =>
                    {
                        delimiter = Some(parse_delimiter(nested_child)?);
                    },
                    DatabaseElementKind::CharacterSet | DatabaseElementKind::FontCharset
                        if character_set.is_none() =>
                    {
                        character_set = Some(parse_character_set(nested_child)?);
                    },
                    _ => return order_error("db:table-setting"),
                }
            }
            Ok(OdfDatabaseTableSetting {
                is_first_row_header_line: bool_attr(child, "is-first-row-header-line")?,
                show_deleted: bool_attr(child, "show-deleted")?,
                delimiter,
                character_set,
            })
        })
        .collect()
}

fn parse_application(
    element: &DatabaseElement,
) -> Result<OdfDatabaseApplicationConnectionSettings> {
    allow_attrs(
        element,
        &[
            "is-table-name-length-limited",
            "enable-sql92-check",
            "append-table-alias-name",
            "ignore-driver-privileges",
            "boolean-comparison-mode",
            "use-catalog",
            "max-row-count",
            "suppress-version-columns",
        ],
    )?;
    let direct = children(element)?;
    let mut phase = 0u8;
    let mut table_filter = None;
    let mut table_types = None;
    let mut data_source_settings = None;
    for child in direct {
        match child.kind() {
            DatabaseElementKind::TableFilter if phase == 0 => {
                phase = 1;
                table_filter = Some(parse_table_filter(child)?);
            },
            DatabaseElementKind::TableTypeFilter if phase <= 1 => {
                phase = 2;
                if table_types.is_some() {
                    return duplicate("db:table-type-filter");
                }
                table_types = Some(parse_table_types(child)?);
            },
            DatabaseElementKind::DataSourceSettings if phase <= 2 => {
                phase = 3;
                if data_source_settings.is_some() {
                    return duplicate("db:data-source-settings");
                }
                data_source_settings = Some(parse_data_source_settings(child)?);
            },
            _ => return order_error("db:application-connection-settings"),
        }
    }
    let value = OdfDatabaseApplicationConnectionSettings {
        is_table_name_length_limited: bool_attr(element, "is-table-name-length-limited")?,
        enable_sql92_check: bool_attr(element, "enable-sql92-check")?,
        append_table_alias_name: bool_attr(element, "append-table-alias-name")?,
        ignore_driver_privileges: bool_attr(element, "ignore-driver-privileges")?,
        boolean_comparison_mode: element
            .attribute(Some(DB), "boolean-comparison-mode")
            .map(OdfDatabaseBooleanComparisonMode::parse)
            .transpose()?,
        use_catalog: bool_attr(element, "use-catalog")?,
        max_row_count: element
            .attribute(Some(DB), "max-row-count")
            .map(OdfDatabaseInteger::new)
            .transpose()?,
        suppress_version_columns: bool_attr(element, "suppress-version-columns")?,
        table_filter,
        table_types,
        data_source_settings,
    };
    validate_application(&value)?;
    Ok(value)
}

fn parse_table_filter(element: &DatabaseElement) -> Result<OdfDatabaseTableFilter> {
    no_attrs(element)?;
    let direct = children(element)?;
    let mut include = None;
    let mut exclude = None;
    for child in direct {
        match child.kind() {
            DatabaseElementKind::TableIncludeFilter if include.is_none() && exclude.is_none() => {
                include = Some(parse_patterns(child)?);
            },
            DatabaseElementKind::TableExcludeFilter if exclude.is_none() => {
                exclude = Some(parse_patterns(child)?);
            },
            _ => return order_error("db:table-filter"),
        }
    }
    Ok(OdfDatabaseTableFilter { include, exclude })
}

fn parse_patterns(element: &DatabaseElement) -> Result<Vec<String>> {
    no_attrs(element)?;
    let direct = children(element)?;
    if direct.is_empty() {
        return Err(Error::InvalidFormat(
            "table include/exclude filters require at least one pattern".to_string(),
        ));
    }
    bounded(direct.len(), MAX_ITEMS, "table filter patterns")?;
    direct
        .into_iter()
        .map(|child| {
            if child.kind() != DatabaseElementKind::TableFilterPattern {
                return Err(Error::InvalidFormat(
                    "invalid table filter child".to_string(),
                ));
            }
            simple_text(child)
        })
        .collect()
}

fn parse_table_types(element: &DatabaseElement) -> Result<Vec<String>> {
    no_attrs(element)?;
    let direct = children(element)?;
    bounded(direct.len(), MAX_ITEMS, "table types")?;
    direct
        .into_iter()
        .map(|child| {
            if child.kind() != DatabaseElementKind::TableType {
                return Err(Error::InvalidFormat("invalid table type child".to_string()));
            }
            simple_text(child)
        })
        .collect()
}

fn parse_data_source_settings(
    element: &DatabaseElement,
) -> Result<Vec<OdfDatabaseDataSourceSetting>> {
    no_attrs(element)?;
    let direct = children(element)?;
    if direct.is_empty() {
        return Err(Error::InvalidFormat(
            "db:data-source-settings requires at least one setting".to_string(),
        ));
    }
    bounded(direct.len(), MAX_ITEMS, "data source settings")?;
    let mut total_values = 0usize;
    direct
        .into_iter()
        .map(|setting| {
            if setting.kind() != DatabaseElementKind::DataSourceSetting {
                return Err(Error::InvalidFormat(
                    "invalid data source setting child".to_string(),
                ));
            }
            allow_attrs(
                setting,
                &[
                    "data-source-setting-is-list",
                    "data-source-setting-name",
                    "data-source-setting-type",
                ],
            )?;
            let nested = children(setting)?;
            if nested.is_empty() {
                return Err(Error::InvalidFormat(
                    "data source setting requires at least one value".to_string(),
                ));
            }
            total_values = total_values.saturating_add(nested.len());
            bounded(total_values, MAX_VALUES, "data source setting values")?;
            let values = nested
                .into_iter()
                .map(|value| {
                    if value.kind() != DatabaseElementKind::DataSourceSettingValue {
                        return Err(Error::InvalidFormat(
                            "invalid setting value child".to_string(),
                        ));
                    }
                    simple_text(value)
                })
                .collect::<Result<Vec<_>>>()?;
            Ok(OdfDatabaseDataSourceSetting {
                is_list: bool_attr(setting, "data-source-setting-is-list")?,
                name: required_attr(setting, "data-source-setting-name")?.to_string(),
                value_type: OdfDatabaseSettingType::parse(required_attr(
                    setting,
                    "data-source-setting-type",
                )?)?,
                values,
            })
        })
        .collect()
}

fn validate_driver(value: &OdfDatabaseDriverSettings) -> Result<()> {
    let mut aggregate = 0usize;
    for item in [
        value.system_driver_settings.as_deref(),
        value.base_dn.as_deref(),
        value
            .auto_increment
            .as_ref()
            .and_then(|v| v.additional_column_statement.as_deref()),
        value
            .auto_increment
            .as_ref()
            .and_then(|v| v.row_retrieving_statement.as_deref()),
        value.delimiter.as_ref().and_then(|v| v.field.as_deref()),
        value.delimiter.as_ref().and_then(|v| v.string.as_deref()),
        value.delimiter.as_ref().and_then(|v| v.decimal.as_deref()),
        value.delimiter.as_ref().and_then(|v| v.thousand.as_deref()),
        value
            .character_set
            .as_ref()
            .and_then(|v| v.encoding.as_deref()),
    ]
    .into_iter()
    .flatten()
    {
        validate_string(item, &mut aggregate)?;
    }
    if let Some(settings) = &value.table_settings {
        bounded(settings.len(), MAX_ITEMS, "table settings")?;
        for setting in settings {
            for item in delimiter_strings(setting.delimiter.as_ref())
                .into_iter()
                .chain(
                    setting
                        .character_set
                        .as_ref()
                        .and_then(|v| v.encoding.as_deref()),
                )
            {
                validate_string(item, &mut aggregate)?;
            }
        }
    }
    Ok(())
}

fn validate_application(value: &OdfDatabaseApplicationConnectionSettings) -> Result<()> {
    let mut aggregate = 0usize;
    if let Some(filter) = &value.table_filter {
        for patterns in [filter.include.as_ref(), filter.exclude.as_ref()]
            .into_iter()
            .flatten()
        {
            if patterns.is_empty() {
                return Err(Error::InvalidFormat(
                    "present pattern lists must not be empty".to_string(),
                ));
            }
            bounded(patterns.len(), MAX_ITEMS, "table filter patterns")?;
            for pattern in patterns {
                validate_string(pattern, &mut aggregate)?;
            }
        }
    }
    if let Some(types) = &value.table_types {
        bounded(types.len(), MAX_ITEMS, "table types")?;
        for item in types {
            validate_string(item, &mut aggregate)?;
        }
    }
    if let Some(settings) = &value.data_source_settings {
        if settings.is_empty() {
            return Err(Error::InvalidFormat(
                "present data source settings must not be empty".to_string(),
            ));
        }
        bounded(settings.len(), MAX_ITEMS, "data source settings")?;
        let mut count = 0usize;
        for setting in settings {
            validate_string(&setting.name, &mut aggregate)?;
            if setting.values.is_empty() {
                return Err(Error::InvalidFormat(
                    "setting values must not be empty".to_string(),
                ));
            }
            count = count.saturating_add(setting.values.len());
            bounded(count, MAX_VALUES, "setting values")?;
            for item in &setting.values {
                validate_string(item, &mut aggregate)?;
            }
        }
    }
    Ok(())
}

fn delimiter_strings(value: Option<&OdfDatabaseDelimiter>) -> Vec<&str> {
    value
        .map(|v| {
            [
                v.field.as_deref(),
                v.string.as_deref(),
                v.decimal.as_deref(),
                v.thousand.as_deref(),
            ]
            .into_iter()
            .flatten()
            .collect()
        })
        .unwrap_or_default()
}

fn validate_string(value: &str, aggregate: &mut usize) -> Result<()> {
    if value.len() > MAX_VALUE || !value.chars().all(xml_char) {
        return Err(Error::InvalidFormat(
            "invalid or oversized database setting value".to_string(),
        ));
    }
    *aggregate = aggregate.saturating_add(value.len());
    if *aggregate > MAX_AGGREGATE {
        return Err(Error::InvalidFormat(
            "database settings exceed aggregate limit".to_string(),
        ));
    }
    Ok(())
}

fn xml_char(ch: char) -> bool {
    matches!(ch, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&ch)
        || ('\u{E000}'..='\u{FFFD}').contains(&ch)
        || ('\u{10000}'..='\u{10FFFF}').contains(&ch)
}

fn children(element: &DatabaseElement) -> Result<Vec<&DatabaseElement>> {
    let mut output = Vec::new();
    for content in element.content() {
        match content {
            DatabaseContent::Text(value) if value.trim().is_empty() => {},
            DatabaseContent::Text(_) => {
                return Err(Error::InvalidFormat(format!(
                    "{} must not contain character data",
                    element.local_name()
                )));
            },
            DatabaseContent::Element(child) => output.push(child),
        }
    }
    Ok(output)
}

fn simple_text(element: &DatabaseElement) -> Result<String> {
    no_attrs(element)?;
    let mut output = String::new();
    for content in element.content() {
        match content {
            DatabaseContent::Text(value) => output.push_str(value),
            DatabaseContent::Element(_) => {
                return Err(Error::InvalidFormat(format!(
                    "{} accepts only string content",
                    element.local_name()
                )));
            },
        }
    }
    let mut aggregate = 0;
    validate_string(&output, &mut aggregate)?;
    Ok(output)
}

fn empty(element: &DatabaseElement) -> Result<()> {
    if children(element)?.is_empty() {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "{} must be empty",
            element.local_name()
        )))
    }
}

fn no_attrs(element: &DatabaseElement) -> Result<()> {
    if element.attributes().is_empty() {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "{} accepts no attributes",
            element.local_name()
        )))
    }
}

fn allow_attrs(element: &DatabaseElement, allowed: &[&str]) -> Result<()> {
    for attribute in element.attributes() {
        if attribute.namespace_uri() != Some(DB) || !allowed.contains(&attribute.local_name()) {
            return Err(Error::InvalidFormat(format!(
                "unsupported {} attribute {}",
                element.local_name(),
                attribute.local_name()
            )));
        }
    }
    Ok(())
}

fn opt_attr(element: &DatabaseElement, local: &str) -> Option<String> {
    element.attribute(Some(DB), local).map(str::to_string)
}

fn required_attr<'a>(element: &'a DatabaseElement, local: &str) -> Result<&'a str> {
    element
        .attribute(Some(DB), local)
        .ok_or_else(|| Error::InvalidFormat(format!("{} requires {local}", element.local_name())))
}

fn bool_attr(element: &DatabaseElement, local: &str) -> Result<Option<bool>> {
    element
        .attribute(Some(DB), local)
        .map(|value| match value {
            "true" => Ok(true),
            "false" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "{local} must be true or false"
            ))),
        })
        .transpose()
}

fn bounded(value: usize, limit: usize, name: &str) -> Result<()> {
    if value > limit {
        Err(Error::InvalidFormat(format!("{name} exceeds {limit}")))
    } else {
        Ok(())
    }
}

fn duplicate<T>(name: &str) -> Result<T> {
    Err(Error::InvalidFormat(format!("duplicate {name}")))
}

fn order_error<T>(name: &str) -> Result<T> {
    Err(Error::InvalidFormat(format!(
        "{name} child order or cardinality is invalid"
    )))
}

fn attr_bool(xml: &mut String, local: &str, value: Option<bool>) {
    if let Some(value) = value {
        attr(xml, local, if value { "true" } else { "false" });
    }
}

fn attr_opt(xml: &mut String, local: &str, value: Option<&str>) {
    if let Some(value) = value {
        attr(xml, local, value);
    }
}

fn attr(xml: &mut String, local: &str, value: &str) {
    xml.push_str(" db:");
    xml.push_str(local);
    xml.push_str("=\"");
    attribute(xml, value);
    xml.push('"');
}

fn attribute(xml: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '"' => xml.push_str("&quot;"),
            '\r' => xml.push_str("&#13;"),
            '\n' => xml.push_str("&#10;"),
            '\t' => xml.push_str("&#9;"),
            _ => xml.push(ch),
        }
    }
}

fn text(xml: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            _ => xml.push(ch),
        }
    }
}

fn push_delimiter(xml: &mut String, value: &OdfDatabaseDelimiter) {
    xml.push_str("<db:delimiter");
    attr_opt(xml, "field", value.field.as_deref());
    attr_opt(xml, "string", value.string.as_deref());
    attr_opt(xml, "decimal", value.decimal.as_deref());
    attr_opt(xml, "thousand", value.thousand.as_deref());
    xml.push_str("/>");
}

fn push_character_set(xml: &mut String, value: &OdfDatabaseCharacterSet) {
    xml.push_str("<db:character-set");
    attr_opt(xml, "encoding", value.encoding.as_deref());
    xml.push_str("/>");
}

fn push_patterns(xml: &mut String, local: &str, values: Option<&[String]>) {
    if let Some(values) = values {
        xml.push_str("<db:");
        xml.push_str(local);
        xml.push('>');
        for value in values {
            xml.push_str("<db:table-filter-pattern>");
            text(xml, value);
            xml.push_str("</db:table-filter-pattern>");
        }
        xml.push_str("</db:");
        xml.push_str(local);
        xml.push('>');
    }
}

#[derive(Clone, Copy)]
enum SettingsTarget {
    Driver,
    Application,
}

#[derive(Default)]
struct SourceSpans {
    connection_end: usize,
    data_source_end_start: usize,
    driver: Option<(usize, usize)>,
    application: Option<(usize, usize)>,
}

fn mutate_settings(
    xml: &str,
    target: SettingsTarget,
    replacement: Option<String>,
) -> Result<String> {
    parse_database_trailing_settings_xml(xml)?;
    let spans = locate_spans(xml)?;
    let old = match target {
        SettingsTarget::Driver => spans.driver,
        SettingsTarget::Application => spans.application,
    };
    let (start, end) = if let Some(span) = old {
        span
    } else {
        let insertion = match target {
            SettingsTarget::Driver => spans
                .application
                .map(|span| span.0)
                .unwrap_or(spans.data_source_end_start),
            SettingsTarget::Application => spans.data_source_end_start,
        };
        (insertion, insertion)
    };
    let replacement = replacement.unwrap_or_default();
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(&replacement);
    output.push_str(&xml[end..]);
    Ok(output)
}

fn preflight(xml: &str) -> Result<()> {
    if xml.len() > MAX_XML {
        return Err(Error::InvalidFormat(
            "database settings XML is too large".to_string(),
        ));
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid database settings XML: {error}"))
            })?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let namespace = namespace_value(&namespace)?;
                let local_name = element.local_name();
                let local = std::str::from_utf8(local_name.as_ref())
                    .map_err(|_| Error::InvalidFormat("non-UTF-8 element name".to_string()))?;
                if (namespace.as_deref() == Some(OFFICE) && local == "event-listeners")
                    || namespace.as_deref() == Some(SCRIPT)
                {
                    return Err(Error::InvalidFormat(
                        "active database settings are rejected".to_string(),
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(Error::InvalidFormat(
                    "active XML constructs are rejected".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn locate_spans(xml: &str) -> Result<SourceSpans> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut stack: Vec<(Option<String>, String)> = Vec::new();
    let mut active: Vec<(String, usize, usize)> = Vec::new();
    let mut spans = SourceSpans::default();
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid settings XML: {error}")))?;
        let namespace = namespace_value(&namespace)?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                let local = owned_local(element.local_name().as_ref())?;
                if namespace.as_deref() == Some(DB)
                    && matches!(
                        local.as_str(),
                        "connection-data" | "driver-settings" | "application-connection-settings"
                    )
                    && stack.last().is_some_and(|(ns, parent)| {
                        ns.as_deref() == Some(DB) && parent == "data-source"
                    })
                {
                    active.push((local.clone(), stack.len(), start));
                }
                stack.push((namespace, local));
            },
            Event::Empty(ref element) => {
                let local = owned_local(element.local_name().as_ref())?;
                if namespace.as_deref() == Some(DB)
                    && stack.last().is_some_and(|(ns, parent)| {
                        ns.as_deref() == Some(DB) && parent == "data-source"
                    })
                {
                    set_span(&mut spans, &local, start, end);
                }
            },
            Event::End(ref element) => {
                let local = owned_local(element.local_name().as_ref())?;
                let depth = stack.len().saturating_sub(1);
                if namespace.as_deref() == Some(DB) && local == "data-source" {
                    spans.data_source_end_start = start;
                }
                if let Some(position) = active
                    .iter()
                    .rposition(|(name, active_depth, _)| name == &local && *active_depth == depth)
                {
                    let (name, _, span_start) = active.remove(position);
                    set_span(&mut spans, &name, span_start, end);
                }
                stack.pop();
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if spans.connection_end == 0 || spans.data_source_end_start == 0 {
        return Err(Error::InvalidFormat(
            "could not locate database source settings".to_string(),
        ));
    }
    Ok(spans)
}

fn set_span(spans: &mut SourceSpans, local: &str, start: usize, end: usize) {
    match local {
        "connection-data" => spans.connection_end = end,
        "driver-settings" => spans.driver = Some((start, end)),
        "application-connection-settings" => spans.application = Some((start, end)),
        _ => {},
    }
}

fn namespace_value(value: &ResolveResult<'_>) -> Result<Option<String>> {
    match value {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(value)) => std::str::from_utf8(value)
            .map(|value| Some(value.to_string()))
            .map_err(|_| Error::InvalidFormat("non-UTF-8 namespace".to_string())),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown namespace prefix {}",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn owned_local(value: &[u8]) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat("non-UTF-8 element name".to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;

    fn wrap(driver: &str, application: &str) -> String {
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:d="{DB}" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource x:type="simple" x:href="db"/></d:connection-data>{driver}{application}</d:data-source><d:forms/></o:database></o:body></o:document-content>"#
        )
    }

    #[test]
    fn parses_complete_driver_and_application_families_and_roundtrips() {
        let driver = r#"<d:driver-settings d:show-deleted="false" d:system-driver-settings="x" d:base-dn="dc=x" d:is-first-row-header-line="true" d:parameter-name-substitution="false"><d:auto-increment d:additional-column-statement="AUTO" d:row-retrieving-statement="SELECT"/><d:delimiter d:field="," d:string="&quot;" d:decimal="." d:thousand=" "/><d:font-charset d:encoding="UTF-8"/><d:table-settings><d:table-setting d:show-deleted="true"><d:delimiter d:field=";"/><d:character-set d:encoding="UTF-16"/></d:table-setting><d:table-setting/></d:table-settings></d:driver-settings>"#;
        let application = r#"<d:application-connection-settings d:is-table-name-length-limited="true" d:enable-sql92-check="false" d:append-table-alias-name="true" d:ignore-driver-privileges="false" d:boolean-comparison-mode="equal-integer" d:use-catalog="true" d:max-row-count="+000100" d:suppress-version-columns="false"><d:table-filter><d:table-include-filter><d:table-filter-pattern>A%</d:table-filter-pattern></d:table-include-filter><d:table-exclude-filter><d:table-filter-pattern>tmp%</d:table-filter-pattern></d:table-exclude-filter></d:table-filter><d:table-type-filter><d:table-type>TABLE</d:table-type></d:table-type-filter><d:data-source-settings><d:data-source-setting d:data-source-setting-is-list="true" d:data-source-setting-name="Options" d:data-source-setting-type="string"><d:data-source-setting-value>a</d:data-source-setting-value><d:data-source-setting-value>b</d:data-source-setting-value></d:data-source-setting></d:data-source-settings></d:application-connection-settings>"#;
        let parsed = parse_database_trailing_settings_xml(&wrap(driver, application)).unwrap();
        assert_eq!(
            parsed
                .application
                .as_ref()
                .unwrap()
                .max_row_count
                .as_ref()
                .unwrap()
                .as_str(),
            "100"
        );
        assert_eq!(
            parsed
                .driver
                .as_ref()
                .unwrap()
                .table_settings
                .as_ref()
                .unwrap()
                .len(),
            2
        );
        let canonical_driver = parsed.driver.as_ref().unwrap().to_xml_fragment().unwrap();
        assert!(canonical_driver.contains("<db:character-set"));
        assert!(!canonical_driver.contains("font-charset"));
        let canonical_application = parsed
            .application
            .as_ref()
            .unwrap()
            .to_xml_fragment()
            .unwrap();
        let reparsed =
            parse_database_trailing_settings_xml(&wrap(&canonical_driver, &canonical_application))
                .unwrap();
        assert_eq!(reparsed, parsed);
    }

    #[test]
    fn rejects_order_cardinality_lexical_active_and_bounded_violations() {
        let invalid = [
            (
                "<d:driver-settings><d:delimiter/><d:auto-increment/></d:driver-settings>",
                "",
            ),
            ("<d:driver-settings d:show-deleted=\"1\"/>", ""),
            (
                "<d:driver-settings><d:table-settings><d:bad/></d:table-settings></d:driver-settings>",
                "",
            ),
            (
                "",
                "<d:application-connection-settings><d:data-source-settings/></d:application-connection-settings>",
            ),
            (
                "",
                "<d:application-connection-settings><d:table-filter><d:table-include-filter/></d:table-filter></d:application-connection-settings>",
            ),
            (
                "",
                "<d:application-connection-settings><d:data-source-settings><d:data-source-setting d:data-source-setting-name=\"x\" d:data-source-setting-type=\"byte\"><d:data-source-setting-value>x</d:data-source-setting-value></d:data-source-setting></d:data-source-settings></d:application-connection-settings>",
            ),
            (
                "",
                "<d:application-connection-settings><o:event-listeners/></d:application-connection-settings>",
            ),
        ];
        for (driver, application) in invalid {
            assert!(parse_database_trailing_settings_xml(&wrap(driver, application)).is_err());
        }
        assert!(OdfDatabaseInteger::new("+").is_err());
        assert_eq!(OdfDatabaseInteger::new("-000").unwrap().as_str(), "0");
        let doctype = format!("<!DOCTYPE x>{}", wrap("", ""));
        assert!(parse_database_trailing_settings_xml(&doctype).is_err());
    }

    #[test]
    fn package_accessors_and_lossless_set_remove_preserve_other_source_content() {
        let driver = OdfDatabaseDriverSettings {
            parameter_name_substitution: Some(false),
            ..OdfDatabaseDriverSettings::default()
        };
        let application = OdfDatabaseApplicationConnectionSettings {
            max_row_count: Some(OdfDatabaseInteger::new("25").unwrap()),
            ..OdfDatabaseApplicationConnectionSettings::default()
        };
        let original = wrap("", "");
        let with_driver = set_database_driver_settings_xml(&original, Some(&driver)).unwrap();
        let both =
            set_database_application_connection_settings_xml(&with_driver, Some(&application))
                .unwrap();
        assert!(both.contains("<d:forms/>"));
        let removed = set_database_driver_settings_xml(&both, None).unwrap();
        let parsed = parse_database_trailing_settings_xml(&removed).unwrap();
        assert!(parsed.driver.is_none());
        assert_eq!(parsed.application, Some(application.clone()));

        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_DATABASE).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, both.as_bytes())
            .unwrap();
        let document = DatabaseDocument::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
        assert_eq!(document.driver_settings().unwrap(), Some(driver));
        assert_eq!(
            document.application_connection_settings().unwrap(),
            Some(application)
        );
    }
}
