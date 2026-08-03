//! Compatibility access to the canonical XLSX external-connections owner.

use litchi_core::sheet::Result;
use litchi_opc::OpcPackage;

pub use litchi_xlsx::connections::{
    Connection, ConnectionParameter, Connections, CredentialsMethod, DatabaseProperties,
    HtmlFormatting, OlapProperties, ParameterType, TextField, TextFieldType, TextFileType,
    TextImportProperties, TextQualifier, WebQueryProperties, WebTableSelector, load_from_package,
    remove_from_package,
};

/// Store through the canonical owner while retaining the migration host's
/// complete query-table parser for cross-part validation.
pub fn store_in_package(package: &mut OpcPackage, value: &Connections, strict: bool) -> Result<()> {
    litchi_xlsx::connections::store_in_package_with_query_table_validator(
        package,
        value,
        strict,
        |xml| {
            crate::xlsx::query_table::parse_query_table(xml)
                .map(|table| table.connection_id())
                .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)
        },
    )
}
