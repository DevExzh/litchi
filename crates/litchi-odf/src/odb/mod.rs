//! OpenDocument Database Front End (`.odb`) support.

mod components;
mod connection_data;
mod document;
mod query;
mod schema;
mod settings;

pub use components::*;
pub use connection_data::{
    OdfDatabaseConnectionData, OdfDatabaseConnectionSource, OdfDatabaseFileSource,
    OdfDatabaseLogin, OdfDatabaseLoginIdentity, OdfDatabasePositiveInteger,
    OdfDatabaseServerLocation, OdfDatabaseServerSource, OdfOdbConnectionResource,
    parse_database_connection_data_xml, replace_database_connection_data_xml,
};
pub use document::{
    DatabaseAttribute, DatabaseContent, DatabaseDocument, DatabaseElement, DatabaseElementKind,
};
pub use query::{
    OdfDatabaseColumn, OdfDatabaseColumnValue, OdfDatabaseQueries, OdfDatabaseQuery,
    OdfDatabaseQueryCollection, OdfDatabaseQueryItem, OdfDatabaseQueryModel, OdfDatabaseStatement,
    OdfDatabaseTableRepresentation, OdfDatabaseTableRepresentations, OdfDatabaseUpdateTable,
    parse_database_queries_xml, parse_database_query_model_xml,
    parse_database_table_representations_xml, set_database_queries_xml,
    set_database_table_representations_xml,
};
pub use schema::*;
pub use settings::{
    OdfDatabaseApplicationConnectionSettings, OdfDatabaseAutoIncrement,
    OdfDatabaseBooleanComparisonMode, OdfDatabaseCharacterSet, OdfDatabaseDataSourceSetting,
    OdfDatabaseDelimiter, OdfDatabaseDriverSettings, OdfDatabaseInteger, OdfDatabaseSettingType,
    OdfDatabaseTableFilter, OdfDatabaseTableSetting, OdfDatabaseTrailingSettings,
    parse_database_trailing_settings_xml, set_database_application_connection_settings_xml,
    set_database_driver_settings_xml,
};
