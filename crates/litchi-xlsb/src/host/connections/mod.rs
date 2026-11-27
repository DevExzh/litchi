//! XLSB External Data Connections part (MS-XLSB 2.1.7.24).
//!
//! A typed, inert model of the workbook's external connections: ODBC and
//! OLE DB command properties, OLAP, Web and classic text-import properties,
//! connection parameters, and Web query table references. Connection
//! strings, commands, URLs, file paths, and credential metadata are stored
//! exactly as declared and are never resolved, opened, contacted,
//! refreshed, or executed.

pub mod codec;
pub mod model;
pub mod package;
mod parse;
#[cfg(test)]
mod tests;
mod text;
mod transaction;
pub mod validation;
pub(crate) mod write;

pub use model::{
    CommandType, Connection, Connections, CredentialMethod, DbProperties, HtmlFormat,
    OlapProperties, Parameter, ParameterType, ParameterValue, PasswordState, Properties,
    ReconnectionType, SourceType, TextCodePage, TextField, TextFieldType, TextImportProperties,
    TextQualifier, UnknownRecord, WebProperties, WebTableItem,
};
pub use parse::parse_connections_part;
pub use transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Snapshot};

pub type Result<T> = crate::package::error::Result<T>;

pub(crate) fn invalid(detail: &'static str) -> Error {
    Error::Unrecognized {
        typ: "XLSB External Data Connections",
        val: detail,
    }
}
