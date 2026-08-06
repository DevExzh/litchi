//! XLSB External Data Connections part (MS-XLSB 2.1.7.24).
//!
//! A typed, inert model of the workbook's external connections: ODBC and
//! OLE DB command properties, OLAP and Web connection properties,
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
pub mod transaction;
pub mod validation;
pub(crate) mod write;

pub use model::{
    CommandType, Connection, Connections, CredentialMethod, DbProperties, HtmlFormat,
    OlapProperties, Parameter, ParameterType, ParameterValue, PasswordState, Properties,
    ReconnectionType, SourceType, UnknownRecord, WebProperties, WebTableItem,
};
pub use parse::parse_connections_part;
pub use transaction::{Commit, Patch, Snapshot, Transaction, apply, read};

pub type Result<T> = crate::package::error::Result<T>;

pub(crate) fn invalid(detail: impl Into<String>) -> crate::package::error::Error {
    crate::package::error::Error::Unrecognized {
        typ: "XLSB External Data Connections".to_string(),
        val: detail.into(),
    }
}
