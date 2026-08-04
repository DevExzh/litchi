//! XLSB External Data Connections part (MS-XLSB 2.1.7.24).
//!
//! A typed, inert model of the workbook's external connections: ODBC and
//! OLE DB command properties, OLAP and Web connection properties,
//! connection parameters, and Web query table references. Connection
//! strings, commands, URLs, file paths, and credential metadata are stored
//! exactly as declared and are never resolved, opened, contacted,
//! refreshed, or executed.

mod model;
pub(crate) mod package;
mod parse;
#[cfg(test)]
mod tests;
pub(crate) mod write;

pub use model::{
    CommandType, Connection, Connections, CredentialMethod, DbProperties, HtmlFormat,
    OlapProperties, Parameter, ParameterType, ParameterValue, PasswordState, Properties,
    ReconnectionType, SourceType, WebProperties, WebTableItem,
};
pub use parse::parse_connections_part;
