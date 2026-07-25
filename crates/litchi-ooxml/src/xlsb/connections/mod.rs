//! XLSB External Data Connections part (MS-XLSB 2.1.7.24).
//!
//! A typed, inert model of the workbook's external connections: ODBC and
//! OLE DB command properties, OLAP and Web connection properties,
//! connection parameters, and Web query table references. Connection
//! strings, commands, URLs, file paths, and credential metadata are stored
//! exactly as declared and are never resolved, opened, contacted,
//! refreshed, or executed.

mod model;
mod parse;
#[cfg(test)]
mod tests;

pub use model::{
    XlsbCommandType, XlsbConnection, XlsbConnectionParameter, XlsbConnectionProperties,
    XlsbConnectionSourceType, XlsbConnections, XlsbCredentialMethod, XlsbDbProperties,
    XlsbHtmlFormat, XlsbOlapProperties, XlsbParameterType, XlsbParameterValue, XlsbPasswordState,
    XlsbReconnectionType, XlsbWebProperties, XlsbWebTableItem,
};
pub use parse::parse_connections_part;
