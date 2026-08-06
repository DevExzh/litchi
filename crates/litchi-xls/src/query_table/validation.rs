use litchi_core::binary;

use super::super::pivot_table::parse_qsi_sx_tag;
use super::super::records::Encoding;
use super::super::utils::parse_string_record;
use super::{codec::*, model::*};

#[derive(Debug)]
pub(crate) enum ExtContext {
    None,
    TableNames,
    OleDb(usize),
    TextQuery,
}

/// In-progress assembly of one `QUERYTABLE` sequence.
#[derive(Debug)]
pub(crate) struct QueryTableBuild {
    pub(crate) table: QueryTable,
    pub(crate) dbquery_seen: bool,
    /// Previous record was an SXString or a ParamQry: a following 0x00DC is
    /// a ParamQry rather than a DbQuery (MS-XLS 2.4.79).
    pub(crate) last_string_or_param: bool,
    pub(crate) remaining_query: u16,
    pub(crate) remaining_odbc_conn: u16,
    pub(crate) remaining_web_post: u16,
    pub(crate) remaining_sql_sav: u16,
    pub(crate) pending_param_name: Option<String>,
    pub(crate) query_chunks: Vec<String>,
    pub(crate) odbc_conn_chunks: Vec<String>,
    pub(crate) web_post_chunks: Vec<String>,
    pub(crate) sql_sav_chunks: Vec<String>,
    pub(crate) ext_context: ExtContext,
    pub(crate) ole_db_remaining: u16,
    pub(crate) in_sxaddl_qsi: bool,
    pub(crate) sort_data_remaining: Option<u32>,
}

impl QueryTableBuild {
    fn new(table: QueryTable) -> Self {
        Self {
            table,
            dbquery_seen: false,
            last_string_or_param: false,
            remaining_query: 0,
            remaining_odbc_conn: 0,
            remaining_web_post: 0,
            remaining_sql_sav: 0,
            pending_param_name: None,
            query_chunks: Vec::new(),
            odbc_conn_chunks: Vec::new(),
            web_post_chunks: Vec::new(),
            sql_sav_chunks: Vec::new(),
            ext_context: ExtContext::None,
            ole_db_remaining: 0,
            in_sxaddl_qsi: false,
            sort_data_remaining: None,
        }
    }

    fn finish(mut self) -> QueryTable {
        if !self.query_chunks.is_empty() {
            self.table.command_text = Some(self.query_chunks.concat());
        }
        if !self.odbc_conn_chunks.is_empty() {
            self.table.connection_string = Some(self.odbc_conn_chunks.concat());
        }
        if !self.web_post_chunks.is_empty() {
            self.table.web_post = Some(self.web_post_chunks.concat());
        }
        if !self.sql_sav_chunks.is_empty() {
            self.table.sql_server_fields = Some(self.sql_sav_chunks.concat());
        }
        self.table
    }
}
/// Ordered worksheet `QUERYTABLE` sequence collector. Multiple query tables
/// per sheet are supported. See the module documentation for the inertness
/// contract and the `SORTDATA12` interaction.
#[derive(Debug, Default)]
pub(crate) struct QueryTableCollector {
    completed: Vec<QueryTable>,
    current: Option<QueryTableBuild>,
}

impl QueryTableCollector {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    fn finalize_current(&mut self) {
        if let Some(build) = self.current.take() {
            self.completed.push(build.finish());
        }
    }

    /// Returns true when the record belongs to a `QUERYTABLE` sequence.
    ///
    /// Never fails: malformed core records drop the in-progress sequence and
    /// malformed optional records are ignored, so a broken query table can
    /// not abort worksheet parsing.
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> bool {
        if record_type == QSI_RECORD_TYPE {
            // A new Qsi starts a new sequence (one query table each).
            self.finalize_current();
            if let Some(table) = parse_qsi(data) {
                self.current = Some(QueryTableBuild::new(table));
            }
            return true;
        }

        let Some(build) = self.current.as_mut() else {
            return false;
        };

        // A pending SortData only accepts its declared ContinueFrt12 records.
        if let Some(remaining) = build.sort_data_remaining {
            if record_type == CONTINUE_FRT12_RECORD_TYPE && remaining > 0 {
                build.table.sort_data_bytes.extend_from_slice(data);
                build.sort_data_remaining = Some(remaining - 1);
                return true;
            }
            build.sort_data_remaining = None;
            self.finalize_current();
            return false;
        }

        match record_type {
            DB_OR_PARAM_QRY_RECORD_TYPE => {
                // MS-XLS 2.4.79: after an SXString or a ParamQry the record
                // is a ParamQry; after anything else it is a DbQuery.
                let was_param = build.last_string_or_param;
                let parsed = if was_param {
                    parse_param_qry(build, data)
                } else {
                    parse_db_query(build, data)
                };
                if parsed.is_none() {
                    // Malformed core record: drop the in-progress sequence.
                    self.current = None;
                    return true;
                }
                // A DbQuery restarts the disambiguation chain; a ParamQry
                // extends it.
                self.current
                    .as_mut()
                    .expect("build present")
                    .last_string_or_param = was_param;
                true
            },
            SX_STRING_RECORD_TYPE => {
                let Ok(text) = parse_string_record(data, &Encoding::Utf16Le) else {
                    // Malformed chunk: drop the in-progress sequence.
                    self.current = None;
                    return true;
                };
                let build = self.current.as_mut().expect("build present");
                if build.remaining_query > 0 {
                    build.remaining_query -= 1;
                    build.query_chunks.push(text);
                } else if build.remaining_odbc_conn > 0 {
                    build.remaining_odbc_conn -= 1;
                    build.odbc_conn_chunks.push(text);
                } else if build.remaining_web_post > 0 {
                    build.remaining_web_post -= 1;
                    build.web_post_chunks.push(text);
                } else if build.remaining_sql_sav > 0 {
                    build.remaining_sql_sav -= 1;
                    build.sql_sav_chunks.push(text);
                } else {
                    // A parameter name preceding its ParamQry record.
                    build.pending_param_name = Some(text);
                }
                build.last_string_or_param = true;
                true
            },
            QSI_SX_TAG_RECORD_TYPE => {
                let Ok(tag) = parse_qsi_sx_tag(data) else {
                    // Malformed tag: ignored, the sequence continues.
                    return true;
                };
                if tag.table_type != 0 {
                    // fSx=1: the tag and its collection belong to a
                    // PivotTable view; hand it back to the pivot collector.
                    self.finalize_current();
                    return false;
                }
                let build = self.current.as_mut().expect("build present");
                if tag.table_name == build.table.name {
                    build.table.enable_refresh = Some(tag.flags & 0x0001 != 0);
                    build.table.qsi_future = tag.options;
                }
                // Name mismatches are ignored per MS-XLS 2.4.211.
                true
            },
            DB_QUERY_EXT_RECORD_TYPE => {
                if parse_db_query_ext(build, data).is_none() {
                    // Malformed core record: drop the in-progress sequence.
                    self.current = None;
                }
                true
            },
            EXT_STRING_RECORD_TYPE => {
                let text = if data.len() >= 7 {
                    parse_string_record(&data[4..], &Encoding::Utf16Le).ok()
                } else {
                    None
                };
                let Some(text) = text else { return true };
                let build = self.current.as_mut().expect("build present");
                match build.ext_context {
                    ExtContext::TableNames => {
                        build.table.table_names = Some(text);
                        build.ext_context = ExtContext::None;
                    },
                    ExtContext::OleDb(index) => {
                        if let Some(connection) = build.table.ole_db_connections.get_mut(index) {
                            connection.connection_string.push_str(&text);
                        }
                        build.ole_db_remaining = build.ole_db_remaining.saturating_sub(1);
                        if build.ole_db_remaining == 0 {
                            build.ext_context = ExtContext::None;
                        }
                    },
                    ExtContext::TextQuery => {
                        if let Some(text_query) = build.table.text_query.as_mut() {
                            text_query.connection_string.push_str(&text);
                        }
                    },
                    ExtContext::None => {},
                }
                true
            },
            TXT_QRY_RECORD_TYPE => {
                if let Some(text_query) = parse_txt_qry(data) {
                    build.table.text_query = Some(Box::new(text_query));
                    build.ext_context = ExtContext::TextQuery;
                }
                true
            },
            OLE_DB_CONN_RECORD_TYPE => {
                if data.len() >= 8
                    && binary::read_u16_le_at(data, 0).ok() == Some(OLE_DB_CONN_RECORD_TYPE)
                {
                    let flags = binary::read_u16_le_at(data, 4).unwrap_or(0);
                    build.table.ole_db_connections.push(OleDbConnection {
                        password_stripped: flags & OLECONN_PASSWD_STRIPPED != 0,
                        local: flags & OLECONN_LOCAL != 0,
                        connection_string: String::new(),
                    });
                    build.ole_db_remaining = binary::read_u16_le_at(data, 6).unwrap_or(0);
                    build.ext_context = ExtContext::OleDb(build.table.ole_db_connections.len() - 1);
                }
                true
            },
            // QSIR formatting records are consumed but not interpreted.
            QSIR_RECORD_TYPE | QSIF_RECORD_TYPE => true,
            SXADDL_RECORD_TYPE => {
                let sxc_qsi = data.len() >= 6 && data[4] == SXC_QSI_CLASS;
                if sxc_qsi {
                    build.in_sxaddl_qsi = data[5] != SXD_END;
                    true
                } else {
                    // Another SXAddl class: not part of this sequence.
                    self.finalize_current();
                    false
                }
            },
            SORT_DATA_RECORD_TYPE => {
                build.table.sort_data_bytes.extend_from_slice(data);
                let conditions = if data.len() >= 34 {
                    binary::read_u32_le_at(data, 30).unwrap_or(0)
                } else {
                    0
                };
                build.sort_data_remaining = Some(conditions);
                true
            },
            _ => {
                self.finalize_current();
                false
            },
        }
    }

    pub(crate) fn finish(mut self) -> Vec<QueryTable> {
        self.finalize_current();
        self.completed
    }
}
