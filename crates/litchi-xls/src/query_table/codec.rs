use litchi_core::binary;

use super::super::records::Encoding;
use super::super::utils::parse_string_record;
use super::model::*;
use super::validation::{ExtContext, QueryTableBuild};

/// `Qsi` record type (MS-XLS 2.4.208).
pub(crate) const QSI_RECORD_TYPE: u16 = 0x01AD;
/// `DbOrParamQry` record type: DbQuery or ParamQry (MS-XLS 2.4.79).
pub(crate) const DB_OR_PARAM_QRY_RECORD_TYPE: u16 = 0x00DC;
/// `SXString` record type (MS-XLS 2.4.304).
pub(crate) const SX_STRING_RECORD_TYPE: u16 = 0x00CD;
/// `QsiSXTag` record type (MS-XLS 2.4.211).
pub(crate) const QSI_SX_TAG_RECORD_TYPE: u16 = 0x0802;
/// `DBQueryExt` record type (MS-XLS 2.4.81).
pub(crate) const DB_QUERY_EXT_RECORD_TYPE: u16 = 0x0803;
/// `ExtString` record type (MS-XLS 2.4.108).
pub(crate) const EXT_STRING_RECORD_TYPE: u16 = 0x0804;
/// `TxtQry` record type (MS-XLS 2.4.330).
pub(crate) const TXT_QRY_RECORD_TYPE: u16 = 0x0805;
/// `Qsir` record type (MS-XLS 2.4.210).
pub(crate) const QSIR_RECORD_TYPE: u16 = 0x0806;
/// `Qsif` record type (MS-XLS 2.4.209).
pub(crate) const QSIF_RECORD_TYPE: u16 = 0x0807;
/// `OleDbConn` record type (MS-XLS 2.4.186).
pub(crate) const OLE_DB_CONN_RECORD_TYPE: u16 = 0x080A;

/// `SXAddl` record type (MS-XLS 2.4.273.1); only the `SxcQsi` class belongs
/// to a `QUERYTABLE` sequence.
pub(crate) const SXADDL_RECORD_TYPE: u16 = 0x0864;
/// `SortData` record type (MS-XLS 2.4.264); trailing `SORTDATA12` member.
pub(crate) const SORT_DATA_RECORD_TYPE: u16 = 0x0895;
/// `ContinueFrt12` record type (MS-XLS 2.4.62); continues a `SortData`.
pub(crate) const CONTINUE_FRT12_RECORD_TYPE: u16 = 0x087F;
/// `SXAddl` class of the `SxcQsi` class records (MS-XLS 2.2.5.1.1).
pub(crate) const SXC_QSI_CLASS: u8 = 0x05;
/// `SXAddl` record kind ending a class sequence.
pub(crate) const SXD_END: u8 = 0xFF;
/// Maximum number of `TxtWf` field descriptors in a `TxtQry` record.
pub(crate) const MAX_TXT_FIELDS: usize = 256;

// Qsi flag bits (first flag word).
pub(crate) const QSI_TITLES: u16 = 0x0001;
pub(crate) const QSI_ROW_NUMS: u16 = 0x0002;
pub(crate) const QSI_DISABLE_REFRESH: u16 = 0x0004;
pub(crate) const QSI_ASYNC: u16 = 0x0008;
pub(crate) const QSI_NEW_ASYNC: u16 = 0x0010;
pub(crate) const QSI_AUTO_REFRESH: u16 = 0x0020;
pub(crate) const QSI_SHRINK: u16 = 0x0040;
pub(crate) const QSI_FILL: u16 = 0x0080;
pub(crate) const QSI_AUTO_FORMAT: u16 = 0x0100;
pub(crate) const QSI_SAVE_DATA: u16 = 0x0200;
pub(crate) const QSI_DISABLE_EDIT: u16 = 0x0400;
pub(crate) const QSI_OVERWRITE: u16 = 0x2000;

// Qsi AutoFormat attribute bits (second flag word).
pub(crate) const QSI_ATR_NUM: u16 = 0x0001;
pub(crate) const QSI_ATR_FNT: u16 = 0x0002;
pub(crate) const QSI_ATR_ALC: u16 = 0x0004;
pub(crate) const QSI_ATR_BDR: u16 = 0x0008;
pub(crate) const QSI_ATR_PAT: u16 = 0x0010;
pub(crate) const QSI_ATR_PROT: u16 = 0x0020;

// DbQuery flag bits.
pub(crate) const DBQUERY_DBT_MASK: u16 = 0x0007;
pub(crate) const DBQUERY_ODBC_CONN: u16 = 0x0008;
pub(crate) const DBQUERY_SQL: u16 = 0x0010;
pub(crate) const DBQUERY_SQL_SAV: u16 = 0x0020;
pub(crate) const DBQUERY_WEB: u16 = 0x0040;
pub(crate) const DBQUERY_SAVE_PWD: u16 = 0x0080;
pub(crate) const DBQUERY_TABLES_ONLY_HTML: u16 = 0x0100;

// DBQueryExt flag bits.
pub(crate) const DBEXT_MAINTAIN: u16 = 0x0001;
pub(crate) const DBEXT_NEW_QUERY: u16 = 0x0002;
pub(crate) const DBEXT_IMPORT_XML_SOURCE: u16 = 0x0004;
pub(crate) const DBEXT_SP_LIST_SRC: u16 = 0x0008;
pub(crate) const DBEXT_SP_LIST_REINIT_CACHE: u16 = 0x0010;
pub(crate) const DBEXT_SRC_IS_XML: u16 = 0x0080;

// DBQueryExt trailing flag bits.
pub(crate) const DBEXT_TABLE_NAMES: u16 = 0x0002;

// TxtQry flag bits (first flag word).
pub(crate) const TXT_FILE: u16 = 0x0001;
pub(crate) const TXT_DELIMITED: u16 = 0x0002;
pub(crate) const TXT_CPID_SHIFT: u16 = 2;
pub(crate) const TXT_CPID_MASK: u16 = 0x0003;
pub(crate) const TXT_PROMPT_FOR_FILE: u16 = 0x0010;
pub(crate) const TXT_USE_NEW_CPID: u16 = 0x8000;

// TxtQry delimiter flag bits (second flag byte).
pub(crate) const TXT_DELIM_TAB: u8 = 0x01;
pub(crate) const TXT_DELIM_SPACE: u8 = 0x02;
pub(crate) const TXT_DELIM_COMMA: u8 = 0x04;
pub(crate) const TXT_DELIM_SEMICOLON: u8 = 0x08;
pub(crate) const TXT_DELIM_CUSTOM: u8 = 0x10;
pub(crate) const TXT_DELIM_CONSECUTIVE: u8 = 0x20;
pub(crate) const TXT_TEXT_DELIM_SHIFT: u8 = 6;
pub(crate) const TXT_TEXT_DELIM_MASK: u8 = 0x03;

// OleDbConn flag bits.
pub(crate) const OLECONN_PASSWD_STRIPPED: u16 = 0x0001;
pub(crate) const OLECONN_LOCAL: u16 = 0x0002;

// ParamQry fixed flags.
pub(crate) const PARAMQRY_PBT_MASK: u16 = 0x0003;
pub(crate) const PARAMQRY_NON_DEFAULT_NAME: u16 = 0x0004;
pub(crate) fn parse_qsi(data: &[u8]) -> Option<QueryTable> {
    if data.len() < 13 {
        return None;
    }
    let flags = binary::read_u16_le_at(data, 0).ok()?;
    let attributes = binary::read_u16_le_at(data, 2).ok()?;
    let name = parse_string_record(&data[10..], &Encoding::Utf16Le).ok()?;
    Some(QueryTable {
        name,
        titles: flags & QSI_TITLES != 0,
        row_numbers: flags & QSI_ROW_NUMS != 0,
        disable_refresh: flags & QSI_DISABLE_REFRESH != 0,
        async_refresh: flags & QSI_ASYNC != 0,
        first_refresh_pending: flags & QSI_NEW_ASYNC != 0,
        auto_refresh: flags & QSI_AUTO_REFRESH != 0,
        shrink: flags & QSI_SHRINK != 0,
        fill: flags & QSI_FILL != 0,
        auto_format: flags & QSI_AUTO_FORMAT != 0,
        save_data: flags & QSI_SAVE_DATA != 0,
        disable_edit: flags & QSI_DISABLE_EDIT != 0,
        overwrite: flags & QSI_OVERWRITE != 0,
        auto_format_index: binary::read_u16_le_at(data, 4).ok()?,
        auto_format_number: attributes & QSI_ATR_NUM != 0,
        auto_format_font: attributes & QSI_ATR_FNT != 0,
        auto_format_alignment: attributes & QSI_ATR_ALC != 0,
        auto_format_border: attributes & QSI_ATR_BDR != 0,
        auto_format_pattern: attributes & QSI_ATR_PAT != 0,
        auto_format_protection: attributes & QSI_ATR_PROT != 0,
        ..QueryTable::default()
    })
}

pub(crate) fn parse_db_query(build: &mut QueryTableBuild, data: &[u8]) -> Option<()> {
    if data.len() < 12 {
        return None;
    }
    let flags = binary::read_u16_le_at(data, 0).ok()?;
    build.table.source = QuerySource::from_dbt(flags & DBQUERY_DBT_MASK);
    build.table.save_password = flags & DBQUERY_SAVE_PWD != 0;
    build.table.tables_only_html = flags & DBQUERY_TABLES_ONLY_HTML != 0;
    build.remaining_query = string_count(data, 4, flags & (DBQUERY_SQL | DBQUERY_WEB) != 0)?;
    build.remaining_odbc_conn = string_count(data, 10, flags & DBQUERY_ODBC_CONN != 0)?;
    build.remaining_web_post = string_count(data, 6, flags & DBQUERY_WEB != 0)?;
    build.remaining_sql_sav = string_count(data, 8, flags & DBQUERY_SQL_SAV != 0)?;
    build.dbquery_seen = true;
    Some(())
}

/// A declared SXString chunk count; negative counts are clamped to zero.
pub(crate) fn string_count(data: &[u8], offset: usize, present: bool) -> Option<u16> {
    let count = binary::read_i16_le(data, offset).ok()?;
    Some(if present { count.max(0) as u16 } else { 0 })
}

pub(crate) fn parse_param_qry(build: &mut QueryTableBuild, data: &[u8]) -> Option<()> {
    if data.len() < 8 {
        return None;
    }
    let sql_type = binary::read_u16_le_at(data, 0).ok()?;
    let flags = binary::read_u16_le_at(data, 2).ok()?;
    let value_type = binary::read_u16_le_at(data, 4).ok()?;
    let pbt = flags & PARAMQRY_PBT_MASK;
    let parameter_type = match pbt {
        0 => QueryParameterType::Prompt,
        1 => QueryParameterType::Value,
        2 => QueryParameterType::CellReference,
        other => QueryParameterType::Unknown(other),
    };
    // For prompt parameters the trailing field is an SXString prompt
    // followed by an unused byte; other value encodings stay uninterpreted.
    let prompt = if pbt == 0 && data.len() > 8 {
        parse_string_record(&data[8..], &Encoding::Utf16Le).ok()
    } else {
        None
    };
    build.table.parameters.push(QueryParameter {
        name: build.pending_param_name.take().unwrap_or_default(),
        parameter_type,
        sql_type,
        non_default_name: flags & PARAMQRY_NON_DEFAULT_NAME != 0,
        value_type,
        prompt,
    });
    Some(())
}

pub(crate) fn parse_db_query_ext(build: &mut QueryTableBuild, data: &[u8]) -> Option<()> {
    if data.len() < 28 || binary::read_u16_le_at(data, 0).ok()? != DB_QUERY_EXT_RECORD_TYPE {
        return None;
    }
    let flags = binary::read_u16_le_at(data, 6).ok()?;
    let trailing = binary::read_u16_le_at(data, 10).ok()?;
    let parameter_count = usize::from(binary::read_u16_le_at(data, 26).ok()?);
    let parameter_end = 28usize.checked_add(parameter_count.checked_mul(2)?)?;
    if data.len() < parameter_end {
        return None;
    }
    let future_count = usize::from(binary::read_u16_le_at(data, 20).ok()?);
    let future_end = parameter_end.checked_add(future_count)?;
    // The DBQueryExt DataSourceType supersedes the 3-bit DbQuery dbt.
    build.table.source = QuerySource::from_dbt(binary::read_u16_le_at(data, 4).ok()?);
    build.table.maintain_connection = flags & DBEXT_MAINTAIN != 0;
    build.table.new_query = flags & DBEXT_NEW_QUERY != 0;
    build.table.import_xml_source = flags & DBEXT_IMPORT_XML_SOURCE != 0;
    build.table.sharepoint_list_source = flags & DBEXT_SP_LIST_SRC != 0;
    build.table.sharepoint_list_reinit = flags & DBEXT_SP_LIST_REINIT_CACHE != 0;
    build.table.source_is_xml = flags & DBEXT_SRC_IS_XML != 0;
    build.table.connection_flags = binary::read_u16_le_at(data, 8).ok()?;
    build.table.edited_version = data[12];
    build.table.refreshed_version = data[13];
    build.table.refreshable_min_version = data[14];
    build.table.refresh_interval = binary::read_u16_le_at(data, 22).ok()?;
    build.table.html_formatting = match binary::read_u16_le_at(data, 24).ok()? {
        0x0001 => HtmlFormatting::None,
        0x0002 => HtmlFormatting::RichText,
        0x0003 => HtmlFormatting::Full,
        other => HtmlFormatting::Unknown(other),
    };
    build.table.parameter_flags = data[28..parameter_end]
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect();
    build.table.future_bytes = data[parameter_end..future_end.min(data.len())].to_vec();
    build.ext_context = if trailing & DBEXT_TABLE_NAMES != 0 {
        ExtContext::TableNames
    } else {
        ExtContext::None
    };
    Some(())
}

pub(crate) fn parse_txt_qry(data: &[u8]) -> Option<TextQuery> {
    if data.len() < 22 || binary::read_u16_le_at(data, 0).ok()? != TXT_QRY_RECORD_TYPE {
        return None;
    }
    let flags = binary::read_u16_le_at(data, 4).ok()?;
    if flags & TXT_FILE == 0 {
        return None;
    }
    let delimiters = data[12];
    let field_count = binary::read_i32_le(data, 16).ok()?;
    if field_count <= 0 || field_count as usize > MAX_TXT_FIELDS {
        return None;
    }
    let fields_end = 22usize.checked_add((field_count as usize).checked_mul(8)?)?;
    if data.len() < fields_end {
        return None;
    }
    let fields = data[22..fields_end]
        .chunks_exact(8)
        .map(|chunk| TextField {
            format: match u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]) {
                0 => TextFieldFormat::General,
                1 => TextFieldFormat::Text,
                2 => TextFieldFormat::DateMdy,
                3 => TextFieldFormat::DateDmy,
                4 => TextFieldFormat::DateYmd,
                5 => TextFieldFormat::DateMyd,
                6 => TextFieldFormat::DateDym,
                7 => TextFieldFormat::DateYdm,
                8 => TextFieldFormat::Skip,
                other => TextFieldFormat::Unknown(other),
            },
            start: i32::from_le_bytes([chunk[4], chunk[5], chunk[6], chunk[7]]),
        })
        .collect();
    let file = parse_string_record(&data[fields_end..], &Encoding::Utf16Le).ok()?;
    Some(TextQuery {
        delimited: flags & TXT_DELIMITED != 0,
        codepage: match (flags >> TXT_CPID_SHIFT) & TXT_CPID_MASK {
            0 => TextCodePage::Macintosh,
            1 => TextCodePage::WindowsAnsi,
            2 => TextCodePage::MsDos,
            other => TextCodePage::Unknown(other),
        },
        new_codepage: (flags >> 5) & 0x03FF,
        use_new_codepage: flags & TXT_USE_NEW_CPID != 0,
        prompt_for_file: flags & TXT_PROMPT_FOR_FILE != 0,
        row_start_at: binary::read_i32_le(data, 8).ok()?,
        tab: delimiters & TXT_DELIM_TAB != 0,
        space: delimiters & TXT_DELIM_SPACE != 0,
        comma: delimiters & TXT_DELIM_COMMA != 0,
        semicolon: delimiters & TXT_DELIM_SEMICOLON != 0,
        custom_delimiter: if delimiters & TXT_DELIM_CUSTOM != 0 {
            char::from_u32(u32::from(binary::read_u16_le_at(data, 13).ok()?))
        } else {
            None
        },
        consecutive: delimiters & TXT_DELIM_CONSECUTIVE != 0,
        text_delimiter: match (delimiters >> TXT_TEXT_DELIM_SHIFT) & TXT_TEXT_DELIM_MASK {
            0 => TextDelimiter::QuotationMark,
            1 => TextDelimiter::Apostrophe,
            2 => TextDelimiter::None,
            other => TextDelimiter::Unknown(other),
        },
        decimal_separator: char::from(data[20]),
        thousands_separator: char::from(data[21]),
        fields,
        file,
        connection_string: String::new(),
    })
}
