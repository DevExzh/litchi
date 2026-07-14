//! Spreadsheet-wide formula calculation settings.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::num::NonZeroUsize;

const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum IterationStatus {
    Enable,
    Disable,
}

impl IterationStatus {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "enable" => Ok(Self::Enable),
            "disable" => Ok(Self::Disable),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table:iteration status '{value}'"
            ))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Enable => "enable",
            Self::Disable => "disable",
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CalculationNullDate {
    /// Whether `table:value-type="date"` was explicitly present.
    pub value_type_date: bool,
    /// XML Schema date lexical value, preserved without timezone normalization.
    pub date_value: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CalculationIteration {
    pub status: Option<IterationStatus>,
    pub steps: Option<NonZeroUsize>,
    /// XML Schema double lexical value, preserving `INF`, `-INF`, and `NaN`.
    pub maximum_difference: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CalculationSettings {
    pub case_sensitive: Option<bool>,
    pub precision_as_shown: Option<bool>,
    pub search_criteria_must_apply_to_whole_cell: Option<bool>,
    pub automatic_find_labels: Option<bool>,
    pub use_regular_expressions: Option<bool>,
    pub use_wildcards: Option<bool>,
    pub null_year: Option<NonZeroUsize>,
    pub null_date: Option<CalculationNullDate>,
    pub iteration: Option<CalculationIteration>,
}

impl CalculationSettings {
    pub fn validate(&self) -> Result<()> {
        if let Some(value) = self
            .null_date
            .as_ref()
            .and_then(|null_date| null_date.date_value.as_deref())
            && !is_xsd_date(value)
        {
            return Err(Error::InvalidFormat(format!(
                "invalid calculation null date '{value}'"
            )));
        }
        if let Some(value) = self
            .iteration
            .as_ref()
            .and_then(|iteration| iteration.maximum_difference.as_deref())
            && !is_xsd_double(value)
        {
            return Err(Error::InvalidFormat(format!(
                "invalid iteration maximum difference '{value}'"
            )));
        }
        Ok(())
    }
}

pub(crate) fn write_calculation_settings(
    out: &mut String,
    settings: Option<&CalculationSettings>,
) -> Result<()> {
    let Some(settings) = settings else {
        return Ok(());
    };
    settings.validate()?;
    out.push_str("<table:calculation-settings");
    write_optional_bool(out, "table:case-sensitive", settings.case_sensitive);
    write_optional_bool(out, "table:precision-as-shown", settings.precision_as_shown);
    write_optional_bool(
        out,
        "table:search-criteria-must-apply-to-whole-cell",
        settings.search_criteria_must_apply_to_whole_cell,
    );
    write_optional_bool(
        out,
        "table:automatic-find-labels",
        settings.automatic_find_labels,
    );
    write_optional_bool(
        out,
        "table:use-regular-expressions",
        settings.use_regular_expressions,
    );
    write_optional_bool(out, "table:use-wildcards", settings.use_wildcards);
    if let Some(year) = settings.null_year {
        write_attribute(out, "table:null-year", &year.get().to_string());
    }
    if settings.null_date.is_none() && settings.iteration.is_none() {
        out.push_str("/>");
        return Ok(());
    }
    out.push('>');
    if let Some(null_date) = &settings.null_date {
        out.push_str("<table:null-date");
        if null_date.value_type_date {
            write_attribute(out, "table:value-type", "date");
        }
        if let Some(value) = &null_date.date_value {
            write_attribute(out, "table:date-value", value);
        }
        out.push_str("/>");
    }
    if let Some(iteration) = &settings.iteration {
        out.push_str("<table:iteration");
        if let Some(status) = iteration.status {
            write_attribute(out, "table:status", status.as_str());
        }
        if let Some(steps) = iteration.steps {
            write_attribute(out, "table:steps", &steps.get().to_string());
        }
        if let Some(value) = &iteration.maximum_difference {
            write_attribute(out, "table:maximum-difference", value);
        }
        out.push_str("/>");
    }
    out.push_str("</table:calculation-settings>");
    Ok(())
}

pub(crate) fn parse_calculation_settings(xml: &str) -> Result<Option<CalculationSettings>> {
    let mut reader = NsReader::from_str(xml);
    let mut buf = Vec::new();
    let mut current = None;
    let mut result = None;
    let mut child_open = false;
    let mut depth = 0usize;
    let mut spreadsheet_depth = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let is_table = is_namespace(&namespace);
        let is_start = matches!(&event, Event::Start(_));
        let is_end = matches!(&event, Event::End(_));
        if let Event::Start(element) = &event
            && is_namespace_uri(&namespace, OFFICE_NAMESPACE)
            && element.local_name().as_ref() == b"spreadsheet"
        {
            spreadsheet_depth = Some(depth);
        }
        let is_spreadsheet_child = spreadsheet_depth.is_some_and(|value| depth == value + 1);
        match event {
            Event::Start(element)
                if is_table && element.local_name().as_ref() == b"calculation-settings" =>
            {
                if !is_spreadsheet_child {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings must be a direct office:spreadsheet child"
                            .to_string(),
                    ));
                }
                if current.is_some() || result.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate table:calculation-settings".to_string(),
                    ));
                }
                current = Some(parse_settings_attributes(
                    reader.resolver(),
                    reader.decoder(),
                    &element,
                )?);
            },
            Event::Empty(element)
                if is_table && element.local_name().as_ref() == b"calculation-settings" =>
            {
                if !is_spreadsheet_child {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings must be a direct office:spreadsheet child"
                            .to_string(),
                    ));
                }
                if current.is_some() || result.is_some() {
                    return Err(Error::InvalidFormat(
                        "duplicate table:calculation-settings".to_string(),
                    ));
                }
                result = Some(parse_settings_attributes(
                    reader.resolver(),
                    reader.decoder(),
                    &element,
                )?);
            },
            Event::Start(element) if current.is_some() => {
                if child_open {
                    return Err(Error::InvalidFormat(
                        "calculation setting children must be empty".to_string(),
                    ));
                }
                parse_settings_child(
                    current.as_mut().expect("settings were checked"),
                    reader.resolver(),
                    reader.decoder(),
                    is_table,
                    &element,
                )?;
                child_open = true;
            },
            Event::Empty(element) if current.is_some() => {
                if child_open {
                    return Err(Error::InvalidFormat(
                        "calculation setting children must be empty".to_string(),
                    ));
                }
                parse_settings_child(
                    current.as_mut().expect("settings were checked"),
                    reader.resolver(),
                    reader.decoder(),
                    is_table,
                    &element,
                )?;
            },
            Event::End(element) if current.is_some() => {
                if child_open {
                    child_open = false;
                } else if is_table && element.local_name().as_ref() == b"calculation-settings" {
                    result = current.take();
                }
            },
            Event::Text(text) if current.is_some() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid calculation settings text: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings cannot contain text".to_string(),
                    ));
                }
            },
            Event::CData(text) if current.is_some() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid calculation settings CDATA: {error}"))
                })?;
                if !value.trim().is_empty() {
                    return Err(Error::InvalidFormat(
                        "table:calculation-settings cannot contain CDATA".to_string(),
                    ));
                }
            },
            Event::GeneralRef(_) if current.is_some() => {
                return Err(Error::InvalidFormat(
                    "table:calculation-settings cannot contain entity references".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        if is_start {
            depth = depth.saturating_add(1);
        } else if is_end {
            depth = depth.saturating_sub(1);
        }
        buf.clear();
    }
    if current.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated table:calculation-settings".to_string(),
        ));
    }
    if let Some(settings) = &result {
        settings.validate()?;
    }
    Ok(result)
}

fn parse_settings_attributes(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
) -> Result<CalculationSettings> {
    Ok(CalculationSettings {
        case_sensitive: optional_bool(resolver, decoder, element, b"case-sensitive")?,
        precision_as_shown: optional_bool(resolver, decoder, element, b"precision-as-shown")?,
        search_criteria_must_apply_to_whole_cell: optional_bool(
            resolver,
            decoder,
            element,
            b"search-criteria-must-apply-to-whole-cell",
        )?,
        automatic_find_labels: optional_bool(resolver, decoder, element, b"automatic-find-labels")?,
        use_regular_expressions: optional_bool(
            resolver,
            decoder,
            element,
            b"use-regular-expressions",
        )?,
        use_wildcards: optional_bool(resolver, decoder, element, b"use-wildcards")?,
        null_year: optional_positive(resolver, decoder, element, b"null-year")?,
        null_date: None,
        iteration: None,
    })
}

fn parse_settings_child(
    settings: &mut CalculationSettings,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    is_table: bool,
    element: &BytesStart<'_>,
) -> Result<()> {
    if !is_table {
        return Err(Error::InvalidFormat(
            "unsupported calculation settings child".to_string(),
        ));
    }
    match element.local_name().as_ref() {
        b"null-date" => {
            if settings.null_date.is_some() || settings.iteration.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate or out-of-order table:null-date".to_string(),
                ));
            }
            let value_type = optional_attribute(resolver, decoder, element, b"value-type")?;
            if value_type.as_deref().is_some_and(|value| value != "date") {
                return Err(Error::InvalidFormat(
                    "table:null-date value type must be 'date'".to_string(),
                ));
            }
            settings.null_date = Some(CalculationNullDate {
                value_type_date: value_type.is_some(),
                date_value: optional_attribute(resolver, decoder, element, b"date-value")?,
            });
        },
        b"iteration" => {
            if settings.iteration.is_some() {
                return Err(Error::InvalidFormat(
                    "duplicate table:iteration".to_string(),
                ));
            }
            settings.iteration = Some(CalculationIteration {
                status: optional_attribute(resolver, decoder, element, b"status")?
                    .map(|value| IterationStatus::parse(&value))
                    .transpose()?,
                steps: optional_positive(resolver, decoder, element, b"steps")?,
                maximum_difference: optional_attribute(
                    resolver,
                    decoder,
                    element,
                    b"maximum-difference",
                )?,
            });
        },
        _ => {
            return Err(Error::InvalidFormat(
                "unsupported calculation settings child".to_string(),
            ));
        },
    }
    Ok(())
}

fn optional_bool(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<bool>> {
    optional_attribute(resolver, decoder, element, name)?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid Boolean value '{value}'"
            ))),
        })
        .transpose()
}

fn optional_positive(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<NonZeroUsize>> {
    optional_attribute(resolver, decoder, element, name)?
        .map(|value| {
            value
                .parse::<usize>()
                .ok()
                .and_then(NonZeroUsize::new)
                .ok_or_else(|| Error::InvalidFormat(format!("invalid positive integer '{value}'")))
        })
        .transpose()
}

fn optional_attribute(
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if is_namespace(&namespace) && local.as_ref() == name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")));
        }
    }
    Ok(None)
}

fn is_namespace(namespace: &ResolveResult<'_>) -> bool {
    is_namespace_uri(namespace, TABLE_NAMESPACE)
}

fn is_namespace_uri(namespace: &ResolveResult<'_>, uri: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == uri)
}

fn write_optional_bool(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        write_attribute(out, name, if value { "true" } else { "false" });
    }
}

fn write_attribute(out: &mut String, name: &str, value: &str) {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    out.push_str(&escape_xml(value));
    out.push('"');
}

fn is_xsd_double(value: &str) -> bool {
    if matches!(value, "INF" | "-INF" | "NaN") {
        return true;
    }
    let bytes = value.as_bytes();
    let mut index = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let integer_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    let has_integer = index > integer_start;
    let mut has_fraction = false;
    if bytes.get(index) == Some(&b'.') {
        index += 1;
        let start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        has_fraction = index > start;
    }
    if !has_integer && !has_fraction {
        return false;
    }
    if matches!(bytes.get(index), Some(b'e' | b'E')) {
        index += 1;
        if matches!(bytes.get(index), Some(b'+' | b'-')) {
            index += 1;
        }
        let start = index;
        while bytes.get(index).is_some_and(u8::is_ascii_digit) {
            index += 1;
        }
        if index == start {
            return false;
        }
    }
    index == bytes.len()
}

fn is_xsd_date(value: &str) -> bool {
    if !value.is_ascii() {
        return false;
    }
    let date = if let Some(date) = value.strip_suffix('Z') {
        date
    } else if value.len() >= 6 {
        let split = value.len() - 6;
        let suffix = &value[split..];
        let timezone = matches!(suffix.as_bytes().first(), Some(b'+' | b'-'))
            && suffix.as_bytes().get(3) == Some(&b':')
            && suffix[1..3].bytes().all(|byte| byte.is_ascii_digit())
            && suffix[4..6].bytes().all(|byte| byte.is_ascii_digit())
            && suffix[1..3].parse::<u8>().is_ok_and(|hour| hour <= 14)
            && suffix[4..6].parse::<u8>().is_ok_and(|minute| minute <= 59)
            && (suffix[1..3] != *"14" || suffix[4..6] == *"00");
        if timezone { &value[..split] } else { value }
    } else {
        value
    };
    let date = date.strip_prefix('-').unwrap_or(date);
    let Some((year, rest)) = date.split_once('-') else {
        return false;
    };
    let Some((month, day)) = rest.split_once('-') else {
        return false;
    };
    if year.len() < 4
        || !year.bytes().all(|byte| byte.is_ascii_digit())
        || year.bytes().all(|byte| byte == b'0')
        || (year.len() > 4 && year.starts_with('0'))
        || month.len() != 2
        || day.len() != 2
    {
        return false;
    }
    let (Ok(month), Ok(day)) = (month.parse::<u8>(), day.parse::<u8>()) else {
        return false;
    };
    let leap =
        decimal_mod(year, 4) == 0 && (decimal_mod(year, 100) != 0 || decimal_mod(year, 400) == 0);
    let days = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => return false,
    };
    (1..=days).contains(&day)
}

fn decimal_mod(value: &str, modulus: u16) -> u16 {
    value.bytes().fold(0, |remainder, byte| {
        (remainder * 10 + u16::from(byte - b'0')) % modulus
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_lossless_lexical_forms() {
        for value in ["0", ".5", "5.", "-1.25E-3", "INF", "-INF", "NaN"] {
            assert!(is_xsd_double(value), "{value}");
        }
        for value in ["", ".", "inf", "1E", "1 2"] {
            assert!(!is_xsd_double(value), "{value}");
        }
        for value in ["1899-12-30", "2026-07-14Z", "2026-07-14+08:00"] {
            assert!(is_xsd_date(value), "{value}");
        }
        for value in [
            "2026-07-14+14:01",
            "0000-01-01",
            "02026-01-01",
            "2026-02-29",
        ] {
            assert!(!is_xsd_date(value), "{value}");
        }
        assert!(is_xsd_date("2024-02-29"));
    }

    #[test]
    fn writes_nested_settings_in_schema_order() {
        let settings = CalculationSettings {
            case_sensitive: Some(true),
            null_year: NonZeroUsize::new(1930),
            null_date: Some(CalculationNullDate {
                value_type_date: true,
                date_value: Some("1899-12-30Z".to_string()),
            }),
            iteration: Some(CalculationIteration {
                status: Some(IterationStatus::Enable),
                steps: NonZeroUsize::new(100),
                maximum_difference: Some("1E-6".to_string()),
            }),
            ..CalculationSettings::default()
        };
        let mut xml = String::new();
        write_calculation_settings(&mut xml, Some(&settings)).unwrap();
        assert!(xml.find("<table:null-date").unwrap() < xml.find("<table:iteration").unwrap());
    }

    #[test]
    fn parses_all_settings_with_namespace_aliases() {
        let xml = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:calculation-settings t:case-sensitive="1"
            t:precision-as-shown="false" t:search-criteria-must-apply-to-whole-cell="true"
            t:automatic-find-labels="false" t:use-regular-expressions="true"
            t:use-wildcards="false" t:null-year="1930">
            <t:null-date t:value-type="date" t:date-value="1899-12-30+08:00"></t:null-date>
            <t:iteration t:status="enable" t:steps="100" t:maximum-difference="NaN"/>
          </t:calculation-settings></o:spreadsheet></o:body>
        </o:document-content>"#;
        let settings = parse_calculation_settings(xml).unwrap().unwrap();
        assert_eq!(settings.case_sensitive, Some(true));
        assert_eq!(settings.null_year.unwrap().get(), 1930);
        assert_eq!(
            settings.null_date.unwrap().date_value.as_deref(),
            Some("1899-12-30+08:00")
        );
        let iteration = settings.iteration.unwrap();
        assert_eq!(iteration.status, Some(IterationStatus::Enable));
        assert_eq!(iteration.steps.unwrap().get(), 100);
        assert_eq!(iteration.maximum_difference.as_deref(), Some("NaN"));
    }

    #[test]
    fn rejects_invalid_or_out_of_order_settings() {
        let invalid = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:calculation-settings t:null-year="0">
            <t:iteration t:steps="0"/><t:null-date t:date-value="not-a-date"/>
          </t:calculation-settings></o:spreadsheet></o:body>
        </o:document-content>"#;
        assert!(parse_calculation_settings(invalid).is_err());
        let out_of_order = invalid
            .replace("t:null-year=\"0\"", "t:null-year=\"1930\"")
            .replace("t:steps=\"0\"", "t:steps=\"1\"")
            .replace("not-a-date", "1899-12-30");
        assert!(parse_calculation_settings(&out_of_order).is_err());
        let duplicate = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:calculation-settings/><t:calculation-settings/>
          </o:spreadsheet></o:body></o:document-content>"#;
        assert!(parse_calculation_settings(duplicate).is_err());
        let nested = r#"<o:document-content
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <o:body><o:spreadsheet><t:table t:name="Bad"><t:table-row><t:table-cell>
            <t:calculation-settings/></t:table-cell></t:table-row></t:table>
          </o:spreadsheet></o:body></o:document-content>"#;
        assert!(parse_calculation_settings(nested).is_err());
    }
}
