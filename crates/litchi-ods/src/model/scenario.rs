//! Bounded inspection of ODF spreadsheet scenario declarations.

use core::fmt;
use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::sync::Arc;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_INPUT_BYTES: usize = 256 * 1024 * 1024;
const MAX_SCENARIOS: usize = 65_536;
const MAX_RANGES: usize = 65_536;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_RANGE_LIST_BYTES: usize = 1024 * 1024;
const MAX_DEPTH: usize = 1_024;

/// A scenario metadata inspection result.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors produced while inspecting inert scenario metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A configured or hard resource limit was exceeded.
    ResourceLimit {
        /// The bounded resource.
        resource: &'static str,
        /// The observed or configured value.
        actual: usize,
        /// The maximum accepted value.
        maximum: usize,
    },
    /// The XML stream could not be decoded.
    InvalidXml(String),
    /// The document has invalid scenario structure or content.
    InvalidStructure(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ResourceLimit {
                resource,
                actual,
                maximum,
            } => write!(
                formatter,
                "{resource} limit exceeded: observed {actual}, maximum {maximum}"
            ),
            Self::InvalidXml(message) => write!(formatter, "invalid XML: {message}"),
            Self::InvalidStructure(message) => formatter.write_str(message),
        }
    }
}

impl std::error::Error for Error {}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    input_bytes: usize,
    scenarios: usize,
    ranges: usize,
    text_bytes: usize,
    range_list_bytes: usize,
    depth: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            input_bytes: MAX_INPUT_BYTES,
            scenarios: MAX_SCENARIOS,
            ranges: MAX_RANGES,
            text_bytes: MAX_TEXT_BYTES,
            range_list_bytes: MAX_RANGE_LIST_BYTES,
            depth: MAX_DEPTH,
        }
    }
}

impl Limits {
    #[must_use]
    pub const fn with_input_bytes(mut self, value: usize) -> Self {
        self.input_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_scenarios(mut self, value: usize) -> Self {
        self.scenarios = value;
        self
    }

    #[must_use]
    pub const fn with_ranges(mut self, value: usize) -> Self {
        self.ranges = value;
        self
    }

    #[must_use]
    pub const fn with_text_bytes(mut self, value: usize) -> Self {
        self.text_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_range_list_bytes(mut self, value: usize) -> Self {
        self.range_list_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_depth(mut self, value: usize) -> Self {
        self.depth = value;
        self
    }

    fn validate(self) -> Result<Self> {
        for (name, value, ceiling) in [
            ("input bytes", self.input_bytes, MAX_INPUT_BYTES),
            ("scenarios", self.scenarios, MAX_SCENARIOS),
            ("ranges", self.ranges, MAX_RANGES),
            ("text bytes", self.text_bytes, MAX_TEXT_BYTES),
            (
                "range-list bytes",
                self.range_list_bytes,
                MAX_RANGE_LIST_BYTES,
            ),
            ("XML depth", self.depth, MAX_DEPTH),
        ] {
            if value > ceiling {
                return Err(Error::ResourceLimit {
                    resource: name,
                    actual: value,
                    maximum: ceiling,
                });
            }
        }
        Ok(self)
    }
}

/// Whether a required scenario is active.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum State {
    Active,
    Inactive,
}

impl From<bool> for State {
    fn from(value: bool) -> Self {
        if value { Self::Active } else { Self::Inactive }
    }
}

/// A scenario setting that distinguishes absence from either boolean value.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum OptionalSetting {
    #[default]
    Unspecified,
    Enabled,
    Disabled,
}

impl From<Option<bool>> for OptionalSetting {
    fn from(value: Option<bool>) -> Self {
        match value {
            Some(true) => Self::Enabled,
            Some(false) => Self::Disabled,
            None => Self::Unspecified,
        }
    }
}

/// A checked ODF `#RRGGBB` color.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RgbColor {
    red: u8,
    green: u8,
    blue: u8,
}

impl RgbColor {
    /// Parses an ODF `#RRGGBB` color.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is not exactly six hexadecimal digits
    /// preceded by `#`.
    pub fn from_hex(value: &str) -> Result<Self> {
        if value.len() != 7
            || !value.starts_with('#')
            || !value.as_bytes()[1..].iter().all(u8::is_ascii_hexdigit)
        {
            return Err(invalid("table:border-color must be an RGB color"));
        }
        let component = |range: std::ops::Range<usize>| {
            u8::from_str_radix(&value[range], 16)
                .map_err(|_parse_error| invalid("table:border-color must be an RGB color"))
        };
        Ok(Self {
            red: component(1..3)?,
            green: component(3..5)?,
            blue: component(5..7)?,
        })
    }

    #[must_use]
    pub const fn red(self) -> u8 {
        self.red
    }

    #[must_use]
    pub const fn green(self) -> u8 {
        self.green
    }

    #[must_use]
    pub const fn blue(self) -> u8 {
        self.blue
    }
}

impl fmt::Display for RgbColor {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "#{:02X}{:02X}{:02X}",
            self.red, self.green, self.blue
        )
    }
}

/// One checked ODF cell-range address retained in source form.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RangeAddress(String);

impl RangeAddress {
    /// Creates a checked single range address.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is empty, malformed, oversized, or
    /// contains more than one range address.
    pub fn new(source: impl Into<String>) -> Result<Self> {
        let value = source.into();
        let limits = Limits::default().with_ranges(1);
        preflight_ranges(&value, limits)?;
        let parsed = crate::model::structure::split_cell_range_addresses(&value)
            .map_err(|error| invalid(format!("invalid scenario range address: {error}")))?;
        if parsed.len() != 1 || parsed.first() != Some(&value) {
            return Err(invalid("expected exactly one scenario range address"));
        }
        Ok(Self(value))
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for RangeAddress {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for RangeAddress {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

/// Typed attributes of one empty `table:scenario` element.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Scenario {
    sheet: String,
    ranges: Vec<RangeAddress>,
    state: State,
    display_border: OptionalSetting,
    border_color: Option<RgbColor>,
    copy_back: OptionalSetting,
    copy_styles: OptionalSetting,
    copy_formulas: OptionalSetting,
    comment: Option<String>,
    protected: OptionalSetting,
}

impl Scenario {
    /// Creates a detached, inert scenario descriptor.
    ///
    /// # Errors
    ///
    /// Returns an error when the sheet name is empty or invalid, or when the
    /// range list is empty or exceeds the hard range-count ceiling.
    pub fn new(
        sheet_name: impl Into<String>,
        ranges: Vec<RangeAddress>,
        state: State,
    ) -> Result<Self> {
        let sheet = sheet_name.into();
        if sheet.is_empty() || sheet.len() > MAX_TEXT_BYTES || !xml_text_is_valid(&sheet) {
            return Err(invalid("invalid scenario sheet name"));
        }
        if ranges.is_empty() {
            return Err(invalid("scenario range list must not be empty"));
        }
        if ranges.len() > MAX_RANGES {
            return Err(Error::ResourceLimit {
                resource: "ranges",
                actual: ranges.len(),
                maximum: MAX_RANGES,
            });
        }
        Ok(Self {
            sheet,
            ranges,
            state,
            display_border: OptionalSetting::Unspecified,
            border_color: None,
            copy_back: OptionalSetting::Unspecified,
            copy_styles: OptionalSetting::Unspecified,
            copy_formulas: OptionalSetting::Unspecified,
            comment: None,
            protected: OptionalSetting::Unspecified,
        })
    }

    #[must_use]
    pub fn sheet(&self) -> &str {
        &self.sheet
    }

    #[must_use]
    pub fn ranges(&self) -> &[RangeAddress] {
        &self.ranges
    }

    #[must_use]
    pub const fn state(&self) -> State {
        self.state
    }

    #[must_use]
    pub const fn is_active(&self) -> bool {
        matches!(self.state, State::Active)
    }

    #[must_use]
    pub const fn display_border(&self) -> OptionalSetting {
        self.display_border
    }

    #[must_use]
    pub const fn border_color(&self) -> Option<RgbColor> {
        self.border_color
    }

    #[must_use]
    pub const fn copy_back(&self) -> OptionalSetting {
        self.copy_back
    }

    #[must_use]
    pub const fn copy_styles(&self) -> OptionalSetting {
        self.copy_styles
    }

    #[must_use]
    pub const fn copy_formulas(&self) -> OptionalSetting {
        self.copy_formulas
    }

    #[must_use]
    pub fn comment(&self) -> Option<&str> {
        self.comment.as_deref()
    }

    #[must_use]
    pub const fn protected(&self) -> OptionalSetting {
        self.protected
    }
}

/// Immutable source-bound scenario inventory.
#[derive(Clone, Debug)]
pub struct Snapshot {
    content: Arc<str>,
    scenarios: Vec<Scenario>,
}

impl Snapshot {
    /// Parse the default-bounded scenario inventory without applying it.
    ///
    /// # Errors
    ///
    /// Returns an error when XML is malformed, violates the ODF scenario
    /// grammar, or exceeds a default resource limit.
    pub fn parse(content_xml: &str) -> Result<Self> {
        Self::parse_with(content_xml, Limits::default())
    }

    /// Parse the scenario inventory under caller-provided resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error when XML is malformed, violates the ODF scenario
    /// grammar, or exceeds `limits`.
    pub fn parse_with(content_xml: &str, requested_limits: Limits) -> Result<Self> {
        let limits = requested_limits.validate()?;
        if content_xml.len() > limits.input_bytes {
            return Err(invalid("content.xml exceeds the scenario input limit"));
        }
        let mut reader = NsReader::from_str(content_xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut spreadsheet_depth = None;
        let mut sheet: Option<(usize, String, bool)> = None;
        let mut scenario_depth = None;
        let mut scenarios = Vec::new();

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidXml(error.to_string()))?;
            match event {
                Event::Start(element) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("XML depth overflow"))?;
                    if depth > limits.depth {
                        return Err(invalid("scenario XML exceeds the nesting limit"));
                    }
                    if is(
                        &namespace,
                        element.local_name().as_ref(),
                        OFFICE,
                        b"spreadsheet",
                    ) {
                        spreadsheet_depth = Some(depth);
                    } else if spreadsheet_depth.is_some_and(|value| depth == value + 1)
                        && is(&namespace, element.local_name().as_ref(), TABLE, b"table")
                    {
                        sheet = Some((
                            depth,
                            required_attr(&element, &reader, b"name", limits.text_bytes)?,
                            false,
                        ));
                    } else if sheet.as_ref().is_some_and(|value| depth == value.0 + 1)
                        && is(
                            &namespace,
                            element.local_name().as_ref(),
                            TABLE,
                            b"scenario",
                        )
                    {
                        let (sheet_name, seen) = {
                            let Some(current) = sheet.as_mut() else {
                                return Err(invalid("scenario sheet parser state is missing"));
                            };
                            (current.1.clone(), &mut current.2)
                        };
                        if *seen {
                            return Err(invalid("a table may contain only one scenario"));
                        }
                        *seen = true;
                        scenarios.push(parse_scenario(&element, &reader, sheet_name, limits)?);
                        if scenarios.len() > limits.scenarios {
                            return Err(invalid("scenario count exceeds its limit"));
                        }
                        scenario_depth = Some(depth);
                    } else if scenario_depth.is_some() {
                        return Err(invalid("table:scenario must not contain child elements"));
                    }
                },
                Event::Empty(element) => {
                    let event_depth = depth + 1;
                    if scenario_depth.is_some() {
                        return Err(invalid("table:scenario must not contain child elements"));
                    } else if sheet
                        .as_ref()
                        .is_some_and(|value| event_depth == value.0 + 1)
                        && is(
                            &namespace,
                            element.local_name().as_ref(),
                            TABLE,
                            b"scenario",
                        )
                    {
                        let (sheet_name, seen) = {
                            let Some(current) = sheet.as_mut() else {
                                return Err(invalid("scenario sheet parser state is missing"));
                            };
                            (current.1.clone(), &mut current.2)
                        };
                        if *seen {
                            return Err(invalid("a table may contain only one scenario"));
                        }
                        *seen = true;
                        scenarios.push(parse_scenario(&element, &reader, sheet_name, limits)?);
                        if scenarios.len() > limits.scenarios {
                            return Err(invalid("scenario count exceeds its limit"));
                        }
                    }
                },
                Event::End(element) => {
                    if scenario_depth == Some(depth)
                        && is(
                            &namespace,
                            element.local_name().as_ref(),
                            TABLE,
                            b"scenario",
                        )
                    {
                        scenario_depth = None;
                    } else if sheet.as_ref().is_some_and(|value| depth == value.0)
                        && is(&namespace, element.local_name().as_ref(), TABLE, b"table")
                    {
                        sheet = None;
                    } else if spreadsheet_depth == Some(depth)
                        && is(
                            &namespace,
                            element.local_name().as_ref(),
                            OFFICE,
                            b"spreadsheet",
                        )
                    {
                        spreadsheet_depth = None;
                    }
                    depth = depth.saturating_sub(1);
                },
                Event::Text(text) if scenario_depth.is_some() => {
                    let value = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| invalid(format!("invalid scenario text: {error}")))?;
                    if !value.trim().is_empty() {
                        return Err(invalid("table:scenario must be empty"));
                    }
                },
                Event::CData(text) if scenario_depth.is_some() => {
                    let value = text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| invalid(format!("invalid scenario CDATA: {error}")))?;
                    if !value.trim().is_empty() {
                        return Err(invalid("table:scenario must be empty"));
                    }
                },
                Event::GeneralRef(_) if scenario_depth.is_some() => {
                    return Err(invalid("table:scenario must not contain entity references"));
                },
                Event::DocType(_) => return Err(invalid("DTD content is not accepted")),
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }
        if depth != 0 || scenario_depth.is_some() {
            return Err(invalid("unfinished scenario XML structure"));
        }
        Ok(Self {
            content: Arc::from(content_xml),
            scenarios,
        })
    }

    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.content
    }

    #[must_use]
    pub fn scenarios(&self) -> &[Scenario] {
        &self.scenarios
    }
}

fn parse_scenario(
    element: &quick_xml::events::BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    sheet: String,
    limits: Limits,
) -> Result<Scenario> {
    let ranges_value = required_attr(element, reader, b"scenario-ranges", limits.text_bytes)?;
    preflight_ranges(&ranges_value, limits)?;
    let ranges = crate::model::structure::split_cell_range_addresses(&ranges_value)
        .map_err(|error| invalid(format!("invalid scenario range list: {error}")))?;
    if ranges.is_empty() || ranges.len() > limits.ranges {
        return Err(invalid("scenario range list is empty or exceeds its limit"));
    }
    let active = parse_bool(&required_attr(
        element,
        reader,
        b"is-active",
        limits.text_bytes,
    )?)?;
    let optional_setting = |name| -> Result<OptionalSetting> {
        Ok(optional_attr(element, reader, name, limits.text_bytes)?
            .as_deref()
            .map(parse_bool)
            .transpose()?
            .into())
    };
    let border_color = optional_attr(element, reader, b"border-color", limits.text_bytes)?
        .as_deref()
        .map(RgbColor::from_hex)
        .transpose()?;
    let mut scenario = Scenario::new(
        sheet,
        ranges.into_iter().map(RangeAddress).collect(),
        active.into(),
    )?;
    scenario.display_border = optional_setting(b"display-border")?;
    scenario.border_color = border_color;
    scenario.copy_back = optional_setting(b"copy-back")?;
    scenario.copy_styles = optional_setting(b"copy-styles")?;
    scenario.copy_formulas = optional_setting(b"copy-formulas")?;
    scenario.comment = optional_attr(element, reader, b"comment", limits.text_bytes)?;
    scenario.protected = optional_setting(b"protected")?;
    Ok(scenario)
}

fn preflight_ranges(value: &str, limits: Limits) -> Result<()> {
    if value.len() > limits.range_list_bytes {
        return Err(invalid("scenario range list exceeds its byte limit"));
    }
    let mut ranges = 0usize;
    let mut token = false;
    let mut quoted = false;
    let mut characters = value.chars().peekable();
    while let Some(character) = characters.next() {
        if character == '\'' {
            token = true;
            if quoted && characters.peek() == Some(&'\'') {
                characters.next();
            } else {
                quoted = !quoted;
            }
        } else if character.is_whitespace() && !quoted {
            if token {
                ranges = ranges
                    .checked_add(1)
                    .ok_or_else(|| invalid("scenario range count overflows"))?;
                if ranges > limits.ranges {
                    return Err(invalid("scenario range count exceeds its limit"));
                }
                token = false;
            }
        } else {
            token = true;
        }
    }
    if quoted {
        return Err(invalid(
            "scenario range list contains an unterminated quoted table name",
        ));
    }
    if token {
        ranges = ranges
            .checked_add(1)
            .ok_or_else(|| invalid("scenario range count overflows"))?;
    }
    if ranges == 0 || ranges > limits.ranges {
        return Err(invalid(
            "scenario range list is empty or exceeds its count limit",
        ));
    }
    Ok(())
}

fn required_attr(
    element: &quick_xml::events::BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    local: &[u8],
    limit: usize,
) -> Result<String> {
    optional_attr(element, reader, local, limit)?.ok_or_else(|| {
        invalid(format!(
            "missing required table:{} attribute",
            String::from_utf8_lossy(local)
        ))
    })
}

fn optional_attr(
    element: &quick_xml::events::BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    local: &[u8],
    limit: usize,
) -> Result<Option<String>> {
    let mut value = None;
    for raw_attribute in element.attributes().with_checks(true) {
        let attribute = raw_attribute
            .map_err(|error| invalid(format!("invalid scenario attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == TABLE)
            && name.as_ref() == local
        {
            if value.is_some() {
                return Err(invalid("duplicate scenario attribute"));
            }
            let decoded = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid(format!("invalid scenario attribute value: {error}")))?
                .into_owned();
            if decoded.len() > limit || !xml_text_is_valid(&decoded) {
                return Err(invalid("invalid or oversized scenario attribute"));
            }
            value = Some(decoded);
        }
    }
    Ok(value)
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("invalid XML boolean '{value}'"))),
    }
}

fn xml_text_is_valid(value: &str) -> bool {
    !value.chars().any(|character| {
        matches!(
            character,
            '\u{0000}'..='\u{0008}' | '\u{000B}'..='\u{000C}' | '\u{000E}'..='\u{001F}'
        )
    })
}

fn is(namespace: &ResolveResult<'_>, local: &[u8], expected_ns: &[u8], expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected_ns)
        && local == expected
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidStructure(message.into())
}
