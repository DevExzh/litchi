//! Bounded inspection of ODF spreadsheet scenario declarations.

use core::fmt;
use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::sync::Arc;

mod transaction;

pub use transaction::{Commit, Edit, Patch};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_INPUT_BYTES: usize = 256 * 1024 * 1024;
const MAX_SCENARIOS: usize = 65_536;
const MAX_RANGES: usize = 65_536;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_RANGE_LIST_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_AGGREGATE_BYTES: usize = 16 * 1024 * 1024;
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
    aggregate_bytes: usize,
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
            aggregate_bytes: MAX_AGGREGATE_BYTES,
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

    /// Bound the aggregate authored scenario metadata retained by one snapshot.
    #[must_use]
    pub const fn with_aggregate_bytes(mut self, value: usize) -> Self {
        self.aggregate_bytes = value;
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
            (
                "aggregate scenario bytes",
                self.aggregate_bytes,
                MAX_AGGREGATE_BYTES,
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
    /// Creates a checked RGB color from its individual components.
    #[must_use]
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }

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
    /// Returns an error when `value` is empty or whitespace-only, malformed,
    /// oversized, or contains more than one range address.
    pub fn new(source: impl Into<String>) -> Result<Self> {
        let value = source.into();
        if value.trim().is_empty() {
            return Err(invalid("scenario range address must not be empty"));
        }
        validate_cell_range_address(&value)?;
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
        let scenario = Self {
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
        };
        scenario.validate()?;
        Ok(scenario)
    }

    /// Set whether the scenario border is displayed when this metadata is
    /// interpreted by an ODF consumer.
    #[must_use]
    pub const fn with_display_border(mut self, value: OptionalSetting) -> Self {
        self.display_border = value;
        self
    }

    /// Set or clear the optional scenario border color.
    #[must_use]
    pub const fn with_border_color(mut self, value: Option<RgbColor>) -> Self {
        self.border_color = value;
        self
    }

    /// Set the optional copy-back setting.
    #[must_use]
    pub const fn with_copy_back(mut self, value: OptionalSetting) -> Self {
        self.copy_back = value;
        self
    }

    /// Set the optional copy-styles setting.
    #[must_use]
    pub const fn with_copy_styles(mut self, value: OptionalSetting) -> Self {
        self.copy_styles = value;
        self
    }

    /// Set the optional copy-formulas setting.  This only records metadata;
    /// it never evaluates or copies formulas.
    #[must_use]
    pub const fn with_copy_formulas(mut self, value: OptionalSetting) -> Self {
        self.copy_formulas = value;
        self
    }

    /// Set or clear the optional scenario comment.
    #[must_use]
    pub fn with_comment(mut self, value: impl Into<String>) -> Self {
        self.comment = Some(value.into());
        self
    }

    /// Clear the optional scenario comment.
    #[must_use]
    pub fn without_comment(mut self) -> Self {
        self.comment = None;
        self
    }

    /// Set the optional protection setting.
    #[must_use]
    pub const fn with_protected(mut self, value: OptionalSetting) -> Self {
        self.protected = value;
        self
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if self.sheet.is_empty()
            || self.sheet.len() > MAX_TEXT_BYTES
            || !xml_text_is_valid(&self.sheet)
        {
            return Err(invalid("invalid scenario sheet name"));
        }
        if self.ranges.is_empty() {
            return Err(invalid("scenario range list must not be empty"));
        }
        if self.ranges.len() > MAX_RANGES {
            return Err(Error::ResourceLimit {
                resource: "ranges",
                actual: self.ranges.len(),
                maximum: MAX_RANGES,
            });
        }
        let mut range_bytes = 0usize;
        for range in &self.ranges {
            // Re-run the bounded range validation so detached descriptors
            // cannot smuggle an invalid or over-budget list into a writer.
            RangeAddress::new(range.as_str())?;
            range_bytes = range_bytes
                .checked_add(range.as_str().len())
                .and_then(|value| value.checked_add(if range_bytes == 0 { 0 } else { 1 }))
                .ok_or_else(|| invalid("scenario range list size overflows"))?;
        }
        if range_bytes > MAX_RANGE_LIST_BYTES {
            return Err(Error::ResourceLimit {
                resource: "range-list bytes",
                actual: range_bytes,
                maximum: MAX_RANGE_LIST_BYTES,
            });
        }
        if self
            .comment
            .as_deref()
            .is_some_and(|comment| comment.len() > MAX_TEXT_BYTES || !xml_text_is_valid(comment))
        {
            return Err(invalid("invalid or oversized scenario comment"));
        }
        Ok(())
    }

    pub(crate) fn aggregate_bytes(&self) -> Result<usize> {
        let mut total = self.sheet.len();
        for (index, range) in self.ranges.iter().enumerate() {
            total = total
                .checked_add(range.as_str().len())
                .and_then(|value| value.checked_add(usize::from(index != 0)))
                .ok_or_else(|| invalid("scenario aggregate size overflows"))?;
        }
        if let Some(comment) = self.comment.as_deref() {
            total = total
                .checked_add(comment.len())
                .ok_or_else(|| invalid("scenario aggregate size overflows"))?;
        }
        Ok(total)
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
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    content: Arc<str>,
    scenarios: Vec<Scenario>,
    aggregate_bytes: usize,
    limits: Limits,
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
        let mut sheet: Option<(usize, Option<String>, bool)> = None;
        let mut scenario_depth = None;
        let mut scenarios = Vec::new();
        let mut aggregate_bytes = 0usize;

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
                            optional_attr(&element, &reader, b"name", limits.text_bytes)?,
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
                        precharge_scenario_slot(&mut scenarios, limits)?;
                        let (sheet_name, seen) = {
                            let Some(current) = sheet.as_mut() else {
                                return Err(invalid("scenario sheet parser state is missing"));
                            };
                            (
                                current.1.clone().ok_or_else(|| {
                                    invalid("table:scenario requires a named worksheet")
                                })?,
                                &mut current.2,
                            )
                        };
                        if *seen {
                            return Err(invalid("a table may contain only one scenario"));
                        }
                        *seen = true;
                        let parsed = parse_scenario(&element, &reader, sheet_name, limits)?;
                        aggregate_bytes = add_aggregate_bytes(aggregate_bytes, &parsed, limits)?;
                        scenarios.push(parsed);
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
                        precharge_scenario_slot(&mut scenarios, limits)?;
                        let (sheet_name, seen) = {
                            let Some(current) = sheet.as_mut() else {
                                return Err(invalid("scenario sheet parser state is missing"));
                            };
                            (
                                current.1.clone().ok_or_else(|| {
                                    invalid("table:scenario requires a named worksheet")
                                })?,
                                &mut current.2,
                            )
                        };
                        if *seen {
                            return Err(invalid("a table may contain only one scenario"));
                        }
                        *seen = true;
                        let parsed = parse_scenario(&element, &reader, sheet_name, limits)?;
                        aggregate_bytes = add_aggregate_bytes(aggregate_bytes, &parsed, limits)?;
                        scenarios.push(parsed);
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
            aggregate_bytes,
            limits,
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

    /// Begin a source-checked, failure-atomic edit of scenario declarations.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit::new(self.clone())
    }

    /// Find one scenario by its exact worksheet name.
    #[must_use]
    pub fn for_sheet(&self, sheet_name: &str) -> Option<&Scenario> {
        let mut matches = self
            .scenarios
            .iter()
            .filter(|scenario| scenario.sheet() == sheet_name);
        let first = matches.next();
        first.filter(|_| matches.next().is_none())
    }
}

fn add_aggregate_bytes(current: usize, scenario: &Scenario, limits: Limits) -> Result<usize> {
    let next = current
        .checked_add(scenario.aggregate_bytes()?)
        .ok_or_else(|| invalid("scenario aggregate size overflows"))?;
    if next > limits.aggregate_bytes {
        return Err(Error::ResourceLimit {
            resource: "aggregate scenario bytes",
            actual: next,
            maximum: limits.aggregate_bytes,
        });
    }
    Ok(next)
}

fn precharge_scenario_slot(scenarios: &mut Vec<Scenario>, limits: Limits) -> Result<()> {
    let actual = scenarios
        .len()
        .checked_add(1)
        .ok_or_else(|| invalid("scenario count overflows"))?;
    if actual > limits.scenarios {
        return Err(Error::ResourceLimit {
            resource: "scenarios",
            actual,
            maximum: limits.scenarios,
        });
    }
    scenarios
        .try_reserve(1)
        .map_err(|_| invalid("scenario catalog allocation failed"))?;
    Ok(())
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
    for range in &ranges {
        validate_cell_range_address(range)?;
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
    let ranges = ranges.into_iter().map(RangeAddress).collect::<Vec<_>>();
    let mut scenario = Scenario::new(sheet, ranges, active.into())?;
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CellRangeKind {
    Cell,
    Column,
    Row,
}

/// Validate one ODF `cellRangeAddress` against the three alternatives in the
/// schema: cell or cell-range, column-range, and row-range.  This is kept
/// lexical and inert; it does not resolve sheet names or cell bounds.
fn validate_cell_range_address(value: &str) -> Result<()> {
    if value.is_empty() || value.trim() != value {
        return Err(invalid(
            "scenario range address has invalid surrounding whitespace",
        ));
    }
    let mut cursor = 0usize;
    parse_sheet_qualifier(value, &mut cursor)?;
    let kind = parse_first_range_endpoint(value, &mut cursor)?;
    match kind {
        CellRangeKind::Cell => {
            if cursor == value.len() {
                return Ok(());
            }
            expect_char(value, &mut cursor, ':')?;
            parse_sheet_qualifier(value, &mut cursor)?;
            parse_cell_endpoint(value, &mut cursor)?;
        },
        CellRangeKind::Column => {
            expect_char(value, &mut cursor, ':')?;
            parse_sheet_qualifier(value, &mut cursor)?;
            parse_column_endpoint(value, &mut cursor)?;
        },
        CellRangeKind::Row => {
            expect_char(value, &mut cursor, ':')?;
            parse_sheet_qualifier(value, &mut cursor)?;
            parse_row_endpoint(value, &mut cursor)?;
        },
    }
    if cursor != value.len() {
        return Err(invalid("scenario range address has trailing characters"));
    }
    Ok(())
}

fn parse_sheet_qualifier(value: &str, cursor: &mut usize) -> Result<()> {
    if char_at(value, *cursor) == Some('.') {
        *cursor += 1;
        return Ok(());
    }
    // The optional absolute marker is ambiguous with a legal unquoted sheet
    // name consisting of `$`.  When `$` is immediately followed by the
    // separator, the grammar must take it as the sheet name instead of
    // consuming it as the marker.
    if char_at(value, *cursor) == Some('$') && char_at(value, *cursor + 1) != Some('.') {
        *cursor += 1;
    }
    match char_at(value, *cursor) {
        Some('\'') => {
            *cursor += 1;
            let mut has_content = false;
            loop {
                let Some(character) = char_at(value, *cursor) else {
                    return Err(invalid(
                        "scenario range address has an unterminated sheet name",
                    ));
                };
                if character != '\'' {
                    *cursor += character.len_utf8();
                    has_content = true;
                    continue;
                }
                if char_at(value, *cursor + character.len_utf8()) == Some('\'') {
                    *cursor += character.len_utf8() * 2;
                    has_content = true;
                    continue;
                }
                *cursor += character.len_utf8();
                if !has_content {
                    return Err(invalid("scenario range address has an empty sheet name"));
                }
                break;
            }
        },
        Some(character) if character != ' ' && character != '\'' && character != '.' => {
            let mut has_content = false;
            while let Some(character) = char_at(value, *cursor) {
                if character == '.' {
                    break;
                }
                if character == ' ' || character == '\'' {
                    return Err(invalid("scenario range address has an invalid sheet name"));
                }
                *cursor += character.len_utf8();
                has_content = true;
            }
            if !has_content {
                return Err(invalid("scenario range address has an empty sheet name"));
            }
        },
        _ => return Err(invalid("scenario range address requires a sheet separator")),
    }
    expect_char(value, cursor, '.')
}

fn parse_first_range_endpoint(value: &str, cursor: &mut usize) -> Result<CellRangeKind> {
    consume_optional_dollar(value, cursor);
    match char_at(value, *cursor) {
        Some(character) if character.is_ascii_uppercase() => {
            parse_uppercase_run(value, cursor);
            if char_at(value, *cursor) == Some('$') {
                *cursor += 1;
                parse_digits(value, cursor)?;
                Ok(CellRangeKind::Cell)
            } else if char_at(value, *cursor).is_some_and(|value| value.is_ascii_digit()) {
                parse_digits(value, cursor)?;
                Ok(CellRangeKind::Cell)
            } else if char_at(value, *cursor) == Some(':') {
                Ok(CellRangeKind::Row)
            } else {
                Err(invalid(
                    "scenario range address has an invalid cell or row endpoint",
                ))
            }
        },
        Some(character) if character.is_ascii_digit() => {
            parse_digits(value, cursor)?;
            if char_at(value, *cursor) != Some(':') {
                return Err(invalid(
                    "scenario range address requires a column range endpoint",
                ));
            }
            Ok(CellRangeKind::Column)
        },
        _ => Err(invalid("scenario range address has an invalid endpoint")),
    }
}

fn parse_cell_endpoint(value: &str, cursor: &mut usize) -> Result<()> {
    consume_optional_dollar(value, cursor);
    if !char_at(value, *cursor).is_some_and(|value| value.is_ascii_uppercase()) {
        return Err(invalid(
            "scenario range address requires uppercase cell columns",
        ));
    }
    parse_uppercase_run(value, cursor);
    consume_optional_dollar(value, cursor);
    parse_digits(value, cursor)
}

fn parse_column_endpoint(value: &str, cursor: &mut usize) -> Result<()> {
    consume_optional_dollar(value, cursor);
    parse_digits(value, cursor)
}

fn parse_row_endpoint(value: &str, cursor: &mut usize) -> Result<()> {
    consume_optional_dollar(value, cursor);
    if !char_at(value, *cursor).is_some_and(|value| value.is_ascii_uppercase()) {
        return Err(invalid(
            "scenario range address requires uppercase row columns",
        ));
    }
    parse_uppercase_run(value, cursor);
    Ok(())
}

fn parse_uppercase_run(value: &str, cursor: &mut usize) {
    while let Some(character) = char_at(value, *cursor) {
        if !character.is_ascii_uppercase() {
            break;
        }
        *cursor += character.len_utf8();
    }
}

fn parse_digits(value: &str, cursor: &mut usize) -> Result<()> {
    let start = *cursor;
    while let Some(character) = char_at(value, *cursor) {
        if !character.is_ascii_digit() {
            break;
        }
        *cursor += character.len_utf8();
    }
    if *cursor == start {
        return Err(invalid(
            "scenario range address requires decimal row digits",
        ));
    }
    Ok(())
}

fn consume_optional_dollar(value: &str, cursor: &mut usize) {
    if char_at(value, *cursor) == Some('$') {
        *cursor += 1;
    }
}

fn expect_char(value: &str, cursor: &mut usize, expected: char) -> Result<()> {
    if char_at(value, *cursor) == Some(expected) {
        *cursor += expected.len_utf8();
        Ok(())
    } else {
        Err(invalid(format!(
            "scenario range address expected '{expected}'"
        )))
    }
}

fn char_at(value: &str, cursor: usize) -> Option<char> {
    value.get(cursor..)?.chars().next()
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
