//! Immutable static XLSX data-validation read model.

use crate::error::{OoxmlError, Result};
use litchi_core::xml::escape::escape_xml;
use litchi_ooxml_common::xml::decode_xml_reference;
use litchi_ooxml_common::{ExpandedName, MceCapabilities, MceLimits, process_markup_compatibility};
use quick_xml::Writer;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashSet, TryReserveError};
use std::fmt;
use std::ops::Range;

const CORE_URI: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_URI: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const X14_URI: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_URI: &str = "http://schemas.microsoft.com/office/excel/2006/main";
const X12AC_URI: &str = "http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac";
const XR_URI: &str = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const CORE: &[u8] = CORE_URI.as_bytes();
const STRICT: &[u8] = STRICT_URI.as_bytes();
const X14: &[u8] = X14_URI.as_bytes();
const XM: &[u8] = XM_URI.as_bytes();
const X12AC: &[u8] = X12AC_URI.as_bytes();
const XR: &[u8] = XR_URI.as_bytes();
const EXTENSION_URI: &str = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_EVENTS: usize = 1_000_000;
const MAX_NODES: usize = 1_000_000;
const MAX_CAPTURED_COLLECTIONS: usize = 1_024;
const MAX_VALIDATIONS: usize = 65_534;
const MAX_REFERENCES: usize = 32_767;
const MAX_FRAGMENT_BYTES: usize = 8 * 1024 * 1024;
const MAX_FORMULA_BYTES: usize = 1024 * 1024;
const MAX_ATTRIBUTE_BYTES: usize = MAX_FORMULA_BYTES;
const MAX_RETAINED_BYTES: usize = 16 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataValidationConformance {
    Transitional,
    Strict,
}

impl DataValidationConformance {
    fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => CORE_URI,
            Self::Strict => STRICT_URI,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DataValidationSource {
    Core,
    Office2010,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParsedDataValidationType {
    None,
    Whole,
    Decimal,
    List,
    Date,
    Time,
    TextLength,
    Custom,
}

impl ParsedDataValidationType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "whole" => Ok(Self::Whole),
            "decimal" => Ok(Self::Decimal),
            "list" => Ok(Self::List),
            "date" => Ok(Self::Date),
            "time" => Ok(Self::Time),
            "textLength" => Ok(Self::TextLength),
            "custom" => Ok(Self::Custom),
            _ => Err(invalid(format!("invalid data-validation type '{value}'"))),
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Whole => "whole",
            Self::Decimal => "decimal",
            Self::List => "list",
            Self::Date => "date",
            Self::Time => "time",
            Self::TextLength => "textLength",
            Self::Custom => "custom",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParsedDataValidationOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
}

impl ParsedDataValidationOperator {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "between" => Ok(Self::Between),
            "notBetween" => Ok(Self::NotBetween),
            "equal" => Ok(Self::Equal),
            "notEqual" => Ok(Self::NotEqual),
            "lessThan" => Ok(Self::LessThan),
            "lessThanOrEqual" => Ok(Self::LessThanOrEqual),
            "greaterThan" => Ok(Self::GreaterThan),
            "greaterThanOrEqual" => Ok(Self::GreaterThanOrEqual),
            _ => Err(invalid(format!(
                "invalid data-validation operator '{value}'"
            ))),
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::Between => "between",
            Self::NotBetween => "notBetween",
            Self::Equal => "equal",
            Self::NotEqual => "notEqual",
            Self::LessThan => "lessThan",
            Self::LessThanOrEqual => "lessThanOrEqual",
            Self::GreaterThan => "greaterThan",
            Self::GreaterThanOrEqual => "greaterThanOrEqual",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParsedDataValidationErrorStyle {
    Stop,
    Warning,
    Information,
}

impl ParsedDataValidationErrorStyle {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "stop" => Ok(Self::Stop),
            "warning" => Ok(Self::Warning),
            "information" => Ok(Self::Information),
            _ => Err(invalid(format!(
                "invalid data-validation errorStyle '{value}'"
            ))),
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::Stop => "stop",
            Self::Warning => "warning",
            Self::Information => "information",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParsedDataValidationImeMode {
    NoControl,
    Off,
    On,
    Disabled,
    Hiragana,
    FullKatakana,
    HalfKatakana,
    FullAlpha,
    HalfAlpha,
    FullHangul,
    HalfHangul,
}

impl ParsedDataValidationImeMode {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "noControl" => Ok(Self::NoControl),
            "off" => Ok(Self::Off),
            "on" => Ok(Self::On),
            "disabled" => Ok(Self::Disabled),
            "hiragana" => Ok(Self::Hiragana),
            "fullKatakana" => Ok(Self::FullKatakana),
            "halfKatakana" => Ok(Self::HalfKatakana),
            "fullAlpha" => Ok(Self::FullAlpha),
            "halfAlpha" => Ok(Self::HalfAlpha),
            "fullHangul" => Ok(Self::FullHangul),
            "halfHangul" => Ok(Self::HalfHangul),
            _ => Err(invalid(format!(
                "invalid data-validation imeMode '{value}'"
            ))),
        }
    }
    fn as_str(self) -> &'static str {
        match self {
            Self::NoControl => "noControl",
            Self::Off => "off",
            Self::On => "on",
            Self::Disabled => "disabled",
            Self::Hiragana => "hiragana",
            Self::FullKatakana => "fullKatakana",
            Self::HalfKatakana => "halfKatakana",
            Self::FullAlpha => "fullAlpha",
            Self::HalfAlpha => "halfAlpha",
            Self::FullHangul => "fullHangul",
            Self::HalfHangul => "halfHangul",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataValidationRange(String);
impl DataValidationRange {
    pub fn parse(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        let sqref = parse_sqref(&value, false, false, false, false)?;
        if sqref.ranges.len() != 1 {
            return Err(invalid("data-validation range must contain one reference"));
        }
        sqref
            .ranges
            .into_iter()
            .next()
            .ok_or_else(|| invalid("data-validation range is empty"))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataValidationFormula(String);
impl DataValidationFormula {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_text(&value, MAX_FORMULA_BYTES, "data-validation formula")?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ValidationListSource {
    Formula(DataValidationFormula),
    QuotedList(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataValidationSqref {
    ranges: Vec<DataValidationRange>,
    edited: bool,
    split: bool,
    adjusted: bool,
    adjust: bool,
}
impl DataValidationSqref {
    pub fn parse(value: impl AsRef<str>) -> Result<Self> {
        parse_sqref(value.as_ref(), false, false, false, false)
    }

    pub fn with_office2010_flags(
        mut self,
        edited: bool,
        split: bool,
        adjusted: bool,
        adjust: bool,
    ) -> Result<Self> {
        if adjusted && !adjust {
            return Err(invalid("sqref adjusted requires adjust"));
        }
        self.edited = edited;
        self.split = split;
        self.adjusted = adjusted;
        self.adjust = adjust;
        Ok(self)
    }

    pub fn ranges(&self) -> &[DataValidationRange] {
        &self.ranges
    }
    pub fn edited(&self) -> bool {
        self.edited
    }
    pub fn split(&self) -> bool {
        self.split
    }
    pub fn adjusted(&self) -> bool {
        self.adjusted
    }
    pub fn adjust(&self) -> bool {
        self.adjust
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParsedDataValidation {
    source: DataValidationSource,
    validation_type: ParsedDataValidationType,
    operator: ParsedDataValidationOperator,
    error_style: ParsedDataValidationErrorStyle,
    ime_mode: ParsedDataValidationImeMode,
    allow_blank: bool,
    show_drop_down: bool,
    show_input_message: bool,
    show_error_message: bool,
    error_title: Option<String>,
    error: Option<String>,
    prompt_title: Option<String>,
    prompt: Option<String>,
    formula1: Option<ValidationListSource>,
    formula2: Option<DataValidationFormula>,
    sqref: DataValidationSqref,
    uid: Option<String>,
}

impl ParsedDataValidation {
    pub fn new(
        source: DataValidationSource,
        validation_type: ParsedDataValidationType,
        sqref: DataValidationSqref,
    ) -> Self {
        Self {
            source,
            validation_type,
            operator: ParsedDataValidationOperator::Between,
            error_style: ParsedDataValidationErrorStyle::Stop,
            ime_mode: ParsedDataValidationImeMode::NoControl,
            allow_blank: false,
            show_drop_down: false,
            show_input_message: false,
            show_error_message: false,
            error_title: None,
            error: None,
            prompt_title: None,
            prompt: None,
            formula1: None,
            formula2: None,
            sqref,
            uid: None,
        }
    }

    pub fn set_operator(&mut self, value: ParsedDataValidationOperator) {
        self.operator = value;
    }
    pub fn set_error_style(&mut self, value: ParsedDataValidationErrorStyle) {
        self.error_style = value;
    }
    pub fn set_ime_mode(&mut self, value: ParsedDataValidationImeMode) {
        self.ime_mode = value;
    }
    pub fn set_allow_blank(&mut self, value: bool) {
        self.allow_blank = value;
    }
    pub fn set_show_drop_down(&mut self, value: bool) {
        self.show_drop_down = value;
    }
    pub fn set_show_input_message(&mut self, value: bool) {
        self.show_input_message = value;
    }
    pub fn set_show_error_message(&mut self, value: bool) {
        self.show_error_message = value;
    }

    pub fn set_error_title(&mut self, value: Option<String>) -> Result<()> {
        validate_optional_text(value.as_deref(), 32, "errorTitle")?;
        self.error_title = value;
        Ok(())
    }
    pub fn set_error(&mut self, value: Option<String>) -> Result<()> {
        validate_optional_text(value.as_deref(), 224, "error")?;
        self.error = value;
        Ok(())
    }
    pub fn set_prompt_title(&mut self, value: Option<String>) -> Result<()> {
        validate_optional_text(value.as_deref(), 32, "promptTitle")?;
        self.prompt_title = value;
        Ok(())
    }
    pub fn set_prompt(&mut self, value: Option<String>) -> Result<()> {
        validate_optional_text(value.as_deref(), 255, "prompt")?;
        self.prompt = value;
        Ok(())
    }
    pub fn set_formula1(&mut self, value: Option<ValidationListSource>) -> Result<()> {
        if let Some(ValidationListSource::Formula(value)) = value.as_ref() {
            validate_text(&value.0, MAX_FORMULA_BYTES, "formula1")?;
        }
        if let Some(ValidationListSource::QuotedList(value)) = value.as_ref() {
            validate_text(value, MAX_FORMULA_BYTES, "quoted validation list")?;
            if self.source != DataValidationSource::Office2010 {
                return Err(invalid(
                    "quoted-list source requires Office 2010 data validation",
                ));
            }
        }
        self.formula1 = value;
        Ok(())
    }
    pub fn set_formula2(&mut self, value: Option<DataValidationFormula>) -> Result<()> {
        if let Some(value) = value.as_ref() {
            validate_text(&value.0, MAX_FORMULA_BYTES, "formula2")?;
        }
        self.formula2 = value;
        Ok(())
    }
    pub fn set_uid(&mut self, value: Option<String>) -> Result<()> {
        if value.as_deref().is_some_and(|value| !valid_guid(value)) {
            return Err(invalid("invalid data-validation uid"));
        }
        self.uid = value;
        Ok(())
    }
    pub fn validate(&self) -> Result<()> {
        validate_rule(self)
    }

    pub fn source(&self) -> DataValidationSource {
        self.source
    }
    pub fn validation_type(&self) -> ParsedDataValidationType {
        self.validation_type
    }
    pub fn operator(&self) -> ParsedDataValidationOperator {
        self.operator
    }
    pub fn error_style(&self) -> ParsedDataValidationErrorStyle {
        self.error_style
    }
    pub fn ime_mode(&self) -> ParsedDataValidationImeMode {
        self.ime_mode
    }
    pub fn allow_blank(&self) -> bool {
        self.allow_blank
    }
    pub fn show_drop_down(&self) -> bool {
        self.show_drop_down
    }
    pub fn show_input_message(&self) -> bool {
        self.show_input_message
    }
    pub fn show_error_message(&self) -> bool {
        self.show_error_message
    }
    pub fn error_title(&self) -> Option<&str> {
        self.error_title.as_deref()
    }
    pub fn error(&self) -> Option<&str> {
        self.error.as_deref()
    }
    pub fn prompt_title(&self) -> Option<&str> {
        self.prompt_title.as_deref()
    }
    pub fn prompt(&self) -> Option<&str> {
        self.prompt.as_deref()
    }
    pub fn formula1(&self) -> Option<&ValidationListSource> {
        self.formula1.as_ref()
    }
    pub fn formula2(&self) -> Option<&DataValidationFormula> {
        self.formula2.as_ref()
    }
    pub fn sqref(&self) -> &DataValidationSqref {
        &self.sqref
    }
    pub fn uid(&self) -> Option<&str> {
        self.uid.as_deref()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataValidationCollection {
    source: DataValidationSource,
    disable_prompts: bool,
    x_window: Option<u32>,
    y_window: Option<u32>,
    declared_count: Option<u32>,
    validations: Vec<ParsedDataValidation>,
}
impl DataValidationCollection {
    pub fn new(
        source: DataValidationSource,
        validations: Vec<ParsedDataValidation>,
    ) -> Result<Self> {
        let value = Self {
            source,
            disable_prompts: false,
            x_window: None,
            y_window: None,
            declared_count: Some(validations.len() as u32),
            validations,
        };
        validate_collection(&value)?;
        Ok(value)
    }

    pub fn set_disable_prompts(&mut self, value: bool) {
        self.disable_prompts = value;
    }
    pub fn set_window(&mut self, x: Option<u32>, y: Option<u32>) -> Result<()> {
        if x.is_some_and(|value| value > 65_535) || y.is_some_and(|value| value > 65_535) {
            return Err(invalid("dataValidations window coordinate exceeds 65535"));
        }
        self.x_window = x;
        self.y_window = y;
        Ok(())
    }
    pub fn set_validations(&mut self, validations: Vec<ParsedDataValidation>) -> Result<()> {
        let candidate = Self {
            source: self.source,
            disable_prompts: self.disable_prompts,
            x_window: self.x_window,
            y_window: self.y_window,
            declared_count: Some(validations.len() as u32),
            validations,
        };
        validate_collection(&candidate)?;
        *self = candidate;
        Ok(())
    }

    pub fn source(&self) -> DataValidationSource {
        self.source
    }
    pub fn disable_prompts(&self) -> bool {
        self.disable_prompts
    }
    pub fn x_window(&self) -> Option<u32> {
        self.x_window
    }
    pub fn y_window(&self) -> Option<u32> {
        self.y_window
    }
    pub fn declared_count(&self) -> Option<u32> {
        self.declared_count
    }
    pub fn validations(&self) -> &[ParsedDataValidation] {
        &self.validations
    }
}

#[derive(Debug)]
struct Captured {
    source: DataValidationSource,
    prefix: Vec<u8>,
    bytes: Vec<u8>,
}

fn allocation(resource: &'static str, source: TryReserveError) -> OoxmlError {
    OoxmlError::Allocation { resource, source }
}

fn reserve_vec<T>(values: &mut Vec<T>, additional: usize, resource: &'static str) -> Result<()> {
    values
        .try_reserve_exact(additional)
        .map_err(|source| allocation(resource, source))
}

fn append_limited_text(
    value: &mut String,
    addition: &str,
    limit: usize,
    field: &str,
) -> Result<()> {
    let length = value
        .len()
        .checked_add(addition.len())
        .ok_or_else(|| invalid(format!("{field} length overflow")))?;
    if length > limit {
        return Err(invalid(format!("{field} is too large")));
    }
    value
        .try_reserve_exact(addition.len())
        .map_err(|source| allocation("data-validation text", source))?;
    value.push_str(addition);
    Ok(())
}

/// A fallible, bounded formatter used by the XML writer.
struct BoundedXml {
    value: String,
    allocation: Option<TryReserveError>,
    exceeded: bool,
}

impl BoundedXml {
    fn new() -> Self {
        Self {
            value: String::new(),
            allocation: None,
            exceeded: false,
        }
    }

    fn write_arguments(&mut self, arguments: fmt::Arguments<'_>) -> Result<()> {
        if fmt::write(self, arguments).is_ok() {
            return Ok(());
        }
        if let Some(source) = self.allocation.take() {
            return Err(allocation("data-validation XML output", source));
        }
        if self.exceeded {
            Err(invalid("data-validation XML output exceeds resource limit"))
        } else {
            Err(invalid("failed to format data-validation XML"))
        }
    }

    fn push_str(&mut self, value: &str) -> Result<()> {
        self.write_arguments(format_args!("{value}"))
    }

    fn finish(self) -> String {
        self.value
    }
}

impl fmt::Write for BoundedXml {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        let length = self
            .value
            .len()
            .checked_add(value.len())
            .ok_or(fmt::Error)?;
        if length > MAX_XML_BYTES {
            self.exceeded = true;
            return Err(fmt::Error);
        }
        if let Err(source) = self.value.try_reserve_exact(value.len()) {
            self.allocation = Some(source);
            return Err(fmt::Error);
        }
        self.value.push_str(value);
        Ok(())
    }
}

fn append_bounded_bytes(output: &mut Vec<u8>, bytes: &[u8]) -> Result<()> {
    let length = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("data-validation XML output length overflow"))?;
    if length > MAX_XML_BYTES {
        return Err(invalid("data-validation XML output exceeds resource limit"));
    }
    output
        .try_reserve_exact(bytes.len())
        .map_err(|source| allocation("data-validation XML output", source))?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn retain_capture(
    values: &mut Vec<Captured>,
    retained: &mut usize,
    captured: Captured,
) -> Result<()> {
    if values.len() >= MAX_CAPTURED_COLLECTIONS {
        return Err(invalid("too many data-validation collections"));
    }
    let size = captured
        .prefix
        .len()
        .checked_add(captured.bytes.len())
        .ok_or_else(|| invalid("data-validation retained-byte overflow"))?;
    *retained = retained
        .checked_add(size)
        .ok_or_else(|| invalid("data-validation retained-byte overflow"))?;
    if *retained > MAX_RETAINED_BYTES {
        return Err(invalid("data-validation content exceeds resource limit"));
    }
    reserve_vec(values, 1, "data-validation collections")?;
    values.push(captured);
    Ok(())
}

pub fn parse_data_validation_collections(xml: &[u8]) -> Result<Vec<DataValidationCollection>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("data-validation worksheet XML is too large"));
    }
    let mut capabilities = MceCapabilities::default();
    capabilities
        .understand_namespace(String::from_utf8_lossy(X14).into_owned())
        .understand_namespace(String::from_utf8_lossy(XM).into_owned())
        .understand_namespace(String::from_utf8_lossy(X12AC).into_owned())
        .understand_namespace(String::from_utf8_lossy(XR).into_owned());
    capabilities.preserve_extension_element(ExpandedName {
        namespace: String::from_utf8_lossy(X14).into_owned(),
        local_name: "dataValidations".into(),
    });
    let limits = MceLimits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        ..MceLimits::default()
    };
    let validated = process_markup_compatibility(xml, &capabilities, &limits)?;
    if validated.xml.len() > MAX_XML_BYTES {
        return Err(invalid("processed data-validation XML is too large"));
    }
    let selected = if validated.report.alternate_content_count == 0 {
        xml
    } else {
        validated.xml.as_ref()
    };
    let fragments = capture_collections(selected)?;
    let mut values = Vec::new();
    reserve_vec(&mut values, fragments.len(), "data-validation collections")?;
    let mut count = 0usize;
    for fragment in fragments {
        let value = parse_collection(&fragment)?;
        count = count
            .checked_add(value.validations.len())
            .ok_or_else(|| invalid("data-validation count overflow"))?;
        if count > MAX_VALIDATIONS {
            return Err(invalid("too many data validations"));
        }
        values.push(value);
    }
    validate_data_validation_collections(&values)?;
    Ok(values)
}

/// In-flight capture state for a single `dataValidation` element.
type CaptureState = Option<(usize, DataValidationSource, Vec<u8>, Writer<Vec<u8>>)>;

fn capture_collections(xml: &[u8]) -> Result<Vec<Captured>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut values = Vec::new();
    let mut depth = 0usize;
    let mut extension_depth = None;
    let mut capture: CaptureState = None;
    let mut events = 0usize;
    let mut nodes = 0usize;
    let mut retained = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("data-validation XML event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("data-validation XML exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Eof) {
            if capture.is_some() || depth != 0 {
                return Err(invalid("unterminated data-validation worksheet XML"));
            }
            break;
        }
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("data-validation XML node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("data-validation XML exceeds node limit"));
            }
        }
        if let Some((capture_depth, _, _, writer)) = capture.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                return Err(invalid("dataValidations fragment is too large"));
            }
            match event {
                Event::Start(_) => {
                    if *capture_depth >= MAX_DEPTH {
                        return Err(invalid("dataValidations nesting is too deep"));
                    }
                    *capture_depth = capture_depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidations nesting overflow"))?;
                },
                Event::End(_) => {
                    *capture_depth = capture_depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid dataValidations nesting"))?;
                },
                _ => {},
            }
            if *capture_depth == 0 {
                let Some((_, source, prefix, writer)) = capture.take() else {
                    return Err(invalid("dataValidations capture state disappeared"));
                };
                let bytes = writer.into_inner();
                retain_capture(
                    &mut values,
                    &mut retained,
                    Captured {
                        source,
                        prefix,
                        bytes,
                    },
                )?;
            }
            continue;
        }
        match event {
            Event::Start(element)
                if element.local_name().as_ref() == b"dataValidations"
                    && depth > 0
                    && depth == 1
                    && spreadsheet(&namespace) =>
            {
                let prefix = prefix(element.name().as_ref())?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidations fragment is too large"));
                }
                capture = Some((1, DataValidationSource::Core, prefix, writer));
            },
            Event::Start(element)
                if element.local_name().as_ref() == b"dataValidations"
                    && depth > 0
                    && exact(&namespace, X14)
                    && extension_depth.is_some() =>
            {
                let prefix = prefix(element.name().as_ref())?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidations fragment is too large"));
                }
                capture = Some((1, DataValidationSource::Office2010, prefix, writer));
            },
            Event::Empty(element)
                if element.local_name().as_ref() == b"dataValidations"
                    && depth == 1
                    && spreadsheet(&namespace) =>
            {
                let prefix = prefix(element.name().as_ref())?;
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidations fragment is too large"));
                }
                retain_capture(
                    &mut values,
                    &mut retained,
                    Captured {
                        source: DataValidationSource::Core,
                        prefix,
                        bytes: writer.into_inner(),
                    },
                )?;
            },
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if depth == 0 {
                    if root_seen
                        || !spreadsheet(&namespace)
                        || element.local_name().as_ref() != b"worksheet"
                    {
                        return Err(invalid("data-validation parser requires a worksheet root"));
                    }
                    root_seen = true;
                }
                if depth >= MAX_DEPTH {
                    return Err(invalid("worksheet nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet nesting overflow"))?;
                if spreadsheet(&namespace)
                    && element.local_name().as_ref() == b"ext"
                    && optional_attr(&element, b"uri", decoder)?.as_deref() == Some(EXTENSION_URI)
                {
                    extension_depth = Some(depth);
                }
            },
            Event::Empty(element) if depth == 0 => {
                if root_seen
                    || !spreadsheet(&namespace)
                    || element.local_name().as_ref() != b"worksheet"
                {
                    return Err(invalid("data-validation parser requires a worksheet root"));
                }
                root_seen = true;
                root_closed = true;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid worksheet nesting"));
                }
                if depth == 1
                    && (!spreadsheet(&namespace) || element.local_name().as_ref() != b"worksheet")
                {
                    return Err(invalid("invalid worksheet closing element"));
                }
                if extension_depth == Some(depth) {
                    extension_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid worksheet nesting"))?;
                if depth == 0 {
                    // The namespace/local-name check is performed by the XML reader for
                    // qualified names; this branch only records the root boundary.
                    root_closed = true;
                }
            },
            Event::Text(value) => {
                if (!root_seen || root_closed)
                    && !value.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("worksheet XML text is outside root"));
                }
                if depth == 1 && !value.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("worksheet cannot contain direct text"));
                }
            },
            Event::CData(_) if depth == 1 || !root_seen || root_closed => {
                return Err(invalid("worksheet XML contains unexpected CDATA"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen {
                    return Err(invalid("invalid worksheet XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::GeneralRef(reference) => {
                decode_xml_reference(&reference)?;
            },
            _ => {},
        }
    }
    if !root_seen || !root_closed || depth != 0 {
        return Err(invalid("incomplete worksheet data-validation XML"));
    }
    Ok(values)
}

fn parse_collection(fragment: &Captured) -> Result<DataValidationCollection> {
    let wrapped = wrap(&fragment.prefix, &fragment.bytes)?;
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root_depth = None;
    let mut closed = false;
    let mut expected = None;
    let mut disable = false;
    let mut x_window = None;
    let mut y_window = None;
    let mut validations = Vec::new();
    let mut capture: Option<(usize, Writer<Vec<u8>>)> = None;
    let mut events = 0usize;
    let mut nodes = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("dataValidations event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("dataValidations exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Eof) {
            if capture.is_some() || depth != 0 || !closed {
                return Err(invalid("unterminated dataValidations"));
            }
            break;
        }
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("dataValidations node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("dataValidations exceeds node limit"));
            }
        }
        if let Some((capture_depth, writer)) = capture.as_mut() {
            writer.write_event(event.clone()).map_err(xml_error)?;
            if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                return Err(invalid("dataValidation rule is too large"));
            }
            match event {
                Event::Start(_) => {
                    if *capture_depth >= MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    *capture_depth = capture_depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                },
                Event::End(_) => {
                    *capture_depth = capture_depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("invalid dataValidation nesting"))?;
                },
                _ => {},
            }
            if *capture_depth == 0 {
                let Some((_, writer)) = capture.take() else {
                    return Err(invalid("dataValidation capture state disappeared"));
                };
                if validations.len() >= MAX_VALIDATIONS {
                    return Err(invalid("too many data validations"));
                }
                let raw = writer.into_inner();
                let value = parse_rule(&raw, fragment.source)?;
                reserve_vec(&mut validations, 1, "data-validation rules")?;
                validations.push(value);
            }
            continue;
        }
        match event {
            Event::Start(element)
                if element.local_name().as_ref() == b"dataValidations"
                    && source_ns(fragment.source, &namespace) =>
            {
                if root_depth.is_some() {
                    return Err(invalid("nested dataValidations"));
                }
                if closed || depth >= MAX_DEPTH {
                    return Err(invalid("dataValidations nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("dataValidations nesting overflow"))?;
                root_depth = Some(depth);
                expected = optional_u32(&element, b"count", decoder)?;
                if expected.is_some_and(|value| value as usize > MAX_VALIDATIONS) {
                    return Err(invalid("too many data validations"));
                }
                disable = optional_bool(&element, b"disablePrompts", decoder)?.unwrap_or(false);
                x_window = optional_u32(&element, b"xWindow", decoder)?;
                y_window = optional_u32(&element, b"yWindow", decoder)?;
            },
            Event::Start(element)
                if element.local_name().as_ref() == b"dataValidation"
                    && source_ns(fragment.source, &namespace)
                    && root_depth == Some(depth) =>
            {
                if closed {
                    return Err(invalid("content follows dataValidations"));
                }
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidation rule is too large"));
                }
                capture = Some((1, writer));
            },
            Event::Empty(element)
                if element.local_name().as_ref() == b"dataValidation"
                    && source_ns(fragment.source, &namespace)
                    && root_depth == Some(depth) =>
            {
                if closed || validations.len() >= MAX_VALIDATIONS {
                    return Err(invalid("too many data validations"));
                }
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Empty(element))
                    .map_err(xml_error)?;
                if writer.get_ref().len() > MAX_FRAGMENT_BYTES {
                    return Err(invalid("dataValidation rule is too large"));
                }
                let raw = writer.into_inner();
                let value = parse_rule(&raw, fragment.source)?;
                reserve_vec(&mut validations, 1, "data-validation rules")?;
                validations.push(value);
            },
            Event::Start(_) => {
                if depth >= MAX_DEPTH {
                    return Err(invalid("dataValidations nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("dataValidations nesting overflow"))?
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected dataValidations closing element"));
                }
                if root_depth == Some(depth)
                    && element.local_name().as_ref() == b"dataValidations"
                    && source_ns(fragment.source, &namespace)
                {
                    closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid dataValidations nesting"))?;
            },
            Event::Text(value) => {
                if !value.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("dataValidations must not contain text"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) => {
                return Err(invalid("dataValidations must not contain character data"));
            },
            Event::Decl(_) => {
                return Err(invalid(
                    "XML declarations are not allowed in dataValidations",
                ));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Empty(_) => {},
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid("unexpected EOF in dataValidations")),
        }
    }
    if validations.is_empty() {
        return Err(invalid("dataValidations must contain at least one rule"));
    }
    if expected.is_some_and(|v| v as usize != validations.len()) {
        return Err(invalid("dataValidations count does not match its children"));
    }
    Ok(DataValidationCollection {
        source: fragment.source,
        disable_prompts: disable,
        x_window,
        y_window,
        declared_count: expected,
        validations,
    })
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum TextTarget {
    Formula1,
    Formula2,
    Sqref,
    List,
}

fn text_target_matches(
    target: TextTarget,
    source: DataValidationSource,
    namespace: &ResolveResult<'_>,
    local: &[u8],
) -> bool {
    match target {
        TextTarget::Formula1 => {
            (source == DataValidationSource::Core
                && local == b"formula1"
                && source_ns(source, namespace))
                || (source == DataValidationSource::Office2010
                    && local == b"f"
                    && exact(namespace, XM))
        },
        TextTarget::Formula2 => {
            (source == DataValidationSource::Core
                && local == b"formula2"
                && source_ns(source, namespace))
                || (source == DataValidationSource::Office2010
                    && local == b"f"
                    && exact(namespace, XM))
        },
        TextTarget::Sqref => local == b"sqref" && exact(namespace, XM),
        TextTarget::List => local == b"list" && exact(namespace, X12AC),
    }
}

fn parse_rule(raw: &[u8], source: DataValidationSource) -> Result<ParsedDataValidation> {
    if raw.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("dataValidation rule is too large"));
    }
    let wrapped = wrap(
        if source == DataValidationSource::Core {
            b""
        } else {
            b"x14"
        },
        raw,
    )?;
    let mut reader = NsReader::from_reader(wrapped.as_slice());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut rule_depth = None;
    let mut closed = false;
    let mut order = 0u8;
    let mut kind = ParsedDataValidationType::None;
    let mut operator = ParsedDataValidationOperator::Between;
    let mut error_style = ParsedDataValidationErrorStyle::Stop;
    let mut ime = ParsedDataValidationImeMode::NoControl;
    let mut allow_blank = false;
    let mut show_drop_down = false;
    let mut show_input = false;
    let mut show_error = false;
    let (mut error_title, mut error, mut prompt_title, mut prompt, mut uid) =
        (None, None, None, None, None);
    let (mut formula1, mut formula2, mut sqref): (
        Option<ValidationListSource>,
        Option<DataValidationFormula>,
        Option<DataValidationSqref>,
    ) = (None, None, None);
    let mut wrapper: Option<(u8, usize, bool)> = None;
    let mut text: Option<(usize, TextTarget, String)> = None;
    let mut events = 0usize;
    let mut nodes = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("dataValidation event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("dataValidation exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Eof) {
            break;
        }
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("dataValidation node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("dataValidation exceeds node limit"));
            }
        }
        match event {
            Event::Start(element) => {
                if closed {
                    return Err(invalid("content follows dataValidation"));
                }
                let local = element.local_name();
                if local.as_ref() == b"dataValidation"
                    && source_ns(source, &namespace)
                    && rule_depth.is_none()
                {
                    if depth >= MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                    rule_depth = Some(depth);
                    kind = ParsedDataValidationType::parse(
                        optional_attr(&element, b"type", decoder)?
                            .as_deref()
                            .unwrap_or("none"),
                    )?;
                    operator = ParsedDataValidationOperator::parse(
                        optional_attr(&element, b"operator", decoder)?
                            .as_deref()
                            .unwrap_or("between"),
                    )?;
                    error_style = ParsedDataValidationErrorStyle::parse(
                        optional_attr(&element, b"errorStyle", decoder)?
                            .as_deref()
                            .unwrap_or("stop"),
                    )?;
                    ime = ParsedDataValidationImeMode::parse(
                        optional_attr(&element, b"imeMode", decoder)?
                            .as_deref()
                            .unwrap_or("noControl"),
                    )?;
                    allow_blank = optional_bool(&element, b"allowBlank", decoder)?.unwrap_or(false);
                    show_drop_down =
                        optional_bool(&element, b"showDropDown", decoder)?.unwrap_or(false);
                    show_input =
                        optional_bool(&element, b"showInputMessage", decoder)?.unwrap_or(false);
                    show_error =
                        optional_bool(&element, b"showErrorMessage", decoder)?.unwrap_or(false);
                    error_title = bounded_attr(&element, b"errorTitle", decoder, 32)?;
                    error = bounded_attr(&element, b"error", decoder, 225)?;
                    prompt_title = bounded_attr(&element, b"promptTitle", decoder, 32)?;
                    prompt = bounded_attr(&element, b"prompt", decoder, 255)?;
                    uid = uid_attr(&element, decoder, &resolver)?;
                    if source == DataValidationSource::Core {
                        sqref = Some(parse_sqref(
                            &required_attr(&element, b"sqref", decoder)?,
                            false,
                            false,
                            false,
                            false,
                        )?);
                    }
                } else if rule_depth == Some(depth)
                    && source_ns(source, &namespace)
                    && matches!(local.as_ref(), b"formula1" | b"formula2")
                {
                    let number = if local.as_ref() == b"formula1" { 1 } else { 2 };
                    if number < order {
                        return Err(invalid("dataValidation children are out of order"));
                    }
                    order = number;
                    if wrapper.is_some() {
                        return Err(invalid("nested data-validation formula wrapper"));
                    }
                    let target_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                    wrapper = Some((number, target_depth, false));
                    if target_depth > MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    depth = target_depth;
                    if source == DataValidationSource::Core {
                        text = Some((
                            depth,
                            if number == 1 {
                                TextTarget::Formula1
                            } else {
                                TextTarget::Formula2
                            },
                            String::new(),
                        ));
                    }
                } else if source == DataValidationSource::Office2010
                    && wrapper.is_some()
                    && rule_depth.is_some_and(|value| depth == value + 1)
                    && exact(&namespace, XM)
                    && local.as_ref() == b"f"
                {
                    let Some(wrapper_state) = wrapper.as_mut() else {
                        return Err(invalid("data-validation formula outside its wrapper"));
                    };
                    if wrapper_state.2 {
                        return Err(invalid("formula wrapper must contain exactly one value"));
                    }
                    wrapper_state.2 = true;
                    depth += 1;
                    let target = if wrapper_state.0 == 1 {
                        TextTarget::Formula1
                    } else {
                        TextTarget::Formula2
                    };
                    text = Some((depth, target, String::new()));
                } else if source == DataValidationSource::Office2010
                    && wrapper.is_some()
                    && rule_depth.is_some_and(|value| depth == value + 1)
                    && exact(&namespace, X12AC)
                    && local.as_ref() == b"list"
                {
                    let Some(wrapper_state) = wrapper.as_mut() else {
                        return Err(invalid("quoted validation list is outside its wrapper"));
                    };
                    if wrapper_state.0 != 1 {
                        return Err(invalid("quoted validation list is only valid in formula1"));
                    }
                    if wrapper_state.2 {
                        return Err(invalid("formula wrapper must contain exactly one value"));
                    }
                    wrapper_state.2 = true;
                    depth += 1;
                    text = Some((depth, TextTarget::List, String::new()));
                } else if source == DataValidationSource::Office2010
                    && rule_depth == Some(depth)
                    && exact(&namespace, XM)
                    && local.as_ref() == b"sqref"
                {
                    if order > 3 {
                        return Err(invalid("dataValidation children are out of order"));
                    }
                    order = 3;
                    let flags = sqref_flags(&element, decoder)?;
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                    if depth > MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    text = Some((depth, TextTarget::Sqref, encode_flags(flags)));
                } else {
                    if depth >= MAX_DEPTH {
                        return Err(invalid("dataValidation nesting is too deep"));
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("dataValidation nesting overflow"))?;
                }
            },
            Event::Text(value) => {
                if let Some((_, _, buffer)) = text.as_mut() {
                    let decoded = value.decode().map_err(xml_error)?;
                    append_limited_text(
                        buffer,
                        &decoded,
                        MAX_FORMULA_BYTES,
                        "data-validation text",
                    )?;
                } else if !value.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("dataValidation contains unexpected text"));
                }
            },
            Event::CData(value) => {
                if let Some((_, _, buffer)) = text.as_mut() {
                    let decoded = value.decode().map_err(xml_error)?;
                    append_limited_text(
                        buffer,
                        &decoded,
                        MAX_FORMULA_BYTES,
                        "data-validation text",
                    )?;
                } else {
                    return Err(invalid("dataValidation contains unexpected CDATA"));
                }
            },
            Event::GeneralRef(value) => {
                if let Some((_, _, buffer)) = text.as_mut() {
                    let decoded = decode_xml_reference(&value)?;
                    append_limited_text(
                        buffer,
                        &decoded,
                        MAX_FORMULA_BYTES,
                        "data-validation text",
                    )?;
                } else {
                    return Err(invalid("dataValidation contains unexpected entity text"));
                }
            },
            Event::Empty(element) => {
                if closed {
                    return Err(invalid("content follows dataValidation"));
                }
                let local = element.local_name();
                if source == DataValidationSource::Core
                    && rule_depth == Some(depth)
                    && spreadsheet(&namespace)
                    && matches!(local.as_ref(), b"formula1" | b"formula2")
                {
                    let number = if local.as_ref() == b"formula1" { 1 } else { 2 };
                    if number < order {
                        return Err(invalid("dataValidation children are out of order"));
                    }
                    order = number;
                    if number == 1 {
                        if formula1
                            .replace(ValidationListSource::Formula(DataValidationFormula(
                                String::new(),
                            )))
                            .is_some()
                        {
                            return Err(invalid("duplicate formula1"));
                        }
                    } else if formula2
                        .replace(DataValidationFormula(String::new()))
                        .is_some()
                    {
                        return Err(invalid("duplicate formula2"));
                    }
                } else if source == DataValidationSource::Office2010
                    && wrapper.is_some()
                    && rule_depth.is_some_and(|value| depth == value + 1)
                    && exact(&namespace, XM)
                    && local.as_ref() == b"f"
                {
                    let Some(wrapper_state) = wrapper.as_mut() else {
                        return Err(invalid("formula value is outside its wrapper"));
                    };
                    if wrapper_state.2 {
                        return Err(invalid("formula wrapper must contain exactly one value"));
                    }
                    let number = wrapper_state.0;
                    wrapper_state.2 = true;
                    if number == 1 {
                        if formula1
                            .replace(ValidationListSource::Formula(DataValidationFormula(
                                String::new(),
                            )))
                            .is_some()
                        {
                            return Err(invalid("duplicate formula1"));
                        }
                    } else if formula2
                        .replace(DataValidationFormula(String::new()))
                        .is_some()
                    {
                        return Err(invalid("duplicate formula2"));
                    }
                } else if source == DataValidationSource::Office2010
                    && wrapper.is_some()
                    && rule_depth.is_some_and(|value| depth == value + 1)
                    && exact(&namespace, X12AC)
                    && local.as_ref() == b"list"
                {
                    let Some(wrapper_state) = wrapper.as_mut() else {
                        return Err(invalid("data-validation list source outside its wrapper"));
                    };
                    if wrapper_state.0 != 1 || wrapper_state.2 {
                        return Err(invalid("invalid quoted-list formula wrapper"));
                    }
                    wrapper_state.2 = true;
                    if formula1
                        .replace(ValidationListSource::QuotedList(String::new()))
                        .is_some()
                    {
                        return Err(invalid("duplicate formula1 source"));
                    }
                } else if source == DataValidationSource::Office2010
                    && rule_depth == Some(depth)
                    && exact(&namespace, XM)
                    && local.as_ref() == b"sqref"
                {
                    return Err(invalid("data-validation sqref is empty"));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected dataValidation closing element"));
                }
                if text.as_ref().is_some_and(|(target, _, _)| *target == depth) {
                    let Some((_, target, value)) = text.take() else {
                        return Err(invalid("data-validation text state disappeared"));
                    };
                    if !text_target_matches(
                        target,
                        source,
                        &namespace,
                        element.local_name().as_ref(),
                    ) {
                        return Err(invalid("invalid data-validation text closing element"));
                    }
                    match target {
                        TextTarget::Formula1 => {
                            if formula1.is_some() {
                                return Err(invalid("duplicate formula1"));
                            }
                            formula1 =
                                Some(ValidationListSource::Formula(DataValidationFormula(value)));
                        },
                        TextTarget::Formula2 => {
                            if formula2.is_some() {
                                return Err(invalid("duplicate formula2"));
                            }
                            formula2 = Some(DataValidationFormula(value));
                        },
                        TextTarget::List => {
                            if formula1.is_some() {
                                return Err(invalid("duplicate formula1 source"));
                            }
                            formula1 = Some(ValidationListSource::QuotedList(value));
                        },
                        TextTarget::Sqref => {
                            let (flags, value) = decode_flags(&value)?;
                            sqref = Some(parse_sqref(&value, flags.0, flags.1, flags.2, flags.3)?);
                        },
                    }
                }
                if wrapper
                    .as_ref()
                    .is_some_and(|(_, target, _)| *target == depth)
                {
                    let Some((number, _, seen)) = wrapper.take() else {
                        return Err(invalid("data-validation wrapper state disappeared"));
                    };
                    let expected = if number == 1 {
                        b"formula1"
                    } else {
                        b"formula2"
                    };
                    if !source_ns(source, &namespace) || element.local_name().as_ref() != expected {
                        return Err(invalid("invalid data-validation wrapper closing element"));
                    }
                    if source == DataValidationSource::Office2010 && !seen {
                        return Err(invalid(
                            "x14 formula wrapper must contain exactly one value",
                        ));
                    }
                }
                if rule_depth == Some(depth)
                    && element.local_name().as_ref() == b"dataValidation"
                    && source_ns(source, &namespace)
                {
                    closed = true;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid dataValidation nesting"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) => {
                return Err(invalid(
                    "XML declarations are not allowed in dataValidation",
                ));
            },
            Event::Comment(_) => {},
            Event::Eof => return Err(invalid("unexpected EOF in dataValidation")),
        }
    }
    if !closed || wrapper.is_some() || text.is_some() || depth != 0 {
        return Err(invalid("unterminated dataValidation"));
    }
    let sqref = sqref.ok_or_else(|| invalid("dataValidation is missing sqref"))?;
    validate_formula_cardinality(kind, operator, &formula1, &formula2)?;
    Ok(ParsedDataValidation {
        source,
        validation_type: kind,
        operator,
        error_style,
        ime_mode: ime,
        allow_blank,
        show_drop_down,
        show_input_message: show_input,
        show_error_message: show_error,
        error_title,
        error,
        prompt_title,
        prompt,
        formula1,
        formula2,
        sqref,
        uid,
    })
}

fn validate_formula_cardinality(
    kind: ParsedDataValidationType,
    operator: ParsedDataValidationOperator,
    f1: &Option<ValidationListSource>,
    f2: &Option<DataValidationFormula>,
) -> Result<()> {
    match kind {
        ParsedDataValidationType::None => {
            if f1.is_some() || f2.is_some() {
                return Err(invalid("type none must not contain formulas"));
            }
        },
        ParsedDataValidationType::List | ParsedDataValidationType::Custom => {
            if f1.is_none() || f2.is_some() {
                return Err(invalid(
                    "list/custom validation requires exactly formula1 or a quoted list",
                ));
            }
        },
        _ if matches!(
            operator,
            ParsedDataValidationOperator::Between | ParsedDataValidationOperator::NotBetween
        ) =>
        {
            if f1.is_none() || f2.is_none() {
                return Err(invalid("between validation requires formula1 and formula2"));
            }
        },
        _ => {
            if f1.is_none() || f2.is_some() {
                return Err(invalid("validation requires formula1 and forbids formula2"));
            }
        },
    }
    Ok(())
}

pub fn validate_data_validation_collections(values: &[DataValidationCollection]) -> Result<()> {
    let mut sources = HashSet::new();
    let mut count = 0usize;
    for collection in values {
        if !sources.insert(collection.source) {
            return Err(invalid("duplicate dataValidations collection source"));
        }
        validate_collection(collection)?;
        count = count
            .checked_add(collection.validations.len())
            .ok_or_else(|| invalid("data-validation count overflow"))?;
        if count > MAX_VALIDATIONS {
            return Err(invalid("too many data validations"));
        }
    }
    Ok(())
}

fn validate_collection(value: &DataValidationCollection) -> Result<()> {
    if value.validations.is_empty() || value.validations.len() > MAX_VALIDATIONS {
        return Err(invalid("dataValidations has an invalid rule count"));
    }
    if value.x_window.is_some_and(|v| v > 65_535) || value.y_window.is_some_and(|v| v > 65_535) {
        return Err(invalid("dataValidations window coordinate exceeds 65535"));
    }
    for rule in &value.validations {
        if rule.source != value.source {
            return Err(invalid(
                "dataValidation source does not match its collection",
            ));
        }
        validate_rule(rule)?;
    }
    Ok(())
}

fn validate_rule(value: &ParsedDataValidation) -> Result<()> {
    validate_formula_cardinality(
        value.validation_type,
        value.operator,
        &value.formula1,
        &value.formula2,
    )?;
    validate_optional_text(value.error_title.as_deref(), 32, "errorTitle")?;
    validate_optional_text(value.error.as_deref(), 224, "error")?;
    validate_optional_text(value.prompt_title.as_deref(), 32, "promptTitle")?;
    validate_optional_text(value.prompt.as_deref(), 255, "prompt")?;
    if value.uid.as_deref().is_some_and(|uid| !valid_guid(uid)) {
        return Err(invalid("invalid data-validation uid"));
    }
    if value.source == DataValidationSource::Core
        && (value.sqref.edited || value.sqref.split || value.sqref.adjusted || value.sqref.adjust)
    {
        return Err(invalid(
            "Office 2010 sqref flags are not valid on core data validation",
        ));
    }
    parse_sqref(
        &sqref_text(&value.sqref)?,
        value.sqref.edited,
        value.sqref.split,
        value.sqref.adjusted,
        value.sqref.adjust,
    )?;
    match value.formula1.as_ref() {
        Some(ValidationListSource::Formula(value)) => {
            validate_text(&value.0, MAX_FORMULA_BYTES, "formula1")?
        },
        Some(ValidationListSource::QuotedList(list)) => {
            if value.source != DataValidationSource::Office2010 {
                return Err(invalid(
                    "quoted-list source requires Office 2010 data validation",
                ));
            }
            validate_text(list, MAX_FORMULA_BYTES, "quoted validation list")?;
        },
        None => {},
    }
    if let Some(value) = value.formula2.as_ref() {
        validate_text(&value.0, MAX_FORMULA_BYTES, "formula2")?;
    }
    Ok(())
}

fn validate_optional_text(value: Option<&str>, max: usize, field: &str) -> Result<()> {
    if let Some(value) = value {
        if value.chars().count() > max {
            return Err(invalid(format!("{field} exceeds {max} characters")));
        }
        validate_xml_chars(value, field)?;
    }
    Ok(())
}

fn validate_text(value: &str, max_bytes: usize, field: &str) -> Result<()> {
    if value.len() > max_bytes {
        return Err(invalid(format!("{field} is too large")));
    }
    validate_xml_chars(value, field)
}

fn validate_xml_chars(value: &str, field: &str) -> Result<()> {
    if value.chars().any(|ch| {
        let code = ch as u32;
        !matches!(code, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
    }) {
        return Err(invalid(format!(
            "{field} contains an invalid XML character"
        )));
    }
    Ok(())
}

fn sqref_text(value: &DataValidationSqref) -> Result<String> {
    let mut text = String::new();
    for (index, range) in value.ranges.iter().enumerate() {
        if index != 0 {
            append_limited_text(&mut text, " ", MAX_FRAGMENT_BYTES, "data-validation sqref")?;
        }
        append_limited_text(
            &mut text,
            range.0.as_str(),
            MAX_FRAGMENT_BYTES,
            "data-validation sqref",
        )?;
    }
    Ok(text)
}

fn parse_sqref(
    value: &str,
    edited: bool,
    split: bool,
    adjusted: bool,
    adjust: bool,
) -> Result<DataValidationSqref> {
    if value.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("data-validation sqref is too large"));
    }
    if adjusted && !adjust {
        return Err(invalid("sqref adjusted requires adjust"));
    }
    let mut ranges = Vec::new();
    for item in value.split_whitespace() {
        if ranges.len() == MAX_REFERENCES {
            return Err(invalid("too many data-validation references"));
        }
        let mut parts = item.split(':');
        let Some(first) = parts.next() else {
            return Err(invalid("invalid empty data-validation range"));
        };
        let second = parts.next();
        if parts.next().is_some() || !valid_cell(first) || second.is_some_and(|v| !valid_cell(v)) {
            return Err(invalid(format!("invalid data-validation range '{item}'")));
        }
        reserve_vec(&mut ranges, 1, "data-validation references")?;
        ranges.push(DataValidationRange(item.to_owned()));
    }
    if ranges.is_empty() {
        return Err(invalid("data-validation sqref is empty"));
    }
    Ok(DataValidationSqref {
        ranges,
        edited,
        split,
        adjusted,
        adjust,
    })
}

fn valid_cell(value: &str) -> bool {
    let raw = value.as_bytes();
    let mut i = 0;
    if i < raw.len() && raw[i] == b'$' {
        i += 1;
    }
    let start = i;
    while i < raw.len() && raw[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == start {
        return false;
    }
    let mut col = 0u32;
    for b in &raw[start..i] {
        col = col
            .saturating_mul(26)
            .saturating_add(u32::from(b.to_ascii_uppercase() - b'A' + 1));
    }
    if !(1..=16_384).contains(&col) {
        return false;
    }
    if i < raw.len() && raw[i] == b'$' {
        i += 1;
    }
    let Ok(row_text) = std::str::from_utf8(&raw[i..]) else {
        return false;
    };
    let Ok(row) = row_text.parse::<u32>() else {
        return false;
    };
    (1..=1_048_576).contains(&row)
}

fn sqref_flags(element: &BytesStart<'_>, decoder: Decoder) -> Result<(bool, bool, bool, bool)> {
    Ok((
        optional_bool(element, b"edited", decoder)?.unwrap_or(false),
        optional_bool(element, b"split", decoder)?.unwrap_or(false),
        optional_bool(element, b"adjusted", decoder)?.unwrap_or(false),
        optional_bool(element, b"adjust", decoder)?.unwrap_or(false),
    ))
}
fn encode_flags(v: (bool, bool, bool, bool)) -> String {
    format!("{}{}{}{}|", v.0 as u8, v.1 as u8, v.2 as u8, v.3 as u8)
}
fn decode_flags(value: &str) -> Result<((bool, bool, bool, bool), String)> {
    let bytes = value.as_bytes();
    if bytes.len() < 5 || bytes[4] != b'|' {
        return Err(invalid("invalid sqref state"));
    }
    Ok((
        (
            bytes[0] == b'1',
            bytes[1] == b'1',
            bytes[2] == b'1',
            bytes[3] == b'1',
        ),
        value[5..].to_owned(),
    ))
}

fn uid_attr(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
        if attribute.key.local_name().as_ref() == b"uid"
            && (exact(&attribute_namespace, XR) || attribute.key.as_ref() == b"xr:uid")
        {
            if result.is_some() {
                return Err(invalid("duplicate data-validation uid"));
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            if value.len() > MAX_ATTRIBUTE_BYTES {
                return Err(invalid("data-validation uid is too large"));
            }
            if !valid_guid(&value) {
                return Err(invalid("invalid data-validation uid"));
            }
            result = Some(value);
        }
    }
    Ok(result)
}
fn valid_guid(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 38
        && bytes[0] == b'{'
        && bytes[37] == b'}'
        && [9, 14, 19, 24].iter().all(|i| bytes[*i] == b'-')
        && bytes[1..37]
            .iter()
            .enumerate()
            .all(|(i, b)| matches!(i, 8 | 13 | 18 | 23) || b.is_ascii_hexdigit())
}

fn optional_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if attribute.key.as_ref() == name {
            if result.is_some() {
                return Err(invalid(format!(
                    "duplicate '{}' attribute",
                    String::from_utf8_lossy(name)
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            if value.len() > MAX_ATTRIBUTE_BYTES {
                return Err(invalid(format!(
                    "data-validation attribute '{}' is too large",
                    String::from_utf8_lossy(name)
                )));
            }
            result = Some(value);
        }
    }
    Ok(result)
}
fn required_attr(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<String> {
    optional_attr(element, name, decoder)?.ok_or_else(|| {
        invalid(format!(
            "missing '{}' attribute",
            String::from_utf8_lossy(name)
        ))
    })
}
fn optional_u32(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<u32>> {
    optional_attr(element, name, decoder)?
        .map(|v| {
            v.parse()
                .map_err(|_| invalid(format!("invalid unsigned integer '{v}'")))
        })
        .transpose()
}
fn optional_bool(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<bool>> {
    optional_attr(element, name, decoder)?
        .map(|v| match v.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid boolean '{v}'"))),
        })
        .transpose()
}
fn bounded_attr(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    max: usize,
) -> Result<Option<String>> {
    let value = optional_attr(element, name, decoder)?;
    if value.as_ref().is_some_and(|v| v.chars().count() > max) {
        return Err(invalid(format!(
            "{} exceeds {max} characters",
            String::from_utf8_lossy(name)
        )));
    }
    Ok(value)
}

fn wrap(prefix: &[u8], fragment: &[u8]) -> Result<Vec<u8>> {
    if fragment.len() > MAX_FRAGMENT_BYTES {
        return Err(invalid("data-validation fragment is too large"));
    }
    let mut out = Vec::new();
    append_bounded_bytes(
        &mut out,
        br#"<root xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:x12ac="http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision""#,
    )?;
    if !prefix.is_empty() && !matches!(prefix, b"s" | b"x14") {
        append_bounded_bytes(&mut out, b" xmlns:")?;
        append_bounded_bytes(&mut out, prefix)?;
        append_bounded_bytes(
            &mut out,
            if prefix == b"x" {
                b"=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
            } else {
                b"=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\""
            },
        )?;
    }
    append_bounded_bytes(&mut out, b">")?;
    append_bounded_bytes(&mut out, fragment)?;
    append_bounded_bytes(&mut out, b"</root>")?;
    Ok(out)
}
fn prefix(name: &[u8]) -> Result<Vec<u8>> {
    let Some(index) = name.iter().position(|value| *value == b':') else {
        return Ok(Vec::new());
    };
    let mut value = Vec::new();
    reserve_vec(&mut value, index, "data-validation namespace prefix")?;
    value.extend_from_slice(&name[..index]);
    Ok(value)
}
fn source_ns(source: DataValidationSource, ns: &ResolveResult<'_>) -> bool {
    match source {
        DataValidationSource::Core => spreadsheet(ns),
        DataValidationSource::Office2010 => exact(ns, X14),
    }
}
fn spreadsheet(ns: &ResolveResult<'_>) -> bool {
    exact(ns, CORE) || exact(ns, STRICT)
}
fn exact(ns: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(ns,ResolveResult::Bound(value)if value.as_ref()==expected)
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

/// Write canonical core and Office 2010 data-validation fragments.
pub fn write_data_validation_collections(
    values: &[DataValidationCollection],
    conformance: DataValidationConformance,
) -> Result<String> {
    let core = write_data_validation_core(values, conformance)?;
    let extensions = write_data_validation_extensions(values, conformance)?;
    let mut xml = BoundedXml::new();
    xml.push_str(&core)?;
    xml.push_str(&extensions)?;
    Ok(xml.finish())
}

pub(crate) fn write_data_validation_core(
    values: &[DataValidationCollection],
    conformance: DataValidationConformance,
) -> Result<String> {
    validate_data_validation_collections(values)?;
    let mut xml = BoundedXml::new();
    if let Some(collection) = values
        .iter()
        .find(|collection| collection.source == DataValidationSource::Core)
    {
        xml.write_arguments(format_args!(
            "<dataValidations xmlns=\"{}\" xmlns:xr=\"{}\"",
            conformance.namespace(),
            XR_URI
        ))?;
        write_collection_attributes(&mut xml, collection)?;
        xml.push_str(">")?;
        for rule in &collection.validations {
            write_rule(&mut xml, rule)?;
        }
        xml.push_str("</dataValidations>")?;
    }
    Ok(xml.finish())
}

pub(crate) fn write_data_validation_extensions(
    values: &[DataValidationCollection],
    conformance: DataValidationConformance,
) -> Result<String> {
    validate_data_validation_collections(values)?;
    let mut xml = BoundedXml::new();
    if let Some(collection) = values
        .iter()
        .find(|collection| collection.source == DataValidationSource::Office2010)
    {
        xml.write_arguments(format_args!(
            "<extLst xmlns=\"{}\"><ext uri=\"{}\"><x14:dataValidations xmlns:x14=\"{}\" xmlns:xm=\"{}\" xmlns:x12ac=\"{}\" xmlns:xr=\"{}\"",
            conformance.namespace(),
            EXTENSION_URI,
            X14_URI,
            XM_URI,
            X12AC_URI,
            XR_URI,
        ))?;
        write_collection_attributes(&mut xml, collection)?;
        xml.push_str(">")?;
        for rule in &collection.validations {
            write_rule(&mut xml, rule)?;
        }
        xml.push_str("</x14:dataValidations></ext></extLst>")?;
    }
    Ok(xml.finish())
}

fn write_collection_attributes(
    xml: &mut BoundedXml,
    value: &DataValidationCollection,
) -> Result<()> {
    if value.disable_prompts {
        xml.push_str(" disablePrompts=\"1\"")?;
    }
    if let Some(value) = value.x_window {
        xml.write_arguments(format_args!(" xWindow=\"{value}\""))?;
    }
    if let Some(value) = value.y_window {
        xml.write_arguments(format_args!(" yWindow=\"{value}\""))?;
    }
    xml.write_arguments(format_args!(" count=\"{}\"", value.validations.len()))?;
    Ok(())
}

fn write_rule(xml: &mut BoundedXml, value: &ParsedDataValidation) -> Result<()> {
    validate_rule(value)?;
    let prefix = if value.source == DataValidationSource::Office2010 {
        "x14:"
    } else {
        ""
    };
    xml.write_arguments(format_args!("<{prefix}dataValidation"))?;
    if value.validation_type != ParsedDataValidationType::None {
        xml.write_arguments(format_args!(" type=\"{}\"", value.validation_type.as_str()))?;
    }
    if value.operator != ParsedDataValidationOperator::Between {
        xml.write_arguments(format_args!(" operator=\"{}\"", value.operator.as_str()))?;
    }
    if value.error_style != ParsedDataValidationErrorStyle::Stop {
        xml.write_arguments(format_args!(
            " errorStyle=\"{}\"",
            value.error_style.as_str()
        ))?;
    }
    if value.ime_mode != ParsedDataValidationImeMode::NoControl {
        xml.write_arguments(format_args!(" imeMode=\"{}\"", value.ime_mode.as_str()))?;
    }
    for (name, enabled) in [
        ("allowBlank", value.allow_blank),
        ("showDropDown", value.show_drop_down),
        ("showInputMessage", value.show_input_message),
        ("showErrorMessage", value.show_error_message),
    ] {
        if enabled {
            xml.write_arguments(format_args!(" {name}=\"1\""))?;
        }
    }
    for (name, text) in [
        ("errorTitle", value.error_title.as_deref()),
        ("error", value.error.as_deref()),
        ("promptTitle", value.prompt_title.as_deref()),
        ("prompt", value.prompt.as_deref()),
    ] {
        if let Some(text) = text {
            xml.write_arguments(format_args!(" {name}=\"{}\"", escape_xml(text)))?;
        }
    }
    if let Some(uid) = value.uid.as_deref() {
        xml.write_arguments(format_args!(" xr:uid=\"{}\"", escape_xml(uid)))?;
    }
    if value.source == DataValidationSource::Core {
        let sqref = sqref_text(&value.sqref)?;
        xml.write_arguments(format_args!(" sqref=\"{}\"", escape_xml(&sqref)))?;
    }
    xml.push_str(">")?;
    write_formula(xml, prefix, 1, value.formula1.as_ref())?;
    if let Some(formula) = value.formula2.as_ref() {
        write_formula_source(xml, prefix, 2, FormulaSource::Formula(&formula.0))?;
    }
    if value.source == DataValidationSource::Office2010 {
        xml.push_str("<xm:sqref")?;
        for (name, enabled) in [
            ("edited", value.sqref.edited),
            ("split", value.sqref.split),
            ("adjusted", value.sqref.adjusted),
            ("adjust", value.sqref.adjust),
        ] {
            if enabled {
                xml.write_arguments(format_args!(" {name}=\"1\""))?;
            }
        }
        let sqref = sqref_text(&value.sqref)?;
        xml.write_arguments(format_args!(">{}</xm:sqref>", escape_xml(&sqref)))?;
    }
    xml.write_arguments(format_args!("</{prefix}dataValidation>"))?;
    Ok(())
}

enum FormulaSource<'a> {
    Formula(&'a str),
    QuotedList(&'a str),
}

fn write_formula(
    xml: &mut BoundedXml,
    prefix: &str,
    number: u8,
    value: Option<&ValidationListSource>,
) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    let source = match value {
        ValidationListSource::Formula(value) => FormulaSource::Formula(&value.0),
        ValidationListSource::QuotedList(value) => FormulaSource::QuotedList(value),
    };
    write_formula_source(xml, prefix, number, source)
}

fn write_formula_source(
    xml: &mut BoundedXml,
    prefix: &str,
    number: u8,
    value: FormulaSource<'_>,
) -> Result<()> {
    xml.write_arguments(format_args!("<{prefix}formula{number}>"))?;
    match (prefix.is_empty(), value) {
        (true, FormulaSource::Formula(value)) => {
            xml.push_str(&escape_xml(value))?;
        },
        (true, FormulaSource::QuotedList(_)) => {
            return Err(invalid(
                "quoted-list source requires Office 2010 data validation",
            ));
        },
        (false, FormulaSource::Formula(value)) => {
            xml.write_arguments(format_args!("<xm:f>{}</xm:f>", escape_xml(value)))?;
        },
        (false, FormulaSource::QuotedList(value)) => {
            xml.write_arguments(format_args!(
                "<x12ac:list>{}</x12ac:list>",
                escape_xml(value)
            ))?;
        },
    }
    xml.write_arguments(format_args!("</{prefix}formula{number}>"))?;
    Ok(())
}

#[derive(Debug)]
struct DataValidationXmlScan {
    conformance: DataValidationConformance,
    worksheet_close: usize,
    core_insert: usize,
    core_ranges: Vec<Range<usize>>,
    x14_ranges: Vec<Range<usize>>,
    matching_ext_close: Option<usize>,
    ext_lst_close: Option<usize>,
}

/// Replace data-validation XML while preserving every unrelated worksheet byte.
pub fn replace_data_validation_collections(
    worksheet_xml: &[u8],
    values: &[DataValidationCollection],
) -> Result<Vec<u8>> {
    let parsed = parse_data_validation_collections(worksheet_xml)?;
    validate_data_validation_collections(&parsed)?;
    validate_data_validation_collections(values)?;
    let scan = scan_data_validation_xml(worksheet_xml)?;
    let parsed_core = parsed
        .iter()
        .any(|value| value.source == DataValidationSource::Core);
    let parsed_x14 = parsed
        .iter()
        .any(|value| value.source == DataValidationSource::Office2010);
    if parsed_core == scan.core_ranges.is_empty() || parsed_x14 == scan.x14_ranges.is_empty() {
        return Err(invalid(
            "data validations selected through MCE cannot be mutated byte-exactly",
        ));
    }
    let core = write_data_validation_core(values, scan.conformance)?;
    let extensions = write_data_validation_extensions(values, scan.conformance)?;
    let edit_count = scan
        .core_ranges
        .len()
        .checked_add(scan.x14_ranges.len())
        .ok_or_else(|| invalid("data-validation edit count overflow"))?;
    let mut edits = Vec::new();
    reserve_vec(&mut edits, edit_count, "data-validation edits")?;
    edits.extend(
        scan.core_ranges
            .iter()
            .chain(scan.x14_ranges.iter())
            .cloned()
            .map(|range| (range, Vec::new())),
    );
    if !core.is_empty() {
        if let Some(range) = scan.core_ranges.first() {
            let Some(edit) = edits.iter_mut().find(|(candidate, _)| candidate == range) else {
                return Err(invalid("missing core data-validation edit"));
            };
            edit.1 = core.into_bytes();
        } else {
            reserve_vec(&mut edits, 1, "data-validation edits")?;
            edits.push((scan.core_insert..scan.core_insert, core.into_bytes()));
        }
    }
    if !extensions.is_empty() {
        let inner = data_validation_extension_inner(&extensions)?;
        if let Some(range) = scan.x14_ranges.first() {
            let Some(edit) = edits.iter_mut().find(|(candidate, _)| candidate == range) else {
                return Err(invalid("missing Office 2010 data-validation edit"));
            };
            edit.1 = inner.into_bytes();
        } else if let Some(position) = scan.matching_ext_close {
            reserve_vec(&mut edits, 1, "data-validation edits")?;
            edits.push((position..position, inner.into_bytes()));
        } else if let Some(position) = scan.ext_lst_close {
            reserve_vec(&mut edits, 1, "data-validation edits")?;
            edits.push((
                position..position,
                data_validation_extension_wrapper(&inner, scan.conformance)?.into_bytes(),
            ));
        } else {
            reserve_vec(&mut edits, 1, "data-validation edits")?;
            edits.push((
                scan.worksheet_close..scan.worksheet_close,
                extensions.into_bytes(),
            ));
        }
    }
    apply_data_validation_edits(worksheet_xml, edits)
}

fn data_validation_extension_inner(fragment: &str) -> Result<String> {
    let start = fragment
        .find("<x14:dataValidations")
        .ok_or_else(|| invalid("invalid generated data-validation extension"))?;
    let end = fragment
        .rfind("</x14:dataValidations>")
        .ok_or_else(|| invalid("invalid generated data-validation extension"))?
        + "</x14:dataValidations>".len();
    Ok(fragment[start..end].to_string())
}

fn data_validation_extension_wrapper(
    inner: &str,
    conformance: DataValidationConformance,
) -> Result<String> {
    let mut wrapper = BoundedXml::new();
    wrapper.write_arguments(format_args!(
        "<ext xmlns=\"{}\" uri=\"{}\">{inner}</ext>",
        conformance.namespace(),
        EXTENSION_URI
    ))?;
    Ok(wrapper.finish())
}

fn apply_data_validation_edits(
    xml: &[u8],
    mut edits: Vec<(Range<usize>, Vec<u8>)>,
) -> Result<Vec<u8>> {
    edits.sort_by_key(|(range, _)| (range.start, range.end));
    let mut output = Vec::new();
    reserve_vec(&mut output, xml.len(), "data-validation XML output")?;
    let mut cursor = 0usize;
    for (range, replacement) in edits {
        if range.start < cursor || range.end < range.start || range.end > xml.len() {
            return Err(invalid("overlapping data-validation XML edits"));
        }
        append_bounded_bytes(&mut output, &xml[cursor..range.start])?;
        append_bounded_bytes(&mut output, &replacement)?;
        cursor = range.end;
    }
    append_bounded_bytes(&mut output, &xml[cursor..])?;
    let reparsed = parse_data_validation_collections(&output)?;
    validate_data_validation_collections(&reparsed)?;
    Ok(output)
}

fn push_scan_range(ranges: &mut Vec<Range<usize>>, range: Range<usize>) -> Result<()> {
    if ranges.len() >= MAX_CAPTURED_COLLECTIONS {
        return Err(invalid("too many physical data-validation collections"));
    }
    reserve_vec(ranges, 1, "data-validation scan ranges")?;
    ranges.push(range);
    Ok(())
}

fn scan_data_validation_xml(xml: &[u8]) -> Result<DataValidationXmlScan> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("data-validation worksheet XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut previous = 0usize;
    let mut conformance = None;
    let mut worksheet_close = None;
    let mut core_insert = None;
    let mut core_start = None;
    let mut core_ranges = Vec::new();
    let mut x14_start = None;
    let mut x14_ranges = Vec::new();
    let mut matching_ext_depth = None;
    let mut matching_ext_close = None;
    let mut ext_lst_depth = None;
    let mut ext_lst_close = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut events = 0usize;
    let mut nodes = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("data-validation worksheet event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("data-validation worksheet exceeds event limit"));
        }
        let start = previous;
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("data-validation XML offset overflow"))?;
        previous = end;
        let decoder = reader.decoder();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if matches!(&event, Event::Eof) {
            break;
        }
        if matches!(&event, Event::Start(_) | Event::Empty(_) | Event::End(_)) {
            nodes = nodes
                .checked_add(1)
                .ok_or_else(|| invalid("data-validation worksheet node count overflow"))?;
            if nodes > MAX_NODES {
                return Err(invalid("data-validation worksheet exceeds node limit"));
            }
        }
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if depth == 0 && root_seen {
                    return Err(invalid("worksheet XML contains multiple roots"));
                }
                if depth >= MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                let local = element.local_name();
                if depth == 1 {
                    conformance = match namespace {
                        ResolveResult::Bound(value) if value.as_ref() == CORE => {
                            Some(DataValidationConformance::Transitional)
                        },
                        ResolveResult::Bound(value) if value.as_ref() == STRICT => {
                            Some(DataValidationConformance::Strict)
                        },
                        _ => None,
                    };
                    if conformance.is_none() || local.as_ref() != b"worksheet" {
                        return Err(invalid("invalid worksheet namespace"));
                    }
                    root_seen = true;
                } else if depth == 2 {
                    if local.as_ref() == b"dataValidations" && !spreadsheet(&namespace) {
                        return Err(invalid("spoofed dataValidations element namespace"));
                    }
                    if spreadsheet(&namespace) {
                        if local.as_ref() == b"dataValidations" {
                            if core_start.is_some() {
                                return Err(invalid("duplicate core dataValidations element"));
                            }
                            core_start = Some((depth, start));
                        } else if core_insert.is_none() && validation_schema_after(local.as_ref()) {
                            core_insert = Some(start);
                        }
                        if local.as_ref() == b"extLst" {
                            ext_lst_depth = Some(depth);
                        }
                    }
                }
                if spreadsheet(&namespace)
                    && local.as_ref() == b"ext"
                    && optional_attr(&element, b"uri", decoder)?.as_deref() == Some(EXTENSION_URI)
                {
                    if matching_ext_depth.is_some() {
                        return Err(invalid("nested data-validation extension"));
                    }
                    matching_ext_depth = Some(depth);
                }
                if local.as_ref() == b"dataValidations" && matching_ext_depth.is_some() {
                    if !exact(&namespace, X14) {
                        return Err(invalid("spoofed x14 dataValidations element namespace"));
                    }
                    if x14_start.is_some() {
                        return Err(invalid("duplicate Office 2010 dataValidations element"));
                    }
                    x14_start = Some((depth, start));
                }
            },
            Event::Empty(element) => {
                if root_closed || depth == 0 {
                    return Err(invalid("worksheet XML contains an element outside root"));
                }
                let local = element.local_name();
                let element_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML nesting overflow"))?;
                if element_depth > MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                if element_depth == 2 {
                    if local.as_ref() == b"dataValidations" && !spreadsheet(&namespace) {
                        return Err(invalid("spoofed dataValidations element namespace"));
                    }
                    if spreadsheet(&namespace) && local.as_ref() == b"dataValidations" {
                        if core_start.is_some() {
                            return Err(invalid("duplicate core dataValidations element"));
                        }
                        push_scan_range(&mut core_ranges, start..end)?;
                    } else if spreadsheet(&namespace)
                        && core_insert.is_none()
                        && validation_schema_after(local.as_ref())
                    {
                        core_insert = Some(start);
                    }
                }
                if local.as_ref() == b"dataValidations" && matching_ext_depth.is_some() {
                    if !exact(&namespace, X14) {
                        return Err(invalid("spoofed x14 dataValidations element namespace"));
                    }
                    push_scan_range(&mut x14_ranges, start..end)?;
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("unexpected worksheet closing element"));
                }
                if core_start.is_some_and(|(element_depth, _)| element_depth == depth) {
                    let Some((_, range_start)) = core_start.take() else {
                        return Err(invalid("missing core data-validation scan state"));
                    };
                    push_scan_range(&mut core_ranges, range_start..end)?;
                }
                if x14_start.is_some_and(|(element_depth, _)| element_depth == depth) {
                    let Some((_, range_start)) = x14_start.take() else {
                        return Err(invalid("missing Office 2010 data-validation scan state"));
                    };
                    push_scan_range(&mut x14_ranges, range_start..end)?;
                }
                if matching_ext_depth == Some(depth) {
                    matching_ext_close = Some(start);
                    matching_ext_depth = None;
                }
                if ext_lst_depth == Some(depth) {
                    ext_lst_close = Some(start);
                    ext_lst_depth = None;
                }
                if depth == 1 && element.local_name().as_ref() == b"worksheet" {
                    if !spreadsheet(&namespace) {
                        return Err(invalid("invalid worksheet closing namespace"));
                    }
                    root_closed = true;
                    worksheet_close = Some(start);
                } else if depth == 1 {
                    return Err(invalid("invalid worksheet closing element"));
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("worksheet XML nesting underflow"))?;
            },
            Event::Text(value) => {
                if (!root_seen || root_closed)
                    && !value.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("worksheet XML text is outside root"));
                }
                if depth == 1 && !value.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("worksheet cannot contain direct text"));
                }
            },
            Event::CData(_) => {
                return Err(invalid("worksheet XML contains unexpected CDATA"));
            },
            Event::GeneralRef(reference) => {
                decode_xml_reference(&reference)?;
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen {
                    return Err(invalid("invalid worksheet XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !root_seen || !root_closed || depth != 0 || core_start.is_some() || x14_start.is_some() {
        return Err(invalid("incomplete worksheet data-validation XML"));
    }
    let worksheet_close = worksheet_close.ok_or_else(|| invalid("worksheet is not closed"))?;
    Ok(DataValidationXmlScan {
        conformance: conformance.ok_or_else(|| invalid("invalid worksheet namespace"))?,
        worksheet_close,
        core_insert: core_insert.unwrap_or(worksheet_close),
        core_ranges,
        x14_ranges,
        matching_ext_close,
        ext_lst_close,
    })
}

fn validation_schema_after(local: &[u8]) -> bool {
    matches!(
        local,
        b"hyperlinks"
            | b"printOptions"
            | b"pageMargins"
            | b"pageSetup"
            | b"headerFooter"
            | b"rowBreaks"
            | b"colBreaks"
            | b"customProperties"
            | b"cellWatches"
            | b"ignoredErrors"
            | b"smartTags"
            | b"drawing"
            | b"legacyDrawing"
            | b"legacyDrawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    fn sheet(relative: &str, index: u32) -> Vec<DataValidationCollection> {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let package = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(root.join(relative))
            .unwrap_or_else(|e| panic!("open {relative}: {e}"));
        let uri = PackURI::new(format!("/xl/worksheets/sheet{index}.xml")).unwrap();
        let bytes = package
            .blob_for(&uri)
            .unwrap_or_else(|e| panic!("sheet {relative}: {e}"));
        parse_data_validation_collections(&bytes)
            .unwrap_or_else(|e| panic!("parse {relative}: {e}"))
    }
    fn count(values: &[DataValidationCollection]) -> usize {
        values.iter().map(|v| v.validations.len()).sum()
    }

    #[test]
    fn parses_poi_and_libreoffice_fixtures() {
        let cases = [
            (
                "test-data/poi/test-data/spreadsheet/DataValidationListTooLong.xlsx",
                1usize,
            ),
            (
                "test-data/poi/test-data/spreadsheet/DataValidations-49244.xlsx",
                52,
            ),
            (
                "test-data/poi/test-data/spreadsheet/dataValidationTableRange.xlsx",
                5,
            ),
            (
                "test-data/poi/test-data/spreadsheet/DataValidationEvaluations.xlsx",
                17,
            ),
            (
                "test-data/libreoffice-core/sc/qa/unit/data/xlsx/textLengthDataValidity.xlsx",
                1,
            ),
            (
                "test-data/libreoffice-core/sc/qa/unit/data/xlsx/invalid_ext_data_validation.xlsx",
                1,
            ),
            (
                "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dataValidity.xlsx",
                1,
            ),
        ];
        let mut parsed = Vec::new();
        for (case, expected) in cases {
            let values = sheet(case, 1);
            assert_eq!(count(&values), expected, "{case}");
            parsed.push(values);
        }
        let extension = &parsed[5][0].validations[0];
        assert_eq!(extension.source, DataValidationSource::Office2010);
        assert_eq!(extension.sqref.ranges[0].as_str(), "F6");
        assert!(
            matches!(&extension.formula1,Some(ValidationListSource::Formula(v))if v.as_str()=="[2]Tabelle1!#REF!")
        );
        let text = &parsed[4][0].validations[0];
        assert_eq!(text.validation_type, ParsedDataValidationType::TextLength);
        assert_eq!(
            text.uid.as_deref(),
            Some("{3FE27F7A-BE41-432D-B94C-05DA7B860A0B}")
        );
        assert!(
            matches!(&parsed[0][0].validations[0].formula1,Some(ValidationListSource::Formula(v))if v.as_str().len()>255)
        );
    }

    #[test]
    fn parses_strict_mce_and_quoted_list() {
        let xml=br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:x12ac="http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac"><dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations><mc:AlternateContent><mc:Choice Requires="x14"><extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1" disablePrompts="1" xWindow="2" yWindow="3"><x14:dataValidation type="list"><x14:formula1><x12ac:list>&quot;a,b&quot;</x12ac:list></x14:formula1><xm:sqref adjusted="1" adjust="1">B2 C3:C4</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></mc:Choice><mc:Fallback><dataValidations count="1"><dataValidation type="none" sqref="D4"/></dataValidations></mc:Fallback></mc:AlternateContent></worksheet>"#;
        let values = parse_data_validation_collections(xml).unwrap();
        assert_eq!(values.len(), 2);
        assert!(values[1].disable_prompts);
        assert_eq!(values[1].x_window, Some(2));
        assert!(
            matches!(&values[1].validations[0].formula1,Some(ValidationListSource::QuotedList(v))if v=="\"a,b\"")
        );
        assert_eq!(values[1].validations[0].sqref.ranges.len(), 2);
    }

    #[test]
    fn rejects_malformed_and_dangerous_input() {
        let bad = [
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="2"><dataValidation type="none" sqref="A1"/></dataValidations></worksheet>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="whole" sqref="A0"><formula1>1</formula1><formula2>2</formula2></dataValidation></dataValidations></worksheet>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula2>2</formula2><formula1>1</formula1></dataValidation></dataValidations></worksheet>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>x</formula1><formula2>y</formula2></dataValidation></dataValidations></worksheet>"#,
            r#"<!DOCTYPE x><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        ];
        for xml in bad {
            assert!(
                parse_data_validation_collections(xml.as_bytes()).is_err(),
                "{xml}"
            );
        }
        for attributes in [
            "type=\"integer\"",
            "allowBlank=\"TRUE\"",
            "operator=\"near\"",
            "errorStyle=\"fatal\"",
            "imeMode=\"automatic\"",
            "xr:uid=\"not-a-guid\"",
        ] {
            let xml = format!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><dataValidations count="1"><dataValidation {attributes} sqref="A1"/></dataValidations></worksheet>"#
            );
            assert!(
                parse_data_validation_collections(xml.as_bytes()).is_err(),
                "{attributes}"
            );
        }
        let long_title = "x".repeat(33);
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="none" errorTitle="{long_title}" sqref="A1"/></dataValidations></worksheet>"#
        );
        assert!(parse_data_validation_collections(xml.as_bytes()).is_err());
        let long_formula = "x".repeat(MAX_FORMULA_BYTES + 1);
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>{long_formula}</formula1></dataValidation></dataValidations></worksheet>"#
        );
        assert!(parse_data_validation_collections(xml.as_bytes()).is_err());
        assert!(parse_data_validation_collections(br#"<?bad x?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#).is_err());
        assert!(parse_data_validation_collections(br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>&bogus;</formula1></dataValidation></dataValidations></worksheet>"#).is_err());
        let wrong_uri=br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><extLst><ext uri="wrong"><x14:dataValidations count="1"><x14:dataValidation type="none"><xm:sqref>A1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"#;
        assert!(
            parse_data_validation_collections(wrong_uri)
                .unwrap()
                .is_empty()
        );
        let bad_flags=br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1"><x14:dataValidation type="none"><xm:sqref adjusted="1">A1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"#;
        assert!(parse_data_validation_collections(bad_flags).is_err());
        let too_many = (0..=MAX_REFERENCES)
            .map(|_| "A1")
            .collect::<Vec<_>>()
            .join(" ");
        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dataValidations count="1"><dataValidation type="none" sqref="{too_many}"/></dataValidations></worksheet>"#
        );
        assert!(parse_data_validation_collections(xml.as_bytes()).is_err());
    }
}
