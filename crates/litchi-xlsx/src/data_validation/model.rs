//! Semantic `SpreadsheetML` data-validation declarations.
///
use super::codec::{
    invalid, parse_sqref, validate_collection, validate_optional_text, validate_rule, validate_text,
};
use super::{CORE_URI, MAX_FORMULA_BYTES, STRICT_URI};
use crate::error::Result;
use litchi_ooxml_common::custom_xml::valid_guid;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => CORE_URI,
            Self::Strict => STRICT_URI,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Source {
    Core,
    Office2010,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValidationType {
    None,
    Whole,
    Decimal,
    List,
    Date,
    Time,
    TextLength,
    Custom,
}

impl ValidationType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
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
    pub(crate) fn as_str(self) -> &'static str {
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
pub enum ValidationOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
}

impl ValidationOperator {
    pub(crate) fn parse(value: &str) -> Result<Self> {
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
    pub(crate) fn as_str(self) -> &'static str {
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
pub enum ValidationErrorStyle {
    Stop,
    Warning,
    Information,
}

impl ValidationErrorStyle {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "stop" => Ok(Self::Stop),
            "warning" => Ok(Self::Warning),
            "information" => Ok(Self::Information),
            _ => Err(invalid(format!(
                "invalid data-validation errorStyle '{value}'"
            ))),
        }
    }
    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Stop => "stop",
            Self::Warning => "warning",
            Self::Information => "information",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValidationImeMode {
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

impl ValidationImeMode {
    pub(crate) fn parse(value: &str) -> Result<Self> {
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
    pub(crate) fn as_str(self) -> &'static str {
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
pub struct Range(pub(crate) String);
impl Range {
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

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Formula(pub(crate) String);
impl Formula {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_text(&value, MAX_FORMULA_BYTES, "data-validation formula")?;
        Ok(Self(value))
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ListSource {
    Formula(Formula),
    QuotedList(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sqref {
    pub(crate) ranges: Vec<Range>,
    pub(crate) edited: bool,
    pub(crate) split: bool,
    pub(crate) adjusted: bool,
    pub(crate) adjust: bool,
}
impl Sqref {
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

    #[must_use]
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }
    #[must_use]
    pub fn edited(&self) -> bool {
        self.edited
    }
    #[must_use]
    pub fn split(&self) -> bool {
        self.split
    }
    #[must_use]
    pub fn adjusted(&self) -> bool {
        self.adjusted
    }
    #[must_use]
    pub fn adjust(&self) -> bool {
        self.adjust
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Validation {
    pub(crate) source: Source,
    pub(crate) validation_type: ValidationType,
    pub(crate) operator: ValidationOperator,
    pub(crate) error_style: ValidationErrorStyle,
    pub(crate) ime_mode: ValidationImeMode,
    pub(crate) allow_blank: bool,
    pub(crate) show_drop_down: bool,
    pub(crate) show_input_message: bool,
    pub(crate) show_error_message: bool,
    pub(crate) error_title: Option<String>,
    pub(crate) error: Option<String>,
    pub(crate) prompt_title: Option<String>,
    pub(crate) prompt: Option<String>,
    pub(crate) formula1: Option<ListSource>,
    pub(crate) formula2: Option<Formula>,
    pub(crate) sqref: Sqref,
    pub(crate) uid: Option<String>,
}

impl Validation {
    #[must_use]
    pub fn new(source: Source, validation_type: ValidationType, sqref: Sqref) -> Self {
        Self {
            source,
            validation_type,
            operator: ValidationOperator::Between,
            error_style: ValidationErrorStyle::Stop,
            ime_mode: ValidationImeMode::NoControl,
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

    pub fn set_operator(&mut self, value: ValidationOperator) {
        self.operator = value;
    }
    pub fn set_error_style(&mut self, value: ValidationErrorStyle) {
        self.error_style = value;
    }
    pub fn set_ime_mode(&mut self, value: ValidationImeMode) {
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
    pub fn set_formula1(&mut self, value: Option<ListSource>) -> Result<()> {
        if let Some(ListSource::Formula(value)) = value.as_ref() {
            validate_text(&value.0, MAX_FORMULA_BYTES, "formula1")?;
        }
        if let Some(ListSource::QuotedList(value)) = value.as_ref() {
            validate_text(value, MAX_FORMULA_BYTES, "quoted validation list")?;
            if self.source != Source::Office2010 {
                return Err(invalid(
                    "quoted-list source requires Office 2010 data validation",
                ));
            }
        }
        self.formula1 = value;
        Ok(())
    }
    pub fn set_formula2(&mut self, value: Option<Formula>) -> Result<()> {
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

    #[must_use]
    pub fn source(&self) -> Source {
        self.source
    }
    #[must_use]
    pub fn validation_type(&self) -> ValidationType {
        self.validation_type
    }
    #[must_use]
    pub fn operator(&self) -> ValidationOperator {
        self.operator
    }
    #[must_use]
    pub fn error_style(&self) -> ValidationErrorStyle {
        self.error_style
    }
    #[must_use]
    pub fn ime_mode(&self) -> ValidationImeMode {
        self.ime_mode
    }
    #[must_use]
    pub fn allow_blank(&self) -> bool {
        self.allow_blank
    }
    #[must_use]
    pub fn show_drop_down(&self) -> bool {
        self.show_drop_down
    }
    #[must_use]
    pub fn show_input_message(&self) -> bool {
        self.show_input_message
    }
    #[must_use]
    pub fn show_error_message(&self) -> bool {
        self.show_error_message
    }
    #[must_use]
    pub fn error_title(&self) -> Option<&str> {
        self.error_title.as_deref()
    }
    #[must_use]
    pub fn error(&self) -> Option<&str> {
        self.error.as_deref()
    }
    #[must_use]
    pub fn prompt_title(&self) -> Option<&str> {
        self.prompt_title.as_deref()
    }
    #[must_use]
    pub fn prompt(&self) -> Option<&str> {
        self.prompt.as_deref()
    }
    #[must_use]
    pub fn formula1(&self) -> Option<&ListSource> {
        self.formula1.as_ref()
    }
    #[must_use]
    pub fn formula2(&self) -> Option<&Formula> {
        self.formula2.as_ref()
    }
    #[must_use]
    pub fn sqref(&self) -> &Sqref {
        &self.sqref
    }
    #[must_use]
    pub fn uid(&self) -> Option<&str> {
        self.uid.as_deref()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Collection {
    pub(crate) source: Source,
    pub(crate) disable_prompts: bool,
    pub(crate) x_window: Option<u32>,
    pub(crate) y_window: Option<u32>,
    pub(crate) declared_count: Option<u32>,
    pub(crate) validations: Vec<Validation>,
}
impl Collection {
    pub fn new(source: Source, validations: Vec<Validation>) -> Result<Self> {
        let declared_count = u32::try_from(validations.len())
            .map_err(|_source| invalid("data-validation count exceeds u32"))?;
        let value = Self {
            source,
            disable_prompts: false,
            x_window: None,
            y_window: None,
            declared_count: Some(declared_count),
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
    pub fn set_validations(&mut self, validations: Vec<Validation>) -> Result<()> {
        let declared_count = u32::try_from(validations.len())
            .map_err(|_source| invalid("data-validation count exceeds u32"))?;
        let candidate = Self {
            source: self.source,
            disable_prompts: self.disable_prompts,
            x_window: self.x_window,
            y_window: self.y_window,
            declared_count: Some(declared_count),
            validations,
        };
        validate_collection(&candidate)?;
        *self = candidate;
        Ok(())
    }

    #[must_use]
    pub fn source(&self) -> Source {
        self.source
    }
    #[must_use]
    pub fn disable_prompts(&self) -> bool {
        self.disable_prompts
    }
    #[must_use]
    pub fn x_window(&self) -> Option<u32> {
        self.x_window
    }
    #[must_use]
    pub fn y_window(&self) -> Option<u32> {
        self.y_window
    }
    #[must_use]
    pub fn declared_count(&self) -> Option<u32> {
        self.declared_count
    }
    #[must_use]
    pub fn validations(&self) -> &[Validation] {
        &self.validations
    }
}
