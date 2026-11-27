use super::tokens::{
    attr, bool_attr, esc, locale_attrs, validate_part, validate_sequence, write_part,
};
use super::{
    MAX_MAPS, MAX_PARTS, MAX_STYLES, Result, invalid, validate_cell_address, validate_locale,
    validate_name, validate_optional_string, validate_text,
};

/// XML part containing a data style.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Part {
    Content,
    Styles,
    Flat,
}

/// Direct style container containing a data style.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Section {
    Styles,
    AutomaticStyles,
}

/// Core schema version used to validate or serialize a style.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Version {
    V1_2,
    V1_3,
}

/// One of the seven standard data-style containers.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Kind {
    Number,
    Currency,
    Percentage,
    Date,
    Time,
    Boolean,
    Text,
}

impl Kind {
    pub(crate) fn local(self) -> &'static str {
        match self {
            Self::Number => "number-style",
            Self::Currency => "currency-style",
            Self::Percentage => "percentage-style",
            Self::Date => "date-style",
            Self::Time => "time-style",
            Self::Boolean => "boolean-style",
            Self::Text => "text-style",
        }
    }

    pub(crate) fn parse(local: &str) -> Option<Self> {
        Some(match local {
            "number-style" => Self::Number,
            "currency-style" => Self::Currency,
            "percentage-style" => Self::Percentage,
            "date-style" => Self::Date,
            "time-style" => Self::Time,
            "boolean-style" => Self::Boolean,
            "text-style" => Self::Text,
            _ => return None,
        })
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ShortLong {
    Short,
    Long,
}

impl ShortLong {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "short" => Ok(Self::Short),
            "long" => Ok(Self::Long),
            _ => invalid(format!("invalid number:style '{value}'")),
        }
    }

    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Short => "short",
            Self::Long => "long",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum TransliterationStyle {
    Short,
    Medium,
    Long,
}

impl TransliterationStyle {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "short" => Ok(Self::Short),
            "medium" => Ok(Self::Medium),
            "long" => Ok(Self::Long),
            _ => invalid(format!("invalid number:transliteration-style '{value}'")),
        }
    }

    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Short => "short",
            Self::Medium => "medium",
            Self::Long => "long",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum FormatSource {
    Fixed,
    Language,
}

impl FormatSource {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "fixed" => Ok(Self::Fixed),
            "language" => Ok(Self::Language),
            _ => invalid(format!("invalid number:format-source '{value}'")),
        }
    }

    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Fixed => "fixed",
            Self::Language => "language",
        }
    }
}

/// Locale metadata shared by a data style or currency symbol.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Locale {
    pub language: Option<String>,
    pub country: Option<String>,
    pub script: Option<String>,
    pub rfc_language_tag: Option<String>,
}

/// A validated, opaque `style:text-properties` fragment.
///
/// Its independent typed property API remains the authority for the 84
/// text-property attributes; the data-style parser preserves the exact fragment.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextProperties {
    pub(crate) xml: String,
}

impl TextProperties {
    #[must_use]
    pub fn as_xml(&self) -> &str {
        &self.xml
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct NumberToken {
    pub decimal_replacement: Option<String>,
    pub display_factor: Option<f64>,
    pub decimal_places: Option<i64>,
    pub min_decimal_places: Option<i64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
    pub embedded_text: Vec<EmbeddedText>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedText {
    pub position: i64,
    pub text: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Scientific {
    pub min_exponent_digits: Option<i64>,
    pub exponent_interval: Option<u64>,
    pub forced_exponent_sign: Option<bool>,
    pub decimal_places: Option<i64>,
    pub min_decimal_places: Option<i64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Fraction {
    pub min_numerator_digits: Option<i64>,
    pub min_denominator_digits: Option<i64>,
    pub denominator_value: Option<i64>,
    pub max_denominator_value: Option<u64>,
    pub min_integer_digits: Option<i64>,
    pub grouping: Option<bool>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Currency {
    pub locale: Locale,
    pub text: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Calendar {
    pub style: Option<ShortLong>,
    pub calendar: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Month {
    pub style: Option<ShortLong>,
    pub textual: Option<bool>,
    pub possessive_form: Option<bool>,
    pub calendar: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct WeekOfYear {
    pub calendar: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Clock {
    pub style: Option<ShortLong>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Seconds {
    pub style: Option<ShortLong>,
    pub decimal_places: Option<i64>,
}

/// One ordered formatting component within a data style.
#[derive(Clone, Debug, PartialEq)]
pub enum Token {
    Text(String),
    FillCharacter(String),
    Number(NumberToken),
    ScientificNumber(Scientific),
    Fraction(Fraction),
    CurrencySymbol(Currency),
    Day(Calendar),
    Month(Month),
    Year(Calendar),
    Era(Calendar),
    DayOfWeek(Calendar),
    WeekOfYear(WeekOfYear),
    Quarter(Calendar),
    Hours(Clock),
    Minutes(Clock),
    Seconds(Seconds),
    AmPm,
    Boolean,
    TextContent,
}

/// A trailing conditional style map. Conditions remain opaque strings.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Map {
    pub condition: String,
    pub apply_style_name: String,
    pub base_cell_address: Option<String>,
}

/// One complete standard data style.
#[derive(Clone, Debug, PartialEq)]
pub struct Style {
    pub source_part: Part,
    pub section: Section,
    pub source_version: Version,
    pub kind: Kind,
    pub name: String,
    pub display_name: Option<String>,
    pub locale: Locale,
    pub title: Option<String>,
    pub volatile: Option<bool>,
    pub transliteration_format: Option<String>,
    pub transliteration_language: Option<String>,
    pub transliteration_country: Option<String>,
    pub transliteration_style: Option<TransliterationStyle>,
    pub automatic_order: Option<bool>,
    pub format_source: Option<FormatSource>,
    pub truncate_on_overflow: Option<bool>,
    pub text_properties: Option<TextProperties>,
    pub parts: Vec<Token>,
    pub maps: Vec<Map>,
}

impl Style {
    /// Create an empty data style with a validated identity.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied style name is not valid XML text.
    pub fn new(name: impl Into<String>, kind: Kind, section: Section) -> Result<Self> {
        let value = Self {
            source_part: Part::Flat,
            section,
            source_version: Version::V1_3,
            kind,
            name: name.into(),
            display_name: None,
            locale: Locale::default(),
            title: None,
            volatile: None,
            transliteration_format: None,
            transliteration_language: None,
            transliteration_country: None,
            transliteration_style: None,
            automatic_order: None,
            format_source: None,
            truncate_on_overflow: None,
            text_properties: None,
            parts: Vec::new(),
            maps: Vec::new(),
        };
        value.validate(Version::V1_3)?;
        Ok(value)
    }

    /// Validate this style for the requested ODF version.
    ///
    /// # Errors
    ///
    /// Returns an error when an attribute, token, map, or ordering rule is not
    /// valid for the requested version.
    pub fn validate(&self, version: Version) -> Result<()> {
        self.validate_inner(version, false)
    }

    pub(crate) fn validate_inner(&self, version: Version, allow_lo_aliases: bool) -> Result<()> {
        validate_name(&self.name, "style:name")?;
        validate_optional_string(self.display_name.as_deref(), "style:display-name")?;
        validate_locale(&self.locale)?;
        for (value, name) in [
            (self.title.as_deref(), "number:title"),
            (
                self.transliteration_format.as_deref(),
                "number:transliteration-format",
            ),
            (
                self.transliteration_language.as_deref(),
                "number:transliteration-language",
            ),
            (
                self.transliteration_country.as_deref(),
                "number:transliteration-country",
            ),
        ] {
            validate_optional_string(value, name)?;
        }
        if self.parts.len() > MAX_PARTS || self.maps.len() > MAX_MAPS {
            return invalid("data style exceeds component limits");
        }
        if self.automatic_order.is_some() && !matches!(self.kind, Kind::Currency | Kind::Date) {
            return invalid("number:automatic-order is invalid for this data style");
        }
        if self.format_source.is_some() && !matches!(self.kind, Kind::Date | Kind::Time) {
            return invalid("number:format-source is invalid for this data style");
        }
        if self.truncate_on_overflow.is_some() && self.kind != Kind::Time {
            return invalid("number:truncate-on-overflow is valid only for time styles");
        }
        validate_sequence(self.kind, &self.parts, version, allow_lo_aliases)?;
        for part in &self.parts {
            validate_part(part, version, allow_lo_aliases)?;
        }
        for map in &self.maps {
            validate_text(&map.condition, "style:condition")?;
            validate_name(&map.apply_style_name, "style:apply-style-name")?;
            if let Some(address) = &map.base_cell_address {
                validate_cell_address(address)?;
            }
        }
        Ok(())
    }

    /// Serialize a normative compact ODF fragment for the requested core version.
    ///
    /// # Errors
    ///
    /// Returns an error when the style is not valid for the requested version.
    pub fn to_xml_fragment(&self, version: Version) -> Result<String> {
        self.validate(version)?;
        let mut out = format!(
            "<number:{} style:name=\"{}\"",
            self.kind.local(),
            esc(&self.name)
        );
        attr(&mut out, "style:display-name", self.display_name.as_deref());
        locale_attrs(&mut out, &self.locale);
        attr(&mut out, "number:title", self.title.as_deref());
        bool_attr(&mut out, "style:volatile", self.volatile);
        attr(
            &mut out,
            "number:transliteration-format",
            self.transliteration_format.as_deref(),
        );
        attr(
            &mut out,
            "number:transliteration-language",
            self.transliteration_language.as_deref(),
        );
        attr(
            &mut out,
            "number:transliteration-country",
            self.transliteration_country.as_deref(),
        );
        if let Some(value) = self.transliteration_style {
            attr(
                &mut out,
                "number:transliteration-style",
                Some(value.as_str()),
            );
        }
        bool_attr(&mut out, "number:automatic-order", self.automatic_order);
        if let Some(value) = self.format_source {
            attr(&mut out, "number:format-source", Some(value.as_str()));
        }
        bool_attr(
            &mut out,
            "number:truncate-on-overflow",
            self.truncate_on_overflow,
        );
        out.push('>');
        if let Some(properties) = &self.text_properties {
            out.push_str(properties.as_xml());
        }
        for part in &self.parts {
            write_part(&mut out, part, version)?;
        }
        for map in &self.maps {
            out.push_str("<style:map");
            attr(&mut out, "style:condition", Some(&map.condition));
            attr(
                &mut out,
                "style:apply-style-name",
                Some(&map.apply_style_name),
            );
            attr(
                &mut out,
                "style:base-cell-address",
                map.base_cell_address.as_deref(),
            );
            out.push_str("/>");
        }
        out.push_str("</number:");
        out.push_str(self.kind.local());
        out.push('>');
        Ok(out)
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct Styles {
    pub styles: Vec<Style>,
}

impl Styles {
    #[must_use]
    pub fn get(&self, part: Part, section: Section, name: &str) -> Option<&Style> {
        self.styles.iter().find(|style| {
            style.source_part == part && style.section == section && style.name == name
        })
    }

    pub fn in_section(&self, section: Section) -> impl Iterator<Item = &Style> {
        self.styles
            .iter()
            .filter(move |style| style.section == section)
    }

    pub fn named<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a Style> {
        self.styles.iter().filter(move |style| style.name == name)
    }

    pub(crate) fn append(&mut self, mut other: Self) -> Result<()> {
        for style in other.styles.drain(..) {
            if self
                .get(style.source_part, style.section, &style.name)
                .is_some()
            {
                return invalid("duplicate data style identity");
            }
            if self.styles.len() >= MAX_STYLES {
                return invalid("too many data styles");
            }
            self.styles.push(style);
        }
        Ok(())
    }
}
