//! Typed worksheet data-consolidation values and validation.

use crate::error::Result;

use super::{
    MAX_DATA_REFERENCES, MAX_RELATIONSHIP_ID_CHARS, MAX_XSTRING_CHARS, STRICT_MAIN, STRICT_REL,
    TRANSITIONAL_MAIN, TRANSITIONAL_REL, invalid,
};

/// Namespace form used when serializing a consolidation fragment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) fn main_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_MAIN,
            Self::Strict => STRICT_MAIN,
        }
    }

    pub(crate) fn relationship_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_REL,
            Self::Strict => STRICT_REL,
        }
    }
}

/// Mathematical aggregator selected by `ST_DataConsolidateFunction`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Function {
    Average,
    Count,
    CountNumbers,
    Maximum,
    Minimum,
    Product,
    StandardDeviation,
    PopulationStandardDeviation,
    #[default]
    Sum,
    Variance,
    PopulationVariance,
}

impl Function {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "average" => Ok(Self::Average),
            "count" => Ok(Self::Count),
            "countNums" => Ok(Self::CountNumbers),
            "max" => Ok(Self::Maximum),
            "min" => Ok(Self::Minimum),
            "product" => Ok(Self::Product),
            "stdDev" => Ok(Self::StandardDeviation),
            "stdDevp" => Ok(Self::PopulationStandardDeviation),
            "sum" => Ok(Self::Sum),
            "var" => Ok(Self::Variance),
            "varp" => Ok(Self::PopulationVariance),
            _ => Err(invalid(format!(
                "invalid dataConsolidate function {value:?}"
            ))),
        }
    }

    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Average => "average",
            Self::Count => "count",
            Self::CountNumbers => "countNums",
            Self::Maximum => "max",
            Self::Minimum => "min",
            Self::Product => "product",
            Self::StandardDeviation => "stdDev",
            Self::PopulationStandardDeviation => "stdDevp",
            Self::Sum => "sum",
            Self::Variance => "var",
            Self::PopulationVariance => "varp",
        }
    }
}

/// A validated A1 cell or rectangular range reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RangeReference(String);

impl RangeReference {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if !valid_range_reference(&value) {
            return Err(invalid(format!("invalid dataRef ref {value:?}")));
        }
        Ok(Self(value))
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// The mutually exclusive source forms of `CT_DataRef`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ReferenceSource {
    DefinedName(String),
    Range {
        sheet: String,
        reference: RangeReference,
    },
}

/// A single consolidation source, optionally in an external workbook.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    source: ReferenceSource,
    relationship_id: Option<String>,
}

impl Reference {
    pub fn named(name: impl Into<String>) -> Result<Self> {
        let name = checked_xstring(name.into(), "dataRef name")?;
        Ok(Self {
            source: ReferenceSource::DefinedName(name),
            relationship_id: None,
        })
    }

    pub fn range(sheet: impl Into<String>, reference: RangeReference) -> Result<Self> {
        let sheet = checked_xstring(sheet.into(), "dataRef sheet")?;
        Ok(Self {
            source: ReferenceSource::Range { sheet, reference },
            relationship_id: None,
        })
    }

    pub fn with_relationship_id(mut self, relationship_id: impl Into<String>) -> Result<Self> {
        self.relationship_id = Some(checked_relationship_id(relationship_id.into())?);
        Ok(self)
    }

    #[must_use]
    pub fn source(&self) -> &ReferenceSource {
        &self.source
    }

    #[must_use]
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    pub(crate) fn from_parts(source: ReferenceSource, relationship_id: Option<String>) -> Self {
        Self {
            source,
            relationship_id,
        }
    }
}

/// Bounded `dataRefs` collection and its optional source count.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct References {
    references: Vec<Reference>,
    declared_count: Option<u32>,
}

impl References {
    pub fn new(references: Vec<Reference>) -> Result<Self> {
        validate_reference_count(references.len())?;
        Ok(Self {
            declared_count: Some(references.len() as u32),
            references,
        })
    }

    #[must_use]
    pub fn references(&self) -> &[Reference] {
        &self.references
    }

    #[must_use]
    pub fn declared_count(&self) -> Option<u32> {
        self.declared_count
    }

    pub(crate) fn from_parts(references: Vec<Reference>, declared_count: Option<u32>) -> Self {
        Self {
            references,
            declared_count,
        }
    }
}

/// Complete immutable worksheet `dataConsolidate` settings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataConsolidation {
    function: Function,
    left_labels: bool,
    start_labels: bool,
    top_labels: bool,
    link: bool,
    data_references: Option<References>,
}

impl DataConsolidation {
    #[must_use]
    pub fn new(function: Function, data_references: Option<References>) -> Self {
        Self {
            function,
            left_labels: false,
            start_labels: false,
            top_labels: false,
            link: false,
            data_references,
        }
    }

    #[must_use]
    pub fn with_left_labels(mut self, value: bool) -> Self {
        self.left_labels = value;
        self
    }

    #[must_use]
    pub fn with_start_labels(mut self, value: bool) -> Self {
        self.start_labels = value;
        self
    }

    #[must_use]
    pub fn with_top_labels(mut self, value: bool) -> Self {
        self.top_labels = value;
        self
    }

    #[must_use]
    pub fn with_link(mut self, value: bool) -> Self {
        self.link = value;
        self
    }

    #[must_use]
    pub fn function(&self) -> Function {
        self.function
    }

    #[must_use]
    pub fn left_labels(&self) -> bool {
        self.left_labels
    }

    #[must_use]
    pub fn start_labels(&self) -> bool {
        self.start_labels
    }

    #[must_use]
    pub fn top_labels(&self) -> bool {
        self.top_labels
    }

    #[must_use]
    pub fn link(&self) -> bool {
        self.link
    }

    #[must_use]
    pub fn data_references(&self) -> Option<&References> {
        self.data_references.as_ref()
    }

    pub(crate) fn from_parts(
        function: Function,
        left_labels: bool,
        start_labels: bool,
        top_labels: bool,
        link: bool,
        data_references: Option<References>,
    ) -> Self {
        Self {
            function,
            left_labels,
            start_labels,
            top_labels,
            link,
            data_references,
        }
    }
}

pub(crate) fn checked_xstring(value: String, name: &str) -> Result<String> {
    if value.is_empty() {
        return Err(invalid(format!("{name} must not be empty")));
    }
    if value.chars().count() > MAX_XSTRING_CHARS {
        return Err(invalid(format!(
            "{name} exceeds {MAX_XSTRING_CHARS} characters"
        )));
    }
    Ok(value)
}

pub(crate) fn checked_relationship_id(value: String) -> Result<String> {
    if value.is_empty() || value.chars().count() > MAX_RELATIONSHIP_ID_CHARS {
        return Err(invalid("dataRef r:id is empty or exceeds the safety limit"));
    }
    Ok(value)
}

pub(crate) fn validate_reference_count(count: usize) -> Result<()> {
    if count > MAX_DATA_REFERENCES {
        Err(invalid(format!(
            "dataRefs exceeds safety limit {MAX_DATA_REFERENCES}"
        )))
    } else {
        Ok(())
    }
}

fn valid_range_reference(value: &str) -> bool {
    let mut parts = value.split(':');
    let Some(first) = parts.next() else {
        return false;
    };
    let second = parts.next();
    parts.next().is_none() && valid_cell_reference(first) && second.is_none_or(valid_cell_reference)
}

fn valid_cell_reference(value: &str) -> bool {
    let value = value.strip_prefix('$').unwrap_or(value);
    let letter_count = value.bytes().take_while(u8::is_ascii_alphabetic).count();
    if !(1..=3).contains(&letter_count) {
        return false;
    }
    let (letters, row) = value.split_at(letter_count);
    let row = row.strip_prefix('$').unwrap_or(row);
    if row.is_empty() || !row.bytes().all(|byte| byte.is_ascii_digit()) || row.starts_with('0') {
        return false;
    }
    let column = letters.bytes().try_fold(0u32, |value, byte| {
        value
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
    });
    let row = row.parse::<u32>().ok();
    column.is_some_and(|column| column <= 16_384)
        && row.is_some_and(|row| (1..=1_048_576).contains(&row))
}
