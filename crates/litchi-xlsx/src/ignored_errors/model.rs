/// One of the nine independent error conditions that a user may suppress.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(usize)]
pub enum IgnoredErrorType {
    CalculatedColumn,
    EmptyCellReference,
    EvaluationError,
    Formula,
    FormulaRange,
    ListDataValidation,
    NumberStoredAsText,
    TwoDigitTextYear,
    UnlockedFormula,
}

/// A validated A1 cell or cell-range reference from `sqref`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct IgnoredErrorRangeReference(pub(crate) String);

impl IgnoredErrorRangeReference {
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Inert, bounded markup retained from an `ignoredErrors/extLst/ext` entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IgnoredErrorsExtension {
    pub(crate) uri: String,
    pub(crate) markup: Vec<u8>,
}

impl IgnoredErrorsExtension {
    pub fn uri(&self) -> &str {
        &self.uri
    }
    /// MCE-processed extension markup. It is retained but never executed.
    pub fn markup(&self) -> &[u8] {
        &self.markup
    }
}

/// Error conditions suppressed for one or more worksheet ranges.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IgnoredError {
    pub(crate) ranges: Vec<IgnoredErrorRangeReference>,
    pub(crate) flags: [bool; 9],
}

impl IgnoredError {
    pub fn ranges(&self) -> &[IgnoredErrorRangeReference] {
        &self.ranges
    }
    pub fn ignores(&self, error_type: IgnoredErrorType) -> bool {
        self.flags[error_type as usize]
    }
}

/// Worksheet ignored-error collection in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IgnoredErrors {
    pub(crate) entries: Vec<IgnoredError>,
    pub(crate) extensions: Vec<IgnoredErrorsExtension>,
}

impl IgnoredErrors {
    pub fn entries(&self) -> &[IgnoredError] {
        &self.entries
    }
    pub fn extensions(&self) -> &[IgnoredErrorsExtension] {
        &self.extensions
    }
}
