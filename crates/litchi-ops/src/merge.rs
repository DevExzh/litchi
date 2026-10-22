//! Template merge (design §13 P0, record 0274): fill `{{key}}` placeholders
//! and Word `MERGEFIELD` fields from a data [`Record`].
//!
//! [`File::merge`](crate::File::merge) replaces every placeholder of a
//! template (DOCX, PPTX, XLSX, XLSB, DOC, PPT, XLS, RTF, ODT, OTH, ODM, ODS
//! or ODP) with its record value in **one** atomic native commit and returns the filled copy; the template handle is never
//! changed, so one opened template fills any number of records. `merge`
//! always means template (mail) merge here, as in OfficeCLI; document
//! merging is a separate concern.
//!
//! ```
//! # #[cfg(feature = "docx")]
//! # {
//! use litchi_core::FileFormat;
//! use litchi_ops::merge::{Data, Options};
//! use litchi_ops::transaction::SemanticLimits;
//! use litchi_ops::{NewNode, create};
//!
//! let template = create(FileFormat::Docx)?
//!     .add("/body", NewNode::new("paragraph").prop("text", "Dear {{name}}, {{ total }} due."))?
//!     .into_file();
//! let data = Data::from_json(
//!     br#"[{"name": "Ada", "total": 5200.00}, {"name": "Grace"}]"#,
//!     &SemanticLimits::default(),
//! )?;
//! let first = template.merge(&data.records()[0], &Options::new())?;
//! assert_eq!(first.file().text()?, "Dear Ada, 5200.00 due.");
//! let second = template.merge(&data.records()[1], &Options::new())?;
//! assert_eq!(second.file().text()?, "Dear Grace, {{ total }} due.");
//! assert_eq!(second.value().unresolved(), ["total"]);
//! # }
//! # Ok::<(), litchi_ops::Error>(())
//! ```
//!
//! # Placeholders
//!
//! A placeholder is OfficeCLI's `{{\s*(\w[\w.\-\[\] ]*?)\s*}}`: two opening
//! braces, optional white space, a key that starts with a letter, digit or
//! `_` and continues with those, `.`, `-`, `[`, `]` and spaces, optional
//! white space and two closing braces. The key keeps its inner spaces
//! (`{{ first name }}` is `first name`). Matching is one left-to-right pass,
//! and inserted values are never matched again.
//!
//! | Format | Where | Formatting |
//! |---|---|---|
//! | DOCX | every paragraph of the body (tables, content controls and text boxes included), headers, footers, footnotes, endnotes and comments | A placeholder may span runs. Its value takes the formatting of the placeholder's first character; the rest of the placeholder is removed from the runs it spans, which keep their own formatting. Tabs and line feeds in values become `w:tab` and `w:br`. |
//! | PPTX | every paragraph of every slide's shapes (group members included) | As DOCX, run by run; a value may not hold a line break. |
//! | XLSX | every text cell of every worksheet; formulas are never touched | The cell keeps its style; a rich-text cell is filled run by run and keeps each run's formatting (record 0295). |
//! | XLSB, XLS, ODS | as XLSX (record 0295) | The cell keeps its style. An XLSB rich-text cell (inline or a formatted shared string) and an ODS cell holding hyperlinks are fixed content; an XLS cell showing a formatted shared string becomes plain text. |
//! | DOC, RTF | every main-story paragraph, run by run (record 0295) | As DOCX; runs holding fields or drawings are fixed content, DOC tables are not merged; a value may not hold a line break. |
//! | PPT | every text shape and notes page, paragraph by paragraph (record 0295) | The native editor replaces a shape's text whole, keeping a single-run body's formatting; a value may not hold a line break. |
//! | ODT, OTH, ODM, ODP | every paragraph of the body (comments excluded) or of every slide's objects, table cells and notes, run by run (record 0295) | Each run keeps its style; runs holding fields or nested markup are fixed content; line feeds and tabs become `text:line-break` and `text:tab`. |
//!
//! A placeholder broken by other content (a tab, a drawing, a field) is not
//! a placeholder.
//!
//! # Word merge fields
//!
//! With [`Options::fields`] (the default) a DOCX `MERGEFIELD` field, simple
//! (`w:fldSimple`) or complex (`w:fldChar`), whose name is in the record is
//! replaced by one run holding its value, formatted as the field's result;
//! the name is matched exactly, then ignoring ASCII case. The `\b` (text
//! before) and `\f` (text after) switches apply to a non-empty value, and
//! `\* Upper`, `\* Lower`, `\* Caps` and `\* FirstCap` change its case;
//! numeric (`\#`) and date (`\@`) pictures are not applied. A field whose
//! markup holds other content (a bookmark, a nested field, a drawing) or
//! spans paragraphs is left as it is and counted in
//! [`Summary::skipped_fields`].
//!
//! # Missing keys
//!
//! A key the record lacks is listed in [`Summary::unresolved`]. With
//! [`Missing::Keep`] (the default, as OfficeCLI) its placeholder or field
//! stays; with [`Missing::Empty`] it is replaced by nothing, as Word's mail
//! merge does.
//!
//! # Data
//!
//! [`Data::from_json`] reads an object (one record) or an array of objects;
//! [`Data::from_csv`] reads a header row of keys and one record per row.
//! JSON values become text as OfficeCLI renders them: strings as
//! themselves, numbers as written, `true`/`false`, `null` as nothing, and
//! objects and arrays as indented JSON. Nested members are also reachable
//! as `a.b` and `items[0]` keys; a literal key of the same spelling wins.
//!
//! # Limits
//!
//! The transaction's [`SemanticLimits`]: the replacements of one merge are
//! bounded by `max_expanded_targets`, and their text plus a fixed charge
//! each by `max_staged_bytes`. Data sources are bounded by
//! `max_staged_bytes` for the source and for the rendered records.
//!
//! An exact no-op (no placeholder or field matches) makes no native commit:
//! the returned file is the template and its patch is a no-op.

#![cfg_attr(
    not(any(
        feature = "docx",
        feature = "pptx",
        feature = "xlsx",
        feature = "doc",
        feature = "ppt",
        feature = "xls",
        feature = "xlsb",
        feature = "rtf",
        feature = "odt",
        feature = "ods",
        feature = "odp"
    )),
    allow(dead_code, reason = "only the format adapters plan merges")
)]

use std::collections::BTreeSet;
use std::fmt;
use std::ops::Range;

use crate::error::Result;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

mod json;
#[cfg(test)]
mod tests;

/// The fixed staging charge of one replacement, beside its text.
const REPLACEMENT_CHARGE: usize = 64;

/// One data record: keys and their text values.
///
/// ```
/// use litchi_ops::merge::Record;
///
/// let record: Record = [("name", "Ada"), ("city", "London")].into_iter().collect();
/// assert_eq!(record.get("name"), Some("Ada"));
/// assert_eq!(record.len(), 2);
/// ```
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Record {
    entries: std::collections::BTreeMap<String, String>,
}

impl Record {
    /// An empty record.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set `key` to `value`, returning the previous value.
    pub fn insert(&mut self, key: impl Into<String>, value: impl Into<String>) -> Option<String> {
        self.entries.insert(key.into(), value.into())
    }

    /// The value of `key`.
    #[must_use]
    pub fn get(&self, key: &str) -> Option<&str> {
        self.entries.get(key).map(String::as_str)
    }

    /// The number of keys.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether the record has no key.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// The keys and values, ordered by key.
    pub fn iter(&self) -> impl Iterator<Item = (&str, &str)> {
        self.entries
            .iter()
            .map(|(key, value)| (key.as_str(), value.as_str()))
    }

    /// The value of a Word merge field name: exact, then the single key
    /// equal to it ignoring ASCII case.
    #[cfg(feature = "docx")]
    fn field<'s>(&'s self, name: &str) -> Option<(&'s str, &'s str)> {
        if let Some((key, value)) = self.entries.get_key_value(name) {
            return Some((key, value));
        }
        let mut matches = self
            .entries
            .iter()
            .filter(|(key, _)| key.eq_ignore_ascii_case(name));
        match (matches.next(), matches.next()) {
            (Some((key, value)), None) => Some((key, value)),
            _ => None,
        }
    }

    fn bytes(&self) -> usize {
        self.entries.iter().fold(0usize, |total, (key, value)| {
            total.saturating_add(key.len()).saturating_add(value.len())
        })
    }

    /// Flatten one JSON object as OfficeCLI does: literal members first,
    /// then nested `a.b` and `a[0]` paths for keys not yet set.
    fn from_object(members: &[(String, json::Value)]) -> Self {
        let mut record = Self::new();
        for (key, value) in members {
            record.entries.insert(key.clone(), json::render(value));
        }
        for (key, value) in members {
            match value {
                json::Value::Object(nested) => record.flatten_object(nested, key),
                json::Value::Array(items) => record.flatten_array(items, key),
                _ => {},
            }
        }
        record
    }

    fn fill(&mut self, path: String, value: &json::Value) {
        self.entries
            .entry(path)
            .or_insert_with(|| json::render(value));
    }

    fn flatten_object(&mut self, members: &[(String, json::Value)], prefix: &str) {
        for (key, value) in members {
            let path = format!("{prefix}.{key}");
            match value {
                json::Value::Object(nested) => {
                    self.flatten_object(nested, &path);
                    self.fill(path, value);
                },
                json::Value::Array(items) => self.flatten_array(items, &path),
                _ => self.fill(path, value),
            }
        }
    }

    fn flatten_array(&mut self, items: &[json::Value], prefix: &str) {
        for (index, item) in items.iter().enumerate() {
            let path = format!("{prefix}[{index}]");
            match item {
                json::Value::Object(nested) => {
                    self.flatten_object(nested, &path);
                    self.fill(path, item);
                },
                json::Value::Array(inner) => self.flatten_array(inner, &path),
                _ => self.fill(path, item),
            }
        }
    }
}

impl<K: Into<String>, V: Into<String>> FromIterator<(K, V)> for Record {
    fn from_iter<I: IntoIterator<Item = (K, V)>>(iter: I) -> Self {
        Self {
            entries: iter
                .into_iter()
                .map(|(key, value)| (key.into(), value.into()))
                .collect(),
        }
    }
}

impl fmt::Debug for Record {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Record")
            .field("keys", &self.entries.len())
            .finish_non_exhaustive()
    }
}

/// The records of a data source, in source order.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Data {
    records: Vec<Record>,
}

impl Data {
    /// Read a JSON object (one record) or an array of objects, under
    /// caller-selected limits ([`SemanticLimits::default`] for the defaults).
    ///
    /// ```
    /// use litchi_ops::merge::Data;
    /// use litchi_ops::transaction::SemanticLimits;
    ///
    /// let data = Data::from_json(
    ///     br#"{"client": {"name": "Acme"}, "items": [1, 2.50], "paid": true}"#,
    ///     &SemanticLimits::default(),
    /// )?;
    /// let record = &data.records()[0];
    /// assert_eq!(record.get("client.name"), Some("Acme"));
    /// assert_eq!(record.get("items[1]"), Some("2.50"));
    /// assert_eq!(record.get("paid"), Some("true"));
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// [`Code::InvalidJson`](crate::Code::InvalidJson) for malformed JSON
    /// or a value that is not an object or an array of objects;
    /// [`Code::LimitExceeded`](crate::Code::LimitExceeded) past the limits.
    pub fn from_json(source: &[u8], limits: &SemanticLimits) -> Result<Self> {
        check_bytes(source.len(), limits)?;
        let records = match json::parse(source)? {
            json::Value::Object(members) => vec![Record::from_object(&members)],
            json::Value::Array(items) => items
                .iter()
                .enumerate()
                .map(|(index, item)| match item {
                    json::Value::Object(members) => Ok(Record::from_object(members)),
                    _ => Err(Error::NotARecord { record: index + 1 }),
                })
                .collect::<std::result::Result<Vec<_>, _>>()?,
            _ => return Err(Error::NotAnObject.into()),
        };
        let data = Self { records };
        check_bytes(data.bytes(), limits)?;
        Ok(data)
    }

    /// Read CSV: the first record names the keys, and every later record is
    /// one data record. A short row's missing columns are empty; a row
    /// whose fields are all empty (a blank line) is skipped; an empty or
    /// repeated key column is ignored. Parsing follows
    /// [`crate::import`] (byte-order marks, `sep=`, quotes) without type
    /// inference. `options` selects the delimiter, quote and encoding
    /// (`import::Options::csv()` or `tsv()` for the common cases; inference
    /// is always off); `limits` bound the data.
    ///
    /// ```
    /// use litchi_ops::import::Options;
    /// use litchi_ops::merge::Data;
    /// use litchi_ops::transaction::SemanticLimits;
    ///
    /// let data = Data::from_csv(
    ///     b"name,total\nAda,\"5,200\"\nGrace\n",
    ///     &Options::csv(),
    ///     &SemanticLimits::default(),
    /// )?;
    /// assert_eq!(data.len(), 2);
    /// assert_eq!(data.records()[0].get("total"), Some("5,200"));
    /// assert_eq!(data.records()[1].get("total"), Some(""));
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// The [`crate::import`] parse errors, and
    /// [`Code::LimitExceeded`](crate::Code::LimitExceeded) past the limits.
    pub fn from_csv(
        source: &[u8],
        options: &crate::import::Options,
        limits: &SemanticLimits,
    ) -> Result<Self> {
        let options = (*options).infer(false).formulas(false);
        // The CSV front end bounds its fields by `max_expanded_targets`;
        // merge data is bounded by bytes only.
        let parse_limits = SemanticLimits::builder()
            .max_staged_bytes(limits.max_staged_bytes())?
            .max_expanded_targets(usize::MAX >> 1)?
            .build()?;
        let table = crate::import::parse(source, &options, &parse_limits)?;
        let mut keys: Vec<Option<String>> = Vec::new();
        let mut records: Vec<(Record, bool)> = Vec::new();
        for field in table.fields() {
            let text = match &field.value {
                crate::import::Value::Text(range) => table.text(range),
                _ => "",
            };
            let column = field.column as usize;
            if field.row == 0 {
                if keys.len() <= column {
                    keys.resize(column + 1, None);
                }
                let duplicate = keys.iter().flatten().any(|key| key == text);
                if !text.is_empty() && !duplicate {
                    keys[column] = Some(text.to_owned());
                }
                continue;
            }
            let row = field.row as usize - 1;
            while records.len() <= row {
                records.push((
                    keys.iter()
                        .flatten()
                        .map(|key| (key.clone(), String::new()))
                        .collect(),
                    false,
                ));
            }
            let (record, filled) = &mut records[row];
            *filled |= !text.is_empty();
            if let Some(Some(key)) = keys.get(column) {
                record.insert(key.clone(), text);
            }
        }
        let data = Self {
            records: records
                .into_iter()
                .filter_map(|(record, filled)| filled.then_some(record))
                .collect(),
        };
        check_bytes(data.bytes(), limits)?;
        Ok(data)
    }

    /// The records.
    #[must_use]
    pub fn records(&self) -> &[Record] {
        &self.records
    }

    /// Take the records.
    #[must_use]
    pub fn into_records(self) -> Vec<Record> {
        self.records
    }

    /// The number of records.
    #[must_use]
    pub fn len(&self) -> usize {
        self.records.len()
    }

    /// Whether there is no record.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.records.is_empty()
    }

    fn bytes(&self) -> usize {
        self.records
            .iter()
            .fold(0usize, |total, record| total.saturating_add(record.bytes()))
    }
}

impl fmt::Debug for Data {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Data")
            .field("records", &self.records.len())
            .finish()
    }
}

fn check_bytes(bytes: usize, limits: &SemanticLimits) -> Result<()> {
    if bytes > limits.max_staged_bytes() {
        return Err(TransactionError::Limit {
            kind: LimitKind::StagedBytes,
            observed: bytes,
            maximum: limits.max_staged_bytes(),
        }
        .into());
    }
    Ok(())
}

/// What a placeholder or merge field whose key the record lacks becomes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[non_exhaustive]
pub enum Missing {
    /// It stays as it is (OfficeCLI).
    #[default]
    Keep,
    /// It is replaced by nothing (Word's mail merge).
    Empty,
}

/// Merge options: missing keys kept, Word merge fields filled.
///
/// ```
/// use litchi_ops::merge::{Missing, Options};
///
/// let options = Options::new().missing(Missing::Empty).fields(false);
/// assert_eq!(options.missing_policy(), Missing::Empty);
/// assert!(!options.merges_fields());
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Options {
    missing: Missing,
    fields: bool,
}

impl Default for Options {
    fn default() -> Self {
        Self::new()
    }
}

impl Options {
    /// The defaults: [`Missing::Keep`] and Word merge fields filled.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            missing: Missing::Keep,
            fields: true,
        }
    }

    /// What a missing key's placeholder or field becomes.
    #[must_use]
    pub const fn missing(mut self, missing: Missing) -> Self {
        self.missing = missing;
        self
    }

    /// Whether DOCX `MERGEFIELD` fields are filled (on by default).
    #[must_use]
    pub const fn fields(mut self, fields: bool) -> Self {
        self.fields = fields;
        self
    }

    /// The missing-key policy.
    #[must_use]
    pub const fn missing_policy(self) -> Missing {
        self.missing
    }

    /// Whether merge fields are filled.
    #[must_use]
    pub const fn merges_fields(self) -> bool {
        self.fields
    }
}

/// What one merge did, the value of the returned
/// [`Commit`](crate::transaction::Commit).
///
/// `Debug` reports counts only; the key lists are authored content.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Summary {
    replacements: usize,
    fields: usize,
    skipped_fields: usize,
    keys: Vec<String>,
    unresolved: Vec<String>,
}

impl Summary {
    /// Placeholders replaced.
    #[must_use]
    pub const fn replacements(&self) -> usize {
        self.replacements
    }

    /// Word merge fields replaced.
    #[must_use]
    pub const fn fields(&self) -> usize {
        self.fields
    }

    /// Word merge fields left in place because their markup holds other
    /// content or spans paragraphs.
    #[must_use]
    pub const fn skipped_fields(&self) -> usize {
        self.skipped_fields
    }

    /// The distinct record keys used, sorted.
    #[must_use]
    pub fn keys(&self) -> &[String] {
        &self.keys
    }

    /// The distinct keys the template uses and the record lacks, sorted.
    #[must_use]
    pub fn unresolved(&self) -> &[String] {
        &self.unresolved
    }
}

impl fmt::Debug for Summary {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Summary")
            .field("replacements", &self.replacements)
            .field("fields", &self.fields)
            .field("skipped_fields", &self.skipped_fields)
            .field("keys", &self.keys.len())
            .field("unresolved", &self.unresolved.len())
            .finish()
    }
}

/// A content-free merge data failure.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The JSON is malformed at a one-based line and column.
    Json {
        /// The line.
        line: usize,
        /// The column, in characters.
        column: usize,
    },
    /// The JSON nests deeper than 128 levels.
    TooDeep,
    /// The JSON is neither an object nor an array.
    NotAnObject,
    /// An array item is not an object.
    NotARecord {
        /// The one-based item.
        record: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Json { line, column } => {
                write!(
                    formatter,
                    "the merge data is not valid JSON at line {line}, column {column}"
                )
            },
            Self::TooDeep => write!(
                formatter,
                "the merge data nests deeper than {} levels",
                json::MAX_DEPTH
            ),
            Self::NotAnObject => {
                formatter.write_str("the merge data must be a JSON object or an array of objects")
            },
            Self::NotARecord { record } => {
                write!(formatter, "merge data item {record} is not a JSON object")
            },
        }
    }
}

impl std::error::Error for Error {}

// ---------------------------------------------------------------------------
// Planning, shared by the format adapters
// ---------------------------------------------------------------------------

/// A `{{key}}` placeholder: its byte range and its key's.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Placeholder {
    pub(crate) range: Range<usize>,
    pub(crate) key: Range<usize>,
}

/// .NET's `\w`, as near as the standard library reaches: letters, digits
/// and `_`.
fn is_word(character: char) -> bool {
    character.is_alphanumeric() || character == '_'
}

fn is_key(character: char) -> bool {
    is_word(character) || matches!(character, '.' | '-' | '[' | ']' | ' ')
}

/// The placeholder starting at byte `start`, which holds `{{`.
fn placeholder_at(text: &str, start: usize) -> Option<Placeholder> {
    let rest = &text[start + 2..];
    let mut chars = rest.char_indices().peekable();
    while chars
        .next_if(|(_, character)| character.is_whitespace())
        .is_some()
    {}
    let (key_start, first) = chars.next()?;
    if !is_word(first) {
        return None;
    }
    let mut key_end = key_start + first.len_utf8();
    while let Some((offset, character)) = chars.next_if(|(_, character)| is_key(*character)) {
        key_end = offset + character.len_utf8();
    }
    let key = rest[key_start..key_end].trim_end_matches(' ');
    let key_end = key_start + key.len();
    let mut after = chars.peek().map_or(rest.len(), |(offset, _)| *offset);
    after += rest[after..].len() - rest[after..].trim_start().len();
    rest[after..].starts_with("}}").then(|| Placeholder {
        range: start..start + 2 + after + 2,
        key: start + 2 + key_start..start + 2 + key_end,
    })
}

/// The placeholders of `text`, left to right.
pub(crate) fn placeholders(text: &str) -> impl Iterator<Item = Placeholder> + '_ {
    let mut from = 0usize;
    let bytes = text.as_bytes();
    std::iter::from_fn(move || {
        while let Some(found) = bytes.get(from..)?.windows(2).position(|pair| pair == b"{{") {
            let start = from + found;
            if let Some(placeholder) = placeholder_at(text, start) {
                from = placeholder.range.end;
                return Some(placeholder);
            }
            from = start + 1;
        }
        None
    })
}

/// One merge's bookkeeping: values, counts and limits.
pub(crate) struct Planner<'a> {
    record: &'a Record,
    options: Options,
    limits: SemanticLimits,
    keys: BTreeSet<&'a str>,
    unresolved: BTreeSet<String>,
    replacements: usize,
    fields: usize,
    skipped_fields: usize,
    staged: usize,
}

impl<'a> Planner<'a> {
    pub(crate) fn new(record: &'a Record, options: &Options, limits: &SemanticLimits) -> Self {
        Self {
            record,
            options: *options,
            limits: *limits,
            keys: BTreeSet::new(),
            unresolved: BTreeSet::new(),
            replacements: 0,
            fields: 0,
            skipped_fields: 0,
            staged: 0,
        }
    }

    #[cfg(feature = "docx")]
    pub(crate) const fn options(&self) -> Options {
        self.options
    }

    fn charge(&mut self, text: &str) -> Result<()> {
        let count = self.replacements + self.fields + 1;
        if count > self.limits.max_expanded_targets() {
            return Err(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed: count,
                maximum: self.limits.max_expanded_targets(),
            }
            .into());
        }
        self.staged = self
            .staged
            .saturating_add(text.len())
            .saturating_add(REPLACEMENT_CHARGE);
        if self.staged > self.limits.max_staged_bytes() {
            return Err(TransactionError::Limit {
                kind: LimitKind::StagedBytes,
                observed: self.staged,
                maximum: self.limits.max_staged_bytes(),
            }
            .into());
        }
        Ok(())
    }

    /// The value of a placeholder key, or `None` to keep it.
    pub(crate) fn value(&mut self, key: &str) -> Result<Option<&'a str>> {
        let value = match self.record.entries.get_key_value(key) {
            Some((key, value)) => {
                // Each value is checked once, when its key is first used.
                if self.keys.insert(key.as_str()) {
                    validate_value(value)?;
                }
                value.as_str()
            },
            None => {
                if !self.unresolved.contains(key) {
                    self.unresolved.insert(key.to_owned());
                }
                match self.options.missing {
                    Missing::Keep => return Ok(None),
                    Missing::Empty => "",
                }
            },
        };
        self.charge(value)?;
        self.replacements += 1;
        Ok(Some(value))
    }

    /// Plan the placeholders of `text`: `is_text` says whether a range may
    /// be replaced, `outside` lists ranges already taken (sorted). Each
    /// accepted placeholder is passed to `accept` with its value, which
    /// [`validate_value`] has accepted.
    pub(crate) fn text(
        &mut self,
        text: &str,
        is_text: impl Fn(Range<usize>) -> bool,
        outside: &[Range<usize>],
        mut accept: impl FnMut(Range<usize>, &'a str),
    ) -> Result<()> {
        if !text.contains("{{") {
            return Ok(());
        }
        for placeholder in placeholders(text) {
            let range = placeholder.range.clone();
            if outside
                .iter()
                .any(|taken| taken.start < range.end && range.start < taken.end)
                || !is_text(range.clone())
            {
                continue;
            }
            if let Some(value) = self.value(&text[placeholder.key])? {
                accept(range, value);
            }
        }
        Ok(())
    }

    /// The text replacing a Word merge field with `instruction`, or `None`
    /// to keep it (not a `MERGEFIELD`, fields off, or a kept missing key).
    #[cfg(feature = "docx")]
    pub(crate) fn field(&mut self, instruction: &str, replaceable: bool) -> Result<Option<String>> {
        if !self.options.fields {
            return Ok(None);
        }
        let field = litchi_docx::field::Field::new(instruction.to_owned(), None, false);
        let Ok(Some(merge)) = field.merge_field() else {
            return Ok(None);
        };
        if !replaceable {
            self.skipped_fields += 1;
            return Ok(None);
        }
        let name = merge.field_name();
        let value = match self.record.field(name) {
            Some((key, value)) => {
                self.keys.insert(key);
                value
            },
            None => {
                if !self.unresolved.contains(name) {
                    self.unresolved.insert(name.to_owned());
                }
                match self.options.missing {
                    Missing::Keep => return Ok(None),
                    Missing::Empty => "",
                }
            },
        };
        let mut text = apply_case(value, merge.switches());
        if !text.is_empty() {
            if let Some(before) = switch(merge.switches(), 'b') {
                text.insert_str(0, before);
            }
            if let Some(after) = switch(merge.switches(), 'f') {
                text.push_str(after);
            }
        }
        self.charge(&text)?;
        self.fields += 1;
        Ok(Some(text))
    }

    pub(crate) fn finish(self) -> Summary {
        Summary {
            replacements: self.replacements,
            fields: self.fields,
            skipped_fields: self.skipped_fields,
            keys: self.keys.into_iter().map(str::to_owned).collect(),
            unresolved: self.unresolved.into_iter().collect(),
        }
    }
}

#[cfg(feature = "docx")]
fn switch(switches: &[litchi_docx::field::Switch], name: char) -> Option<&str> {
    switches
        .iter()
        .find(|switch| switch.name().eq_ignore_ascii_case(&name))
        .and_then(litchi_docx::field::Switch::argument)
}

/// Word's `\*` case formats.
#[cfg(feature = "docx")]
fn apply_case(value: &str, switches: &[litchi_docx::field::Switch]) -> String {
    let format = switches
        .iter()
        .filter(|switch| switch.name() == '*')
        .filter_map(litchi_docx::field::Switch::argument)
        .find(|argument| {
            ["upper", "lower", "caps", "firstcap"]
                .iter()
                .any(|name| argument.eq_ignore_ascii_case(name))
        });
    match format.map(str::to_ascii_lowercase).as_deref() {
        Some("upper") => value.to_uppercase(),
        Some("lower") => value.to_lowercase(),
        Some("caps") => {
            let mut output = String::with_capacity(value.len());
            let mut start = true;
            for character in value.chars() {
                if start && character.is_alphabetic() {
                    output.extend(character.to_uppercase());
                } else {
                    output.extend(character.to_lowercase());
                }
                start = character.is_whitespace();
            }
            output
        },
        Some("firstcap") => {
            let mut chars = value.chars();
            chars.next().map_or_else(String::new, |first| {
                first.to_uppercase().chain(chars).collect()
            })
        },
        _ => value.to_owned(),
    }
}

/// Text values keep tabs and line feeds; a carriage return (alone or before
/// a line feed) becomes a line feed.
#[cfg(any(
    feature = "docx",
    feature = "xlsx",
    feature = "rtf",
    feature = "doc",
    feature = "ppt",
    feature = "xls",
    feature = "xlsb",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    test
))]
pub(crate) fn normalize_breaks(value: &str) -> std::borrow::Cow<'_, str> {
    if value.contains('\r') {
        std::borrow::Cow::Owned(value.replace("\r\n", "\n").replace('\r', "\n"))
    } else {
        std::borrow::Cow::Borrowed(value)
    }
}

/// Check that a value holds only characters XML text can carry.
pub(crate) fn validate_value(value: &str) -> Result<()> {
    if value.chars().any(|character| {
        !matches!(character, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
            || character == '\u{FFFC}'
    }) {
        return Err(crate::Error::new(
            crate::Code::InvalidValue,
            "a merge value holds a character a document cannot store",
        )
        .with_suggestion("remove control characters from the data"));
    }
    Ok(())
}
