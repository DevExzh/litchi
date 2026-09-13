//! OpenFormula 1.4 reference syntax.
//!
//! Bracketed references are deliberately kept separate from the older
//! [`super::CellRef`] and [`super::RangeRef`] projection.  The grammar in
//! OpenFormula 1.4 can carry a source IRI, a subtable locator, whole rows or
//! columns, and an omitted second locator which inherits the first one.  None
//! of those states can be represented by the legacy structures without losing
//! information, so this module owns the complete, inert representation.

use litchi_core::{Error, Resource, ResourceLimit, Result};
use std::{convert::TryFrom, sync::Arc};

mod iri;

use self::iri::is_valid_iri_reference;

/// Default maximum number of bytes in one bracketed reference body.
pub const DEFAULT_MAX_REFERENCE_BYTES: usize = 64 * 1024;
/// Default maximum number of locator and endpoint components.
pub const DEFAULT_MAX_REFERENCE_COMPONENTS: usize = 256;
/// Default maximum size of one source, sheet, or column lexical value.
pub const DEFAULT_MAX_REFERENCE_NAME_BYTES: usize = 16 * 1024;

/// Finite limits applied while parsing one OpenFormula reference.
///
/// The limits cover the complete bracket body and every owned component.  A
/// parser never grows a component vector past `max_components`, and all
/// vectors reserve before they are mutated.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_bytes: usize,
    max_components: usize,
    max_name_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_bytes: DEFAULT_MAX_REFERENCE_BYTES,
            max_components: DEFAULT_MAX_REFERENCE_COMPONENTS,
            max_name_bytes: DEFAULT_MAX_REFERENCE_NAME_BYTES,
        }
    }
}

impl Limits {
    /// Set the maximum UTF-8 byte length of the bracketed reference body.
    #[must_use]
    pub const fn with_max_bytes(mut self, value: usize) -> Self {
        self.max_bytes = value;
        self
    }

    /// Set the maximum number of address, locator, and subtable components.
    #[must_use]
    pub const fn with_max_components(mut self, value: usize) -> Self {
        self.max_components = value;
        self
    }

    /// Set the maximum UTF-8 byte length of one source, sheet, or column
    /// component.
    #[must_use]
    pub const fn with_max_name_bytes(mut self, value: usize) -> Self {
        self.max_name_bytes = value;
        self
    }

    /// Return the maximum reference body length.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    /// Return the maximum component count.
    #[must_use]
    pub const fn max_components(self) -> usize {
        self.max_components
    }

    /// Return the maximum component string length.
    #[must_use]
    pub const fn max_name_bytes(self) -> usize {
        self.max_name_bytes
    }
}

/// A source IRI in the OpenFormula `'IRI'#` prefix.
///
/// The IRI is inert.  Parsing and inspecting it never opens, resolves, or
/// otherwise contacts the referenced source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Source {
    /// The decoded IRI value.  Doubled apostrophes in formula syntax have
    /// already been reduced to one apostrophe.
    pub iri: String,
}

impl Source {
    /// Validate and create an inert source IRI.
    pub fn new(iri: &str) -> Result<Self> {
        if iri.len() > DEFAULT_MAX_REFERENCE_NAME_BYTES {
            return Err(limit_error(
                "source IRI bytes",
                iri.len(),
                DEFAULT_MAX_REFERENCE_NAME_BYTES,
            ));
        }
        if !is_valid_iri_reference(iri) {
            return Err(Error::InvalidFormat(
                "invalid OpenFormula source IRI".to_string(),
            ));
        }
        Ok(Self {
            iri: copy_component(iri, "formula reference source IRI")?,
        })
    }

    /// Return the decoded IRI.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.iri
    }
}

/// A sheet name in a sheet locator.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetName {
    /// Decoded sheet name, without the optional `$` marker or quote pair.
    pub name: String,
    /// Whether the sheet name carried the `$` absolute marker.
    pub absolute: bool,
    /// Whether the sheet name used the quoted lexical form.
    pub quoted: bool,
}

impl SheetName {
    /// Return the decoded sheet name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.name
    }
}

/// A subtable component in a sheet locator.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Subtable {
    /// A cell address selecting a subtable.
    Cell(Cell),
    /// A quoted sheet-name component selecting a subtable.
    Name(SheetName),
}

/// A cell coordinate used by an endpoint or subtable selector.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Cell {
    /// Column coordinate.
    pub column: Column,
    /// Row coordinate.
    pub row: Row,
}

/// A column coordinate, including its absolute marker.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Column {
    /// Uppercase A-Z column label.
    pub label: String,
    /// Whether the column carried the `$` absolute marker.
    pub absolute: bool,
}

/// A one-based row coordinate, including its absolute marker.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Row {
    /// One-based row number.
    pub number: u32,
    /// Whether the row carried the `$` absolute marker.
    pub absolute: bool,
}

/// A locator for a sheet and its optional subtable chain.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetLocator {
    /// The first sheet-name component.
    pub sheet: SheetName,
    /// Ordered subtable components following the sheet name.
    pub subtables: Vec<Subtable>,
}

impl SheetLocator {
    /// Return the first sheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &SheetName {
        &self.sheet
    }

    /// Return the ordered subtable chain.
    #[must_use]
    pub fn subtables(&self) -> &[Subtable] {
        &self.subtables
    }
}

/// The sheet-selection state carried by one reference endpoint.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SheetSelector {
    /// No sheet locator was written; the evaluator's current sheet applies.
    Current,
    /// A sheet locator was written on this endpoint.
    Explicit(SheetLocator),
    /// The endpoint omitted its locator and inherits the first endpoint's
    /// locator according to OpenFormula 1.4 section 5.8.
    Inherited,
}

/// The coordinate kind carried by an endpoint.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum EndpointValue {
    /// A single cell coordinate.
    Cell(Cell),
    /// A whole-column coordinate.
    Column(Column),
    /// A whole-row coordinate.
    Row(Row),
}

/// One endpoint of a cell, row, or column address.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Endpoint {
    /// Sheet selection for this endpoint.
    pub sheet: SheetSelector,
    /// Cell, column, or row coordinate.
    pub value: EndpointValue,
}

impl Endpoint {
    /// Construct a cell endpoint.
    #[must_use]
    pub fn cell(sheet: SheetSelector, cell: Cell) -> Self {
        Self {
            sheet,
            value: EndpointValue::Cell(cell),
        }
    }

    /// Construct a whole-column endpoint.
    #[must_use]
    pub fn column(sheet: SheetSelector, column: Column) -> Self {
        Self {
            sheet,
            value: EndpointValue::Column(column),
        }
    }

    /// Construct a whole-row endpoint.
    #[must_use]
    pub fn row(sheet: SheetSelector, row: Row) -> Self {
        Self {
            sheet,
            value: EndpointValue::Row(row),
        }
    }
}

/// A parsed range address.
///
/// `Cells`, `Columns`, and `Rows` use endpoint sheet selectors directly, so
/// two different explicit locators represent a cross-sheet cuboid without a
/// second cross-sheet enum variant.  The invalidated `#REF!` state belongs to
/// [`Reference::Error`], so it cannot be paired with a source IRI.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Address {
    /// One cell endpoint.
    Cell(Endpoint),
    /// A cell range.
    Cells(Endpoint, Endpoint),
    /// A whole-column range.
    Columns(Endpoint, Endpoint),
    /// A whole-row range.
    Rows(Endpoint, Endpoint),
}

/// A complete OpenFormula 1.4 reference.
///
/// The invalidated `#REF!` state is a direct variant and therefore cannot be
/// combined with the source-bearing variant.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Reference {
    /// A local, non-error range address.
    Local(Address),
    /// A source-qualified, non-error reference.
    Source {
        /// Inert source IRI.
        source: Source,
        /// Valid range address in that source.
        address: Address,
    },
    /// An invalidated reference (`#REF!`).
    Error,
}

impl Reference {
    /// Parse one complete bracketed reference.
    pub fn parse(value: &str) -> Result<Self> {
        parse(value)
    }

    /// Parse one reference with explicit finite limits.
    pub fn parse_with_limits(value: &str, limits: &Limits) -> Result<Self> {
        parse_with_limits(value, limits)
    }

    /// Return the source IRI, if this is source-qualified.
    #[must_use]
    pub fn source(&self) -> Option<&Source> {
        match self {
            Self::Local(_) | Self::Error => None,
            Self::Source { source, .. } => Some(source),
        }
    }

    /// Return whether this reference is an explicit invalidated reference.
    #[must_use]
    pub fn is_error(&self) -> bool {
        matches!(self, Self::Error)
    }

    /// Borrow the parsed address, if this is not `#REF!`.
    #[must_use]
    pub fn address(&self) -> Option<&Address> {
        match self {
            Self::Local(address) | Self::Source { address, .. } => Some(address),
            Self::Error => None,
        }
    }

    /// Build the legacy cell projection used by [`super::extract_cell_refs`].
    ///
    /// The projection is intentionally limited to a local cell endpoint with
    /// no subtable selector.  Quoting and an absolute sheet marker are
    /// representable as a legacy `CellRef` sheet name, although those lexical
    /// markers are necessarily lost by that older shape.  Source-qualified,
    /// whole-axis, subtable, and invalidated references return `None` instead
    /// of being reported as a different local cell.
    pub(crate) fn legacy_cell_ref(&self) -> Option<super::CellRef> {
        let address = match self {
            Self::Local(address) => address,
            Self::Source { .. } | Self::Error => return None,
        };
        match address {
            Address::Cell(endpoint) | Address::Cells(endpoint, _) => {
                legacy_projection_cell(endpoint)
            },
            Address::Columns(_, _) | Address::Rows(_, _) => None,
        }
    }

    /// Convert to a token, retaining a legacy token only when it can carry every meaningful state of this
    /// reference.  Bracketed references with source IRIs, absolute sheet
    /// locators, quoted names, subtables, whole axes, or inherited endpoints
    /// deliberately stay in the rich representation.
    pub(crate) fn into_token(self) -> super::Token {
        match self {
            Self::Local(address) => address.into_token(),
            Self::Source { source, address } => {
                super::Token::Reference(Box::new(Self::Source { source, address }))
            },
            Self::Error => super::Token::Reference(Box::new(Self::Error)),
        }
    }
}

impl Address {
    fn into_token(self) -> super::Token {
        let compatible = self.is_legacy_compatible();
        match self {
            Self::Cell(endpoint) if compatible => {
                super::Token::CellRef(legacy_cell(endpoint).expect("validated legacy cell"))
            },
            Self::Cells(start, end) if compatible => super::Token::RangeRef(super::RangeRef {
                start: legacy_cell(start).expect("validated legacy range start"),
                end: legacy_cell(end).expect("validated legacy range end"),
            }),
            address => super::Token::Reference(Box::new(Reference::Local(address))),
        }
    }

    fn is_legacy_compatible(&self) -> bool {
        match self {
            Self::Cell(endpoint) => legacy_cell_shape(endpoint),
            Self::Cells(start, end) => legacy_cell_shape(start) && legacy_cell_shape(end),
            Self::Columns(_, _) | Self::Rows(_, _) => false,
        }
    }
}

fn legacy_cell_shape(endpoint: &Endpoint) -> bool {
    match &endpoint.sheet {
        SheetSelector::Current => {},
        SheetSelector::Explicit(locator)
            if locator.subtables.is_empty() && !locator.sheet.absolute && !locator.sheet.quoted => {
        },
        SheetSelector::Explicit(_) | SheetSelector::Inherited => return false,
    }
    matches!(endpoint.value, EndpointValue::Cell(_))
}

fn legacy_cell(endpoint: Endpoint) -> Option<super::CellRef> {
    let sheet = match endpoint.sheet {
        SheetSelector::Current => None,
        SheetSelector::Explicit(locator)
            if locator.subtables.is_empty() && !locator.sheet.absolute && !locator.sheet.quoted =>
        {
            Some(locator.sheet.name)
        },
        SheetSelector::Explicit(_) | SheetSelector::Inherited => return None,
    };
    let EndpointValue::Cell(cell) = endpoint.value else {
        return None;
    };
    Some(super::CellRef {
        sheet,
        column: cell.column.label,
        row: cell.row.number,
        column_absolute: cell.column.absolute,
        row_absolute: cell.row.absolute,
    })
}

fn legacy_projection_cell(endpoint: &Endpoint) -> Option<super::CellRef> {
    let sheet = match &endpoint.sheet {
        SheetSelector::Current => None,
        SheetSelector::Explicit(locator) if locator.subtables.is_empty() => {
            Some(locator.sheet.name.clone())
        },
        SheetSelector::Explicit(_) | SheetSelector::Inherited => return None,
    };
    let EndpointValue::Cell(cell) = &endpoint.value else {
        return None;
    };
    Some(super::CellRef {
        sheet,
        column: cell.column.label.clone(),
        row: cell.row.number,
        column_absolute: cell.column.absolute,
        row_absolute: cell.row.absolute,
    })
}

/// Parse one OpenFormula reference with default limits.
pub fn parse(value: &str) -> Result<Reference> {
    parse_with_limits(value, &Limits::default())
}

/// Parse one OpenFormula reference with explicit finite limits.
pub fn parse_with_limits(value: &str, limits: &Limits) -> Result<Reference> {
    let value = value.trim();
    if !value.starts_with('[') || !value.ends_with(']') || value.len() < 2 {
        return Err(invalid("a reference must be enclosed in brackets"));
    }
    parse_body(&value[1..value.len() - 1], limits)
}

pub(crate) fn parse_body(value: &str, limits: &Limits) -> Result<Reference> {
    if value.len() > limits.max_bytes {
        return Err(limit_error(
            "reference bytes",
            value.len(),
            limits.max_bytes,
        ));
    }
    // The public wrapper may be surrounded by formula whitespace, but bytes
    // inside the brackets belong to the reference grammar.  Keep them intact
    // so whitespace cannot be silently accepted inside a sheet or coordinate.
    let mut parser = Parser::new(value, *limits);
    parser.parse()
}

struct Parser<'a> {
    input: &'a str,
    bytes: &'a [u8],
    position: usize,
    limits: Limits,
    components: usize,
}

impl<'a> Parser<'a> {
    fn new(input: &'a str, limits: Limits) -> Self {
        Self {
            input,
            bytes: input.as_bytes(),
            position: 0,
            limits,
            components: 0,
        }
    }

    fn parse(&mut self) -> Result<Reference> {
        if self.input.is_empty() {
            return Err(invalid("empty OpenFormula reference"));
        }

        if self.input == "#REF!" {
            return Ok(Reference::Error);
        }

        let source = self.parse_source_prefix()?;
        let first = self.parse_first_endpoint()?;
        let address = if self.at_end() {
            match first.value {
                EndpointValue::Cell(_) => Address::Cell(first),
                EndpointValue::Column(_) | EndpointValue::Row(_) => {
                    return Err(invalid("whole-row or whole-column reference needs a range"));
                },
            }
        } else {
            self.expect_byte(b':', "reference range separator")?;
            let second = if self.consume_byte(b'.') {
                Endpoint {
                    sheet: SheetSelector::Inherited,
                    value: self.parse_coordinate()?,
                }
            } else {
                if !matches!(first.sheet, SheetSelector::Explicit(_)) {
                    return Err(invalid(
                        "a range endpoint may name its sheet only after an explicit first sheet",
                    ));
                }
                self.parse_explicit_endpoint()?
            };

            if !self.at_end() {
                return Err(invalid("trailing bytes in OpenFormula reference"));
            }
            self.combine_range(first, second)?
        };

        if !self.at_end() {
            return Err(invalid("trailing bytes in OpenFormula reference"));
        }

        if let Some(source) = source {
            Ok(Reference::Source { source, address })
        } else {
            Ok(Reference::Local(address))
        }
    }

    fn parse_source_prefix(&mut self) -> Result<Option<Source>> {
        if self.peek() != Some(b'\'') {
            return Ok(None);
        }

        let checkpoint = self.position;
        let close = self.scan_quoted(checkpoint)?;
        if self.bytes.get(close + 1) != Some(&b'#') {
            return Ok(None);
        }

        self.position = checkpoint;
        self.bump_component()?;
        let (iri, close) = self.parse_quoted_text(true)?;
        debug_assert_eq!(self.bytes.get(close), Some(&b'\''));
        self.position = close + 1;
        self.expect_byte(b'#', "source IRI terminator")?;
        if iri.len() > self.limits.max_name_bytes {
            return Err(limit_error(
                "source IRI bytes",
                iri.len(),
                self.limits.max_name_bytes,
            ));
        }
        if !is_valid_iri_reference(&iri) {
            return Err(invalid("invalid OpenFormula source IRI"));
        }
        Ok(Some(Source { iri }))
    }

    fn parse_first_endpoint(&mut self) -> Result<Endpoint> {
        if self.consume_byte(b'.') {
            return Ok(Endpoint {
                sheet: SheetSelector::Current,
                value: self.parse_coordinate()?,
            });
        }
        self.parse_explicit_endpoint()
    }

    fn parse_explicit_endpoint(&mut self) -> Result<Endpoint> {
        let sheet = self.parse_sheet_name()?;
        let mut subtables = Vec::new();

        loop {
            self.expect_byte(b'.', "sheet and coordinate separator")?;
            if self.peek() == Some(b'\'') || self.peek_dollar_quote() {
                let name = self.parse_sheet_name()?;
                self.push_subtable_into(&mut subtables, Subtable::Name(name))?;
                continue;
            }

            let coordinate = self.parse_coordinate()?;
            if matches!(coordinate, EndpointValue::Cell(_)) && self.peek() == Some(b'.') {
                let EndpointValue::Cell(cell) = coordinate else {
                    unreachable!("coordinate was checked as a cell")
                };
                self.push_subtable_into(&mut subtables, Subtable::Cell(cell))?;
                continue;
            }
            if self.peek() == Some(b'.') {
                return Err(invalid(
                    "whole-row or whole-column cannot select a subtable",
                ));
            }
            let locator = SheetLocator { sheet, subtables };
            return Ok(Endpoint {
                sheet: SheetSelector::Explicit(locator),
                value: coordinate,
            });
        }
    }

    fn combine_range(&self, first: Endpoint, second: Endpoint) -> Result<Address> {
        match (&first.value, &second.value) {
            (EndpointValue::Cell(_), EndpointValue::Cell(_)) => Ok(Address::Cells(first, second)),
            (EndpointValue::Column(_), EndpointValue::Column(_)) => {
                Ok(Address::Columns(first, second))
            },
            (EndpointValue::Row(_), EndpointValue::Row(_)) => Ok(Address::Rows(first, second)),
            _ => Err(invalid(
                "range endpoints must both be cells, columns, or rows",
            )),
        }
    }

    fn parse_coordinate(&mut self) -> Result<EndpointValue> {
        let start = self.position;
        let column_absolute = self.consume_byte(b'$');
        let column_start = self.position;
        while self.peek().is_some_and(|byte| byte.is_ascii_uppercase()) {
            self.advance();
        }
        if self.position != column_start {
            let column_label = &self.input[column_start..self.position];
            if column_label.len() > self.limits.max_name_bytes {
                return Err(limit_error(
                    "column bytes",
                    column_label.len(),
                    self.limits.max_name_bytes,
                ));
            }

            let row_absolute = self.consume_byte(b'$');
            let row_start = self.position;
            if self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                if self.peek() == Some(b'0') {
                    return Err(invalid("OpenFormula rows are one-based"));
                }
                while self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                    self.advance();
                }
                let row_digits = &self.input[row_start..self.position];
                if row_digits.len() > self.limits.max_name_bytes {
                    return Err(limit_error(
                        "row bytes",
                        row_digits.len(),
                        self.limits.max_name_bytes,
                    ));
                }
                let row_number = row_digits
                    .parse::<u32>()
                    .map_err(|_error| invalid("row exceeds the representable range"))?;
                self.bump_component()?;
                self.bump_component()?;
                let column = Column {
                    label: copy_component(column_label, "formula reference column")?,
                    absolute: column_absolute,
                };
                return Ok(EndpointValue::Cell(Cell {
                    column,
                    row: Row {
                        number: row_number,
                        absolute: row_absolute,
                    },
                }));
            }
            if row_absolute {
                return Err(invalid("absolute column must be followed by a row"));
            }
            if self.peek() == Some(b':') || self.at_end() || self.peek() == Some(b'.') {
                self.bump_component()?;
                return Ok(EndpointValue::Column(Column {
                    label: copy_component(column_label, "formula reference column")?,
                    absolute: column_absolute,
                }));
            }
            return Err(invalid("invalid OpenFormula column endpoint"));
        }

        self.position = start;
        let row_absolute = self.consume_byte(b'$');
        let row_start = self.position;
        if !self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
            return Err(invalid("expected uppercase column or one-based row"));
        }
        if self.peek() == Some(b'0') {
            return Err(invalid("OpenFormula rows are one-based"));
        }
        while self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
            self.advance();
        }
        let row_digits = &self.input[row_start..self.position];
        if row_digits.len() > self.limits.max_name_bytes {
            return Err(limit_error(
                "row bytes",
                row_digits.len(),
                self.limits.max_name_bytes,
            ));
        }
        if !self.at_end() && self.peek() != Some(b':') && self.peek() != Some(b'.') {
            return Err(invalid("invalid OpenFormula row endpoint"));
        }
        let number = row_digits
            .parse::<u32>()
            .map_err(|_error| invalid("row exceeds the representable range"))?;
        self.bump_component()?;
        Ok(EndpointValue::Row(Row {
            number,
            absolute: row_absolute,
        }))
    }

    fn parse_sheet_name(&mut self) -> Result<SheetName> {
        let absolute = self.consume_byte(b'$');
        if self.peek() == Some(b'\'') {
            self.bump_component()?;
            let (name, close) = self.parse_quoted_text(false)?;
            self.position = close + 1;
            if name.is_empty() {
                return Err(invalid("sheet names cannot be empty"));
            }
            return Ok(SheetName {
                name,
                absolute,
                quoted: true,
            });
        }

        let start = self.position;
        while let Some(byte) = self.peek() {
            if byte == b'.'
                || byte == b']'
                || byte == b'#'
                || byte == b'$'
                || byte == b'\''
                || byte.is_ascii_whitespace()
            {
                break;
            }
            self.advance();
        }
        if start == self.position {
            return Err(invalid("expected non-empty sheet name"));
        }
        let name = &self.input[start..self.position];
        if name.len() > self.limits.max_name_bytes {
            return Err(limit_error(
                "sheet name bytes",
                name.len(),
                self.limits.max_name_bytes,
            ));
        }
        self.bump_component()?;
        Ok(SheetName {
            name: copy_component(name, "formula reference sheet name")?,
            absolute,
            quoted: false,
        })
    }

    fn parse_quoted_text(&mut self, source: bool) -> Result<(String, usize)> {
        let opening = self.position;
        let close = self.scan_quoted(opening)?;
        let inner_start = opening + 1;
        let inner = &self.input[inner_start..close];
        let mut decoded_len = 0usize;
        let mut position = 0usize;
        while position < inner.len() {
            let byte = inner.as_bytes()[position];
            if byte == b'\'' {
                if inner.as_bytes().get(position + 1) != Some(&b'\'') {
                    return Err(invalid("invalid apostrophe escape"));
                }
                position += 2;
                decoded_len = decoded_len
                    .checked_add(1)
                    .ok_or_else(|| invalid("quoted component length overflow"))?;
            } else {
                let character = inner[position..]
                    .chars()
                    .next()
                    .ok_or_else(|| invalid("invalid UTF-8 in quoted component"))?;
                let width = character.len_utf8();
                position += width;
                decoded_len = decoded_len
                    .checked_add(width)
                    .ok_or_else(|| invalid("quoted component length overflow"))?;
            }
        }
        if decoded_len > self.limits.max_name_bytes {
            return Err(limit_error(
                if source {
                    "source IRI bytes"
                } else {
                    "sheet name bytes"
                },
                decoded_len,
                self.limits.max_name_bytes,
            ));
        }

        let mut decoded = String::new();
        decoded
            .try_reserve_exact(decoded_len)
            .map_err(|source| Error::Allocation {
                resource: "formula reference component",
                source,
            })?;
        position = 0;
        while position < inner.len() {
            if inner.as_bytes()[position] == b'\'' {
                decoded.push('\'');
                position += 2;
            } else {
                let character = inner[position..]
                    .chars()
                    .next()
                    .ok_or_else(|| invalid("invalid UTF-8 in quoted component"))?;
                position += character.len_utf8();
                decoded.push(character);
            }
        }
        Ok((decoded, close))
    }

    fn scan_quoted(&self, opening: usize) -> Result<usize> {
        if self.bytes.get(opening) != Some(&b'\'') {
            return Err(invalid("expected quoted component"));
        }
        let mut position = opening + 1;
        while position < self.bytes.len() {
            if self.bytes[position] != b'\'' {
                let character = self.input[position..]
                    .chars()
                    .next()
                    .ok_or_else(|| invalid("invalid UTF-8 in quoted component"))?;
                position += character.len_utf8();
                continue;
            }
            if self.bytes.get(position + 1) == Some(&b'\'') {
                position += 2;
                continue;
            }
            return Ok(position);
        }
        Err(invalid("unterminated quoted component"))
    }

    fn push_subtable_into(
        &mut self,
        subtables: &mut Vec<Subtable>,
        subtable: Subtable,
    ) -> Result<()> {
        self.bump_component()?;
        if subtables.len() == subtables.capacity() {
            let next_capacity = if subtables.capacity() == 0 {
                1
            } else {
                subtables
                    .capacity()
                    .checked_mul(2)
                    .ok_or_else(|| invalid("subtable capacity overflow"))?
                    .min(self.limits.max_components)
            };
            let additional = next_capacity
                .checked_sub(subtables.len())
                .ok_or_else(|| invalid("subtable capacity overflow"))?;
            subtables
                .try_reserve_exact(additional)
                .map_err(|source| Error::Allocation {
                    resource: "formula reference subtables",
                    source,
                })?;
        }
        subtables.push(subtable);
        Ok(())
    }

    #[inline]
    fn bump_component(&mut self) -> Result<()> {
        self.components = self
            .components
            .checked_add(1)
            .ok_or_else(|| invalid("reference component count overflow"))?;
        if self.components > self.limits.max_components {
            return Err(limit_error(
                "reference components",
                self.components,
                self.limits.max_components,
            ));
        }
        Ok(())
    }

    fn peek_dollar_quote(&self) -> bool {
        self.peek() == Some(b'$') && self.bytes.get(self.position + 1) == Some(&b'\'')
    }

    fn expect_byte(&mut self, expected: u8, label: &str) -> Result<()> {
        if self.consume_byte(expected) {
            Ok(())
        } else {
            Err(invalid(format!("expected {label}")))
        }
    }

    fn consume_byte(&mut self, expected: u8) -> bool {
        if self.peek() == Some(expected) {
            self.position += 1;
            true
        } else {
            false
        }
    }

    fn peek(&self) -> Option<u8> {
        self.bytes.get(self.position).copied()
    }

    fn advance(&mut self) {
        self.position += 1;
    }

    fn at_end(&self) -> bool {
        self.position == self.bytes.len()
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cold]
fn limit_error(resource: &'static str, actual: usize, maximum: usize) -> Error {
    let Some(observed) = u64::try_from(actual).ok() else {
        return invalid("OpenFormula reference limit exceeds u64");
    };
    let Some(limit) = u64::try_from(maximum).ok() else {
        return invalid("OpenFormula reference limit exceeds u64");
    };
    let resource = if resource == "reference components" {
        Resource::Objects
    } else {
        Resource::InputBytes
    };
    Error::ResourceLimit(ResourceLimit {
        resource,
        observed,
        limit,
        scope: Arc::from("ods-formula-reference"),
    })
}

#[inline]
fn copy_component(value: &str, resource: &'static str) -> Result<String> {
    let mut result = String::new();
    result
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    result.push_str(value);
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cell_endpoint(reference: &Reference) -> &Endpoint {
        match reference.address().expect("non-error address") {
            Address::Cell(endpoint) => endpoint,
            other => panic!("expected cell address, got {other:?}"),
        }
    }

    #[test]
    fn parses_local_cell_and_preserves_absolute_metadata() {
        let reference = parse("[. $A$1]");
        assert!(reference.is_err(), "interior whitespace is not grammar");

        let reference = parse("[.$A$1]").expect("relative cell reference");
        let endpoint = cell_endpoint(&reference);
        assert!(matches!(endpoint.sheet, SheetSelector::Current));
        assert!(matches!(
            &endpoint.value,
            EndpointValue::Cell(cell)
                if cell.column.label == "A"
                    && cell.column.absolute
                    && cell.row.number == 1
                    && cell.row.absolute
        ));
    }

    #[test]
    fn parses_colon_rich_unquoted_sheet_name_without_range_split() {
        let reference = parse("[Sheet:Two.A1]").expect("colon is legal in unquoted sheet names");
        let endpoint = cell_endpoint(&reference);
        assert!(matches!(
            &endpoint.sheet,
            SheetSelector::Explicit(locator)
                if locator.sheet.name == "Sheet:Two" && locator.subtables.is_empty()
        ));
    }

    #[test]
    fn parses_inherited_and_cross_sheet_endpoints() {
        let inherited = parse("[$Inputs.$A$1:.$B$2]").expect("inherited endpoint");
        let Address::Cells(start, end) = inherited.address().expect("address") else {
            panic!("expected cell range");
        };
        assert!(matches!(
            start.sheet,
            SheetSelector::Explicit(ref locator)
                if locator.sheet.name == "Inputs" && locator.sheet.absolute
        ));
        assert!(matches!(end.sheet, SheetSelector::Inherited));

        let cross_sheet = parse("[First.A1:Second.B2]").expect("cross-sheet range");
        let Address::Cells(start, end) = cross_sheet.address().expect("address") else {
            panic!("expected cell range");
        };
        assert!(matches!(
            start.sheet,
            SheetSelector::Explicit(ref locator) if locator.sheet.name == "First"
        ));
        assert!(matches!(
            end.sheet,
            SheetSelector::Explicit(ref locator) if locator.sheet.name == "Second"
        ));
    }

    #[test]
    fn parses_whole_axis_ranges() {
        let columns = parse("[Sheet.A:.C]").expect("whole-column range");
        assert!(matches!(
            columns.address(),
            Some(Address::Columns(start, end))
                if matches!(start.value, EndpointValue::Column(ref column) if column.label == "A")
                    && matches!(end.sheet, SheetSelector::Inherited)
                    && matches!(end.value, EndpointValue::Column(ref column) if column.label == "C")
        ));

        let rows = parse("[Sheet.1:.4]").expect("whole-row range");
        assert!(matches!(
            rows.address(),
            Some(Address::Rows(start, end))
                if matches!(start.value, EndpointValue::Row(ref row) if row.number == 1)
                    && matches!(end.sheet, SheetSelector::Inherited)
                    && matches!(end.value, EndpointValue::Row(ref row) if row.number == 4)
        ));
    }

    #[test]
    fn parses_subtable_chain_and_quoted_names() {
        let reference = parse("['Bob''s'.A1.'Sub.Table'.B2]").expect("subtable locator");
        let endpoint = cell_endpoint(&reference);
        let SheetSelector::Explicit(locator) = &endpoint.sheet else {
            panic!("expected explicit locator");
        };
        assert_eq!(locator.sheet.name, "Bob's");
        assert!(locator.sheet.quoted);
        assert!(matches!(
            locator.subtables.as_slice(),
            [
                Subtable::Cell(Cell { column, row }),
                Subtable::Name(SheetName { name, quoted: true, .. })
            ] if column.label == "A" && row.number == 1 && name == "Sub.Table"
        ));
        assert!(matches!(
            &endpoint.value,
            EndpointValue::Cell(Cell { column, row })
                if column.label == "B" && row.number == 2
        ));
    }

    #[test]
    fn parses_source_and_keeps_reference_error_source_free() {
        let reference = parse("['file:///tmp/a''b.ods'#Sheet.A1]").expect("source reference");
        assert_eq!(
            reference.source().map(Source::as_str),
            Some("file:///tmp/a'b.ods")
        );
        assert!(matches!(reference.address(), Some(Address::Cell(_))));

        let error = parse("[#REF!]").expect("reference error");
        assert!(error.is_error());
        assert!(error.address().is_none());
        assert!(error.source().is_none());
        assert!(parse("['file:///tmp/a.ods'# #REF!]").is_err());
    }

    #[test]
    fn rejects_non_normative_columns_and_second_endpoint_without_dot() {
        assert!(parse("[.a1]").is_err());
        assert!(parse("[ .A1 ]").is_err());
        assert!(parse("[Sheet.A1:B2]").is_err());
        assert!(parse("[.A1:Sheet.B2]").is_err());
        assert!(parse("[Sheet.A1:.B]").is_err());
    }

    #[test]
    fn applies_byte_and_component_limits_before_component_allocation() {
        let limits = Limits::default().with_max_bytes(4);
        let error = parse_with_limits("[.A123]", &limits).expect_err("byte limit");
        assert!(
            matches!(error, Error::ResourceLimit(limit) if limit.resource == Resource::InputBytes)
        );

        let limits = Limits::default().with_max_components(1);
        let error = parse_with_limits("[.A1]", &limits).expect_err("component limit");
        assert!(
            matches!(error, Error::ResourceLimit(limit) if limit.resource == Resource::Objects)
        );
    }
}
