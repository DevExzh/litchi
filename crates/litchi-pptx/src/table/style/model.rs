//! Public table-style model types and source-backed mutation methods.

use super::codec::{encode, parse_owned, scan};
use super::validation::{validate_def, validate_list, validate_name};
use super::{A, AS, MAX_STYLES, STRICT_REL, allocation, invalid, limit};
use crate::{Error, Result};
use bitflags::bitflags;
use litchi_opc::constants::relationship_type as rt;
use std::fmt::{self, Write as _};
use std::ops::Range;
use std::str::FromStr;

/// Namespace and relationship profile used by a table-style catalog.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(super) fn drawing(self) -> &'static str {
        match self {
            Self::Transitional => A,
            Self::Strict => AS,
        }
    }

    pub(super) fn relationship(self) -> &'static str {
        match self {
            Self::Transitional => rt::TABLE_STYLES,
            Self::Strict => STRICT_REL,
        }
    }

    pub(super) fn office_document(self) -> &'static str {
        match self {
            Self::Transitional => rt::OFFICE_DOCUMENT,
            Self::Strict => rt::STRICT_OFFICE_DOCUMENT,
        }
    }
}

/// A validated `DrawingML` table-style GUID stored without heap allocation.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Id([u8; 16]);

impl Id {
    /// Parse the required braced GUID wire form.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(value: &str) -> Result<Self> {
        value.parse()
    }

    pub(super) fn write_to(self, output: &mut String) -> Result<()> {
        output
            .try_reserve(38)
            .map_err(|source| allocation("table-style GUID encoding", source))?;
        write!(
            output,
            "{{{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            self.0[0],
            self.0[1],
            self.0[2],
            self.0[3],
            self.0[4],
            self.0[5],
            self.0[6],
            self.0[7],
            self.0[8],
            self.0[9],
            self.0[10],
            self.0[11],
            self.0[12],
            self.0[13],
            self.0[14],
            self.0[15],
        )
        .map_err(|_err| Error::Write)
    }
}

impl FromStr for Id {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        let bytes = value.as_bytes();
        if bytes.len() != 38 || bytes.first() != Some(&b'{') || bytes.last() != Some(&b'}') {
            return Err(invalid("table-style ID must be a braced GUID"));
        }
        for position in [9usize, 14, 19, 24] {
            if bytes.get(position) != Some(&b'-') {
                return Err(invalid("table-style ID has invalid GUID separators"));
            }
        }
        let mut decoded = [0u8; 16];
        let mut source = 1usize;
        for byte in &mut decoded {
            while matches!(source, 9 | 14 | 19 | 24) {
                source += 1;
            }
            let high = hex(*bytes
                .get(source)
                .ok_or_else(|| invalid("short table-style GUID"))?)
            .ok_or_else(|| invalid("table-style ID contains a non-hex digit"))?;
            let low = hex(*bytes
                .get(source + 1)
                .ok_or_else(|| invalid("short table-style GUID"))?)
            .ok_or_else(|| invalid("table-style ID contains a non-hex digit"))?;
            *byte = (high << 4) | low;
            source += 2;
        }
        Ok(Self(decoded))
    }
}

impl fmt::Display for Id {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "{{{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            self.0[0],
            self.0[1],
            self.0[2],
            self.0[3],
            self.0[4],
            self.0[5],
            self.0[6],
            self.0[7],
            self.0[8],
            self.0[9],
            self.0[10],
            self.0[11],
            self.0[12],
            self.0[13],
            self.0[14],
            self.0[15],
        )
    }
}

impl fmt::Debug for Id {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Display::fmt(self, formatter)
    }
}

bitflags! {
    /// Conditional table regions defined by one style, packed into two bytes.
    #[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
    pub struct Parts: u16 {
        const WHOLE = 1 << 0;
        const ODD_ROW = 1 << 1;
        const EVEN_ROW = 1 << 2;
        const ODD_COLUMN = 1 << 3;
        const EVEN_COLUMN = 1 << 4;
        const FIRST_COLUMN = 1 << 5;
        const LAST_COLUMN = 1 << 6;
        const FIRST_ROW = 1 << 7;
        const LAST_ROW = 1 << 8;
        const SOUTH_EAST = 1 << 9;
        const SOUTH_WEST = 1 << 10;
        const NORTH_EAST = 1 << 11;
        const NORTH_WEST = 1 << 12;
        const BACKGROUND = 1 << 13;
    }
}

impl Parts {
    /// Return the `DrawingML` element name for one single-region flag.
    #[must_use]
    pub fn xml_name(self) -> Option<&'static str> {
        PARTS
            .iter()
            .find_map(|(part, name)| (*part == self).then_some(*name))
    }

    pub(super) fn from_xml_name(name: &[u8]) -> Option<Self> {
        PARTS
            .iter()
            .find_map(|(part, candidate)| (candidate.as_bytes() == name).then_some(*part))
    }
}

pub(super) const PARTS: [(Parts, &str); 14] = [
    (Parts::BACKGROUND, "tblBg"),
    (Parts::WHOLE, "wholeTbl"),
    (Parts::ODD_ROW, "band1H"),
    (Parts::EVEN_ROW, "band2H"),
    (Parts::ODD_COLUMN, "band1V"),
    (Parts::EVEN_COLUMN, "band2V"),
    (Parts::LAST_COLUMN, "lastCol"),
    (Parts::FIRST_COLUMN, "firstCol"),
    (Parts::LAST_ROW, "lastRow"),
    (Parts::SOUTH_EAST, "seCell"),
    (Parts::SOUTH_WEST, "swCell"),
    (Parts::FIRST_ROW, "firstRow"),
    (Parts::NORTH_EAST, "neCell"),
    (Parts::NORTH_WEST, "nwCell"),
];

#[derive(Debug)]
pub(super) struct Attr {
    pub(super) name: String,
    pub(super) value: String,
}

#[derive(Debug)]
pub(super) enum Payload {
    Shared {
        raw: Range<usize>,
        body: Range<usize>,
        exact: bool,
    },
    Owned {
        xml: Vec<u8>,
        body: Range<usize>,
        exact: bool,
    },
}

/// One typed `a:tblStyle` definition.
///
/// Loaded definitions retain their complete inert XML payload. Renaming a
/// definition preserves that payload; [`Self::reset_parts`] deliberately
/// replaces detailed cell formatting with empty region declarations.
pub struct Def {
    pub(super) id: Id,
    pub(super) name: String,
    pub(super) parts: Parts,
    pub(super) attrs: Vec<Attr>,
    pub(super) payload: Payload,
}

impl fmt::Debug for Def {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Def")
            .field("id", &self.id)
            .field("name", &self.name)
            .field("parts", &self.parts)
            .finish_non_exhaustive()
    }
}

impl Def {
    /// Create a style with no conditional region payloads.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(id: Id, name: impl Into<String>) -> Result<Self> {
        let name = name.into();
        validate_name(&name)?;
        Ok(Self {
            id,
            name,
            parts: Parts::empty(),
            attrs: Vec::new(),
            payload: Payload::Owned {
                xml: Vec::new(),
                body: 0..0,
                exact: false,
            },
        })
    }

    #[must_use]
    pub fn id(&self) -> Id {
        self.id
    }

    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub fn parts(&self) -> Parts {
        self.parts
    }

    #[must_use]
    pub fn has(&self, parts: Parts) -> bool {
        self.parts.contains(parts)
    }

    /// Rename this detached definition while preserving its cell-style body.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn rename(&mut self, name: impl Into<String>) -> Result<String> {
        let name = name.into();
        validate_name(&name)?;
        if self.name == name {
            return Ok(name);
        }
        invalidate_payload(&mut self.payload);
        Ok(std::mem::replace(&mut self.name, name))
    }

    /// Replace detailed cell formatting with the selected empty regions.
    ///
    /// The explicit `reset` name makes the destructive payload change visible
    /// at the call site. Existing opaque formatting is otherwise preserved.
    pub fn reset_parts(&mut self, parts: Parts) -> Parts {
        let previous = self.parts;
        self.parts = parts;
        self.payload = Payload::Owned {
            xml: Vec::new(),
            body: 0..0,
            exact: false,
        };
        previous
    }

    fn materialize(&mut self, source: &[u8]) -> Result<()> {
        let Payload::Shared { raw, body, exact } = &self.payload else {
            return Ok(());
        };
        let raw_bytes = source
            .get(raw.clone())
            .ok_or_else(|| invalid("table-style source range is invalid"))?;
        let body_start = body
            .start
            .checked_sub(raw.start)
            .ok_or_else(|| invalid("table-style body precedes its element"))?;
        let body_end = body
            .end
            .checked_sub(raw.start)
            .ok_or_else(|| invalid("table-style body precedes its element"))?;
        let mut xml = Vec::new();
        xml.try_reserve_exact(raw_bytes.len())
            .map_err(|source| allocation("detached table-style XML", source))?;
        xml.extend_from_slice(raw_bytes);
        let exact = *exact;
        self.payload = Payload::Owned {
            xml,
            body: body_start..body_end,
            exact,
        };
        Ok(())
    }
}

/// Ordered table-style catalog (`a:tblStyleLst`).
///
/// The facade keeps GUIDs typed, region settings compact, and loaded XML
/// source-backed. An unchanged load→put moves the original producer bytes
/// back into OPC without normalization.
pub struct List {
    pub(super) conformance: Conformance,
    pub(super) default: Id,
    pub(super) defs: Vec<Def>,
    pub(super) root_attrs: Vec<Attr>,
    pub(super) source: Vec<u8>,
    pub(super) dirty: bool,
}

impl fmt::Debug for List {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("List")
            .field("conformance", &self.conformance)
            .field("default", &self.default)
            .field("defs", &self.defs)
            .field("source_bytes", &self.source.len())
            .field("dirty", &self.dirty)
            .finish()
    }
}

impl List {
    /// Create an empty catalog with an explicitly selected default style.
    #[must_use]
    pub fn new(conformance: Conformance, default: Id) -> Self {
        Self {
            conformance,
            default,
            defs: Vec::new(),
            root_attrs: Vec::new(),
            source: Vec::new(),
            dirty: true,
        }
    }

    /// Parse and take ownership of bounded table-style XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(xml: impl Into<Vec<u8>>) -> Result<Self> {
        parse_owned(xml.into())
    }

    #[must_use]
    pub fn conformance(&self) -> Conformance {
        self.conformance
    }

    #[must_use]
    pub fn default(&self) -> Id {
        self.default
    }

    pub fn set_default(&mut self, id: Id) -> Id {
        if self.default == id {
            return id;
        }
        self.dirty = true;
        std::mem::replace(&mut self.default, id)
    }

    #[must_use]
    pub fn styles(&self) -> &[Def] {
        &self.defs
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.defs.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.defs.is_empty()
    }

    /// Checked raw-position lookup for ordered inspection.
    #[must_use]
    pub fn at(&self, index: usize) -> Option<&Def> {
        self.defs.get(index)
    }

    /// Preferred stable-identity lookup.
    #[must_use]
    pub fn get(&self, id: Id) -> Option<&Def> {
        self.defs.iter().find(|style| style.id == id)
    }

    /// Return every definition with this non-identity display name.
    ///
    /// `DrawingML` permits duplicate and empty `styleName` values, so this
    /// method deliberately returns all matches rather than selecting one.
    pub fn named<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a Def> + 'a {
        self.defs.iter().filter(move |style| style.name == name)
    }

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add(&mut self, style: Def) -> Result<()> {
        validate_def(&style)?;
        self.ensure_unique_id(style.id, None)?;
        if self.defs.len() >= MAX_STYLES {
            return Err(limit("table-style count", MAX_STYLES));
        }
        self.defs
            .try_reserve(1)
            .map_err(|source| allocation("table-style insertion", source))?;
        self.defs.push(style);
        self.dirty = true;
        Ok(())
    }

    /// Rename one style by stable ID while retaining its opaque formatting.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn rename(&mut self, id: Id, name: impl Into<String>) -> Result<String> {
        let name = name.into();
        validate_name(&name)?;
        let style = self
            .defs
            .iter_mut()
            .find(|style| style.id == id)
            .ok_or_else(|| invalid(format!("table style {id} was not found")))?;
        if style.name == name {
            return Ok(name);
        }
        invalidate_payload(&mut style.payload);
        let previous = std::mem::replace(&mut style.name, name);
        self.dirty = true;
        Ok(previous)
    }

    /// Replace one style while retaining the selected stable GUID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(&mut self, id: Id, mut replacement: Def) -> Result<Def> {
        replacement.id = id;
        validate_def(&replacement)?;
        self.ensure_unique_id(id, Some(id))?;
        let style = self
            .defs
            .iter_mut()
            .find(|style| style.id == id)
            .ok_or_else(|| invalid(format!("table style {id} was not found")))?;
        style.materialize(&self.source)?;
        let previous = std::mem::replace(style, replacement);
        self.dirty = true;
        Ok(previous)
    }

    /// Remove one non-default style by stable GUID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove(&mut self, id: Id) -> Result<Option<Def>> {
        if id == self.default {
            return Err(invalid(
                "cannot remove the selected default table style; select another default first",
            ));
        }
        let Some(position) = self.defs.iter().position(|style| style.id == id) else {
            return Ok(None);
        };
        self.defs[position].materialize(&self.source)?;
        let removed = self.defs.remove(position);
        self.dirty = true;
        Ok(Some(removed))
    }

    /// Return original producer XML when the list has not been edited.
    #[must_use]
    pub fn source_xml(&self) -> Option<&[u8]> {
        (!self.dirty).then_some(self.source.as_slice())
    }

    /// Consume and encode the catalog, moving unchanged source bytes directly.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn into_xml(self) -> Result<Vec<u8>> {
        validate_list(&self)?;
        if !self.dirty {
            return Ok(self.source);
        }
        let xml = encode(&self)?;
        let parsed = scan(&xml)?;
        if parsed.conformance != self.conformance
            || parsed.default != self.default
            || parsed.defs.len() != self.defs.len()
            || parsed.defs.iter().zip(&self.defs).any(|(left, right)| {
                left.id != right.id || left.name != right.name || left.parts != right.parts
            })
        {
            return Err(invalid("encoded table-style catalog did not round-trip"));
        }
        Ok(xml)
    }

    fn ensure_unique_id(&self, id: Id, except: Option<Id>) -> Result<()> {
        if self
            .defs
            .iter()
            .any(|style| style.id == id && Some(style.id) != except)
        {
            return Err(invalid(format!("duplicate table-style ID {id}")));
        }
        Ok(())
    }
}
/// A borrowed, fully validated presentation relationship to a table-style
/// catalog.
#[derive(Clone, Copy, Debug)]
pub struct Link<'a> {
    pub(super) id: &'a str,
    pub(super) kind: &'a str,
    pub(super) target: &'a str,
}

impl<'a> Link<'a> {
    /// Return the exact relationship ID stored by the producer.
    #[must_use]
    pub fn id(self) -> &'a str {
        self.id
    }

    /// Return the exact Strict or Transitional relationship type.
    #[must_use]
    pub fn kind(self) -> &'a str {
        self.kind
    }

    /// Return the producer's unmodified relative target reference.
    #[must_use]
    pub fn target(self) -> &'a str {
        self.target
    }
}
fn hex(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

fn invalidate_payload(payload: &mut Payload) {
    match payload {
        Payload::Shared { exact, .. } | Payload::Owned { exact, .. } => *exact = false,
    }
}
