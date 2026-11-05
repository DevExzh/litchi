//! The property registry of a format, exported as data: [`schema()`].
//!
//! A [`Schema`] lists, for each element kind a format's resolver knows, the
//! properties its registry declares, each with
//!
//! - its OfficeCLI key (the name `get` prints and `set` accepts) and the
//!   other spellings `set`/`add` accept ([`Property::aliases`]);
//! - its value type in OfficeCLI's help-schema vocabulary, the canonical
//!   values and value aliases of an enumeration, the unit a bare number is
//!   read in for a length, and the color vocabulary for a color;
//! - its constraint ([`Constraint`]) and whether `null` clears it;
//! - its `get`/`set`/`add` flags;
//! - the gap and follow-up when the registry declares it but cannot write
//!   it yet ([`Property::write_gap`]), and the property gaps of the element
//!   (keys OfficeCLI has that Litchi does not answer yet).
//!
//! The export is generated from the same static tables the resolver, the
//! node view, the write admission and [`Capabilities`] run on, so it cannot
//! claim what the code does not do. It describes the format, not a file: no
//! open state, no retained-source or protection refusals (those are
//! [`File::capabilities`](crate::File::capabilities)).
//!
//! # JSON
//!
//! `Schema` serializes per element in OfficeCLI's help-schema shape
//! (`schemas/help/_schema.json`), with Litchi's additions marked below:
//!
//! ```json
//! {
//!   "format": "xlsx",
//!   "elements": [
//!     {
//!       "format": "xlsx",
//!       "element": "row",
//!       "operations": { "add": false, "set": true, "get": true, "query": true, "remove": false },
//!       "paths": { "positional": ["/<sheet>/row[N]"] },
//!       "properties": {
//!         "outline": {
//!           "type": "int",
//!           "aliases": ["group", "outlinelevel"],
//!           "constraint": { "min": 0, "max": 7 },
//!           "clearable": true,
//!           "add": false, "set": true, "get": false
//!         }
//!       }
//!     }
//!   ],
//!   "gaps": []
//! }
//! ```
//!
//! Litchi's members beyond OfficeCLI's: `valueAliases` (enumeration input
//! spellings), `unit` (lengths), `colorContext` (colors), `constraint`,
//! `clearable`, `writeGap` (per property); `propertyGaps`, `predicates` and
//! `gap` (per element); `gaps` (the format's element gaps without a table
//! entry, and its axis, operation and view gaps)
//! and `pending` (a format without an adapter in this build).
//!
//! The schema export is golden-tested per format
//! (`crates/litchi-ops/tests/golden/registry/<format>.json`, regenerated
//! only with `LITCHI_BLESS=1`), and the OfficeCLI delta table is generated
//! from those snapshots (`tools/officecli_capability_delta.py`).
//!
//! ```
//! use litchi_core::FileFormat;
//!
//! let schema = litchi_ops::schema(FileFormat::Docx);
//! assert_eq!(schema.format(), FileFormat::Docx);
//! # #[cfg(feature = "docx")] {
//! let run = schema.element("run").unwrap();
//! let bold = run.property("bold").unwrap();
//! assert_eq!(bold.value_type().schema_type(), "bool");
//! assert!(bold.get() && bold.set());
//! let json = serde_json::to_value(&schema).unwrap();
//! assert_eq!(json["format"], "docx");
//! # }
//! // A format whose adapter is not in this build has no elements.
//! if cfg!(not(feature = "pages")) {
//!     assert_eq!(litchi_ops::schema(FileFormat::Pages).elements().count(), 0);
//! }
//! ```

#![cfg_attr(
    not(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "xlsx",
        feature = "pptx",
        feature = "xlsb",
        feature = "xls",
        feature = "rtf",
        feature = "ods"
    )),
    allow(dead_code, unused_variables, unreachable_code)
)]

use litchi_core::FileFormat;
use serde::ser::{Serialize, SerializeMap, SerializeSeq, Serializer};

use crate::capabilities::{
    Capabilities, Constraint, Element, Elements, Gap, Gaps, Property, Subject, Verb,
};
use crate::read::Adapter;
use crate::registry::PropAlias;
use crate::value::{ColorContext, LengthContext, ValueType};

/// The registry export of `format`: every element kind with its declared
/// properties. See the [module documentation](self).
///
/// Built from static tables; it allocates nothing until serialized.
#[must_use]
pub fn schema(format: FileFormat) -> Schema {
    Schema {
        format,
        capabilities: Capabilities::of(format),
    }
}

/// The property registry of one format, from [`schema()`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Schema {
    format: FileFormat,
    capabilities: Capabilities,
}

impl Schema {
    /// The format described.
    #[must_use]
    pub const fn format(&self) -> FileFormat {
        self.format
    }

    /// Whether this build registers the format: `false` for a format whose
    /// adapter has not landed or whose cargo feature is off; its schema is
    /// then empty.
    #[must_use]
    pub fn registered(&self) -> bool {
        self.capabilities.verb(Verb::Get).is_supported()
    }

    /// The element kinds, in the format's table order.
    pub fn elements(&self) -> Elements {
        self.capabilities.elements()
    }

    /// The element named `name` (its OfficeCLI type or an alias).
    #[must_use]
    pub fn element(&self, name: &str) -> Option<Element> {
        self.capabilities.element(name)
    }

    /// Every gap of the format: element, axis, property, operation and view
    /// gaps, each with its follow-up.
    pub fn gaps(&self) -> Gaps {
        self.capabilities.gaps()
    }
}

/// The registry aliases of the properties of `adapter`'s element `canonical`
/// (its kind name): the alias projection of the element's write table.
///
/// Only tables declared with aliases appear here; a new aliased table adds
/// its arm (the `aliased_tables_are_exported` test lists them).
pub(crate) fn prop_aliases(adapter: Adapter, canonical: &str) -> &'static [PropAlias] {
    #[cfg(feature = "xlsx")]
    if adapter == Adapter::Xlsx {
        return match canonical {
            "cell" | "range" => XLSX_CELL_ALIASES,
            "run" => XLSX_RUN_ALIASES,
            "row" => XLSX_ROW_ALIASES,
            "col" => XLSX_COLUMN_ALIASES,
            _ => &[],
        };
    }
    #[cfg(feature = "pptx")]
    if adapter == Adapter::Pptx {
        return match canonical {
            "shape" | "placeholder" => crate::format::pptx::SHAPE_PROP_ALIASES,
            "paragraph" => crate::format::pptx::PARAGRAPH_PROP_ALIASES,
            "run" => crate::format::pptx::RUN_PROP_ALIASES,
            _ => &[],
        };
    }
    #[cfg(feature = "xlsb")]
    if adapter == Adapter::Xlsb {
        return match canonical {
            "cell" | "range" => XLSB_CELL_ALIASES,
            _ => &[],
        };
    }
    #[cfg(feature = "xls")]
    if adapter == Adapter::Xls {
        return match canonical {
            "cell" => XLS_CELL_ALIASES,
            _ => &[],
        };
    }
    #[cfg(feature = "ods")]
    if adapter == Adapter::Ods {
        return match canonical {
            "cell" | "range" => ODS_CELL_ALIASES,
            _ => &[],
        };
    }

    let _ = (adapter, canonical);
    &[]
}

#[cfg(feature = "xlsb")]
const XLSB_CELL_ALIASES: &[PropAlias] = &crate::format::xlsb::CELL_WRITE_TABLE.prop_aliases();
#[cfg(feature = "xls")]
const XLS_CELL_ALIASES: &[PropAlias] = &crate::format::xls::CELL_WRITE_TABLE.prop_aliases();
#[cfg(feature = "ods")]
const ODS_CELL_ALIASES: &[PropAlias] = &crate::format::ods::CELL_WRITE_TABLE.prop_aliases();

#[cfg(feature = "xlsx")]
const XLSX_CELL_ALIASES: &[PropAlias] = &crate::format::xlsx::CELL_WRITE_TABLE.prop_aliases();
#[cfg(feature = "xlsx")]
const XLSX_ROW_ALIASES: &[PropAlias] = &crate::format::xlsx::ROW_WRITE_TABLE.prop_aliases();
#[cfg(feature = "xlsx")]
const XLSX_RUN_ALIASES: &[PropAlias] = &crate::format::xlsx::RUN_WRITE_TABLE.prop_aliases();
#[cfg(feature = "xlsx")]
const XLSX_COLUMN_ALIASES: &[PropAlias] = &crate::format::xlsx::COLUMN_WRITE_TABLE.prop_aliases();

/// The unit a bare number of a length is read in.
const fn bare_unit(context: LengthContext) -> &'static str {
    match context {
        LengthContext::Emu | LengthContext::Extent => "emu",
        LengthContext::LineWidth | LengthContext::FontSize => "pt",
        LengthContext::WordSpacing => "twip",
    }
}

const fn color_context(context: ColorContext) -> &'static str {
    match context {
        ColorContext::Drawing => "drawing",
        ColorContext::Spreadsheet => "spreadsheet",
        ColorContext::Word => "word",
    }
}

/// The JSON name of a format (its usual extension).
pub(crate) const fn format_name(format: FileFormat) -> &'static str {
    match format {
        FileFormat::Doc => "doc",
        FileFormat::Docx => "docx",
        FileFormat::Ppt => "ppt",
        FileFormat::Pptx => "pptx",
        FileFormat::Xls => "xls",
        FileFormat::Xlsx => "xlsx",
        FileFormat::Xlsb => "xlsb",
        FileFormat::Rtf => "rtf",
        FileFormat::Pages => "pages",
        FileFormat::Keynote => "keynote",
        FileFormat::Numbers => "numbers",
        FileFormat::Odt => "odt",
        FileFormat::Ods => "ods",
        FileFormat::Odp => "odp",
        FileFormat::Odg => "odg",
        FileFormat::Odc => "odc",
        FileFormat::Odf => "odf",
        FileFormat::Odi => "odi",
        FileFormat::Odm => "odm",
        FileFormat::Oth => "oth",
        FileFormat::Odb => "odb",
        _ => "unknown",
    }
}

impl Serialize for Schema {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("format", format_name(self.format))?;
        if !self.registered() {
            map.serialize_entry("pending", &self.capabilities.verb(Verb::Get))?;
        }
        map.serialize_entry("elements", &ElementsJson(*self))?;
        map.serialize_entry("gaps", &FormatGapsJson(*self))?;
        if self.registered() {
            map.serialize_entry("setKeywords", &KeywordsJson(*self))?;
        }
        map.end()
    }
}

/// The keywords `set` takes beside an element's properties: text
/// find/replace (record 0273), a Litchi addition. OfficeCLI documents them
/// as cross-cutting `set` keywords outside its element schemas.
struct KeywordsJson(Schema);

impl Serialize for KeywordsJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(3))?;
        map.serialize_entry("verb", Verb::Replace.as_str())?;
        map.serialize_entry("support", &self.0.capabilities.verb(Verb::Replace))?;
        map.serialize_entry("keywords", &KeywordJson)?;
        map.end()
    }
}

struct KeywordJson;

impl Serialize for KeywordJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(crate::find::KEYWORDS.len()))?;
        for keyword in &crate::find::KEYWORDS {
            map.serialize_entry(keyword.name, &KeywordEntry(keyword))?;
        }
        map.end()
    }
}

/// One keyword in the property shape of [`PropertyJson`].
struct KeywordEntry(&'static crate::find::Keyword);

impl Serialize for KeywordEntry {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let keyword = self.0;
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("type", keyword.ty)?;
        if keyword.not_empty {
            map.serialize_entry(
                "constraint",
                &[("notEmpty", true)].into_iter().collect::<Pairs<_>>(),
            )?;
        }
        if let Some(default) = keyword.default {
            map.serialize_entry("default", &default)?;
        }
        map.serialize_entry("set", &true)?;
        if let Some((follow_up, reason)) = keyword.gap {
            map.serialize_entry(
                "writeGap",
                &[
                    ("code", crate::Code::UnsupportedCapability.as_str()),
                    ("followUp", follow_up),
                    ("reason", reason),
                ]
                .into_iter()
                .collect::<Pairs<_>>(),
            )?;
        }
        map.end()
    }
}

/// A JSON array of the items of a clonable iterator.
struct Seq<I>(I);

impl<I> Serialize for Seq<I>
where
    I: Iterator + Clone,
    I::Item: Serialize,
{
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(None)?;
        for item in self.0.clone() {
            seq.serialize_element(&item)?;
        }
        seq.end()
    }
}

struct ElementsJson(Schema);

impl Serialize for ElementsJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(None)?;
        for element in self.0.elements() {
            seq.serialize_element(&ElementJson {
                schema: self.0,
                element,
            })?;
        }
        seq.end()
    }
}

/// One element in OfficeCLI's help-schema shape, plus Litchi's members.
struct ElementJson {
    schema: Schema,
    element: Element,
}

impl Serialize for ElementJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let element = self.element;
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("format", format_name(self.schema.format))?;
        map.serialize_entry("element", element.name())?;
        let mut parents = element.parents();
        match (parents.next(), parents.next()) {
            (None, _) => {},
            (Some(parent), None) => map.serialize_entry("parent", parent.name())?,
            (Some(_), Some(_)) => {
                map.serialize_entry("parent", &Seq(element.parents().map(|p| p.name())))?;
            },
        }
        if element.aliases().next().is_some() {
            map.serialize_entry("elementAliases", &Seq(element.aliases()))?;
        }
        let operations = element.operations();
        map.serialize_entry(
            "operations",
            &[
                ("add", operations.add),
                ("set", operations.set),
                ("get", operations.get),
                ("query", operations.query),
                ("remove", operations.remove),
            ]
            .into_iter()
            .collect::<Pairs<_>>(),
        )?;
        map.serialize_entry("paths", &PathsJson(element))?;
        if element.children().next().is_some() {
            map.serialize_entry("children", &ChildrenJson(element))?;
        }
        map.serialize_entry("properties", &PropertiesJson(element))?;
        let gaps = PropertyGapsJson {
            schema: self.schema,
            element: element.name(),
        };
        if gaps.iter().next().is_some() {
            map.serialize_entry("propertyGaps", &gaps)?;
        }
        if operations.query {
            map.serialize_entry("predicates", &PredicatesJson(element))?;
        }
        if let Some(gap) = element.gap() {
            map.serialize_entry("gap", &GapJson(gap))?;
        }
        map.end()
    }
}

/// A JSON object of `(key, value)` pairs in order.
struct Pairs<V>(Vec<(&'static str, V)>);

impl<V> FromIterator<(&'static str, V)> for Pairs<V> {
    fn from_iter<T: IntoIterator<Item = (&'static str, V)>>(iter: T) -> Self {
        Self(iter.into_iter().collect())
    }
}

impl<V: Serialize> Serialize for Pairs<V> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(Some(self.0.len()))?;
        for (key, value) in &self.0 {
            map.serialize_entry(key, value)?;
        }
        map.end()
    }
}

struct PathsJson(Element);

impl Serialize for PathsJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let element = self.0;
        let mut map = serializer.serialize_map(None)?;
        for (key, positional) in [("stable", false), ("positional", true)] {
            let mut forms = element
                .paths()
                .filter(move |(_, is)| *is == positional)
                .map(|(template, _)| template)
                .peekable();
            if forms.peek().is_some() {
                map.serialize_entry(key, &Seq(forms))?;
            }
        }
        map.end()
    }
}

struct ChildrenJson(Element);

impl Serialize for ChildrenJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(None)?;
        for child in self.0.children() {
            let mut entry = vec![("element", child.name())];
            if let Some(segment) = child.path_segment() {
                entry.push(("pathSegment", segment));
            }
            seq.serialize_element(&Pairs(entry))?;
        }
        seq.end()
    }
}

struct PropertiesJson(Element);

impl Serialize for PropertiesJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(None)?;
        for property in self.0.properties() {
            map.serialize_entry(property.name(), &PropertyJson(property))?;
        }
        map.end()
    }
}

/// A property as OfficeCLI's help schema writes it (`type`, `aliases`,
/// `values`, then `add`/`set`/`get`), with the registry's additions.
struct PropertyJson(Property);

impl Serialize for PropertyJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let property = self.0;
        let ty = property.value_type();
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("type", ty.schema_type())?;
        if property.aliases().next().is_some() {
            map.serialize_entry("aliases", &Seq(property.aliases()))?;
        }
        match ty {
            ValueType::Token(tokens) => {
                map.serialize_entry("values", tokens.values())?;
                if !tokens.alias_pairs().is_empty() {
                    map.serialize_entry(
                        "valueAliases",
                        &tokens
                            .alias_pairs()
                            .iter()
                            .map(|&(alias, index)| (alias, tokens.values()[usize::from(index)]))
                            .collect::<Pairs<_>>(),
                    )?;
                }
            },
            ValueType::Length(spec) => map.serialize_entry("unit", bare_unit(spec.context))?,
            ValueType::Color(context) => {
                map.serialize_entry("colorContext", color_context(context))?;
            },
            _ => {},
        }
        if property.constraint() != Constraint::None {
            map.serialize_entry("constraint", &ConstraintJson(property.constraint()))?;
        }
        if property.clearable() {
            map.serialize_entry("clearable", &true)?;
        }
        map.serialize_entry("add", &property.add())?;
        map.serialize_entry("set", &property.set())?;
        map.serialize_entry("get", &property.get())?;
        if let Some(gap) = property.write_gap() {
            map.serialize_entry("writeGap", &GapJson(gap))?;
        }
        map.end()
    }
}

struct ConstraintJson(Constraint);

impl Serialize for ConstraintJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(None)?;
        match self.0 {
            Constraint::Int { min, max } => {
                map.serialize_entry("min", &min)?;
                map.serialize_entry("max", &max)?;
            },
            Constraint::Number { min, max } => {
                map.serialize_entry("min", &min)?;
                map.serialize_entry("max", &max)?;
            },
            Constraint::Length { min, max } => {
                map.serialize_entry("minEmu", &min.emu())?;
                map.serialize_entry("maxEmu", &max.emu())?;
            },
            Constraint::Chars { max } => map.serialize_entry("maxChars", &max)?,
            Constraint::NotEmpty => map.serialize_entry("notEmpty", &true)?,
            Constraint::None => {},
        }
        map.end()
    }
}

/// The property gaps of one element, keyed by property.
struct PropertyGapsJson {
    schema: Schema,
    element: &'static str,
}

impl PropertyGapsJson {
    fn iter(&self) -> impl Iterator<Item = (&'static str, Gap)> + use<> {
        let element = self.element;
        self.schema
            .gaps()
            .filter_map(move |gap| match gap.subject() {
                Subject::Property {
                    element: owner,
                    key,
                } if owner == element => Some((key, gap)),
                _ => None,
            })
    }
}

impl Serialize for PropertyGapsJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut map = serializer.serialize_map(None)?;
        for (key, gap) in self.iter() {
            map.serialize_entry(key, &GapJson(gap))?;
        }
        map.end()
    }
}

struct PredicatesJson(Element);

impl Serialize for PredicatesJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let element = self.0;
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("keys", &Seq(element.predicate_keys()))?;
        if element.predicate_aliases().next().is_some() {
            map.serialize_entry(
                "aliases",
                &element.predicate_aliases().collect::<Pairs<_>>(),
            )?;
        }
        if element.header_columns() {
            map.serialize_entry("headerColumns", &true)?;
        }
        map.end()
    }
}

/// A gap's code, follow-up and reason.
struct GapJson(Gap);

impl Serialize for GapJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let gap = self.0;
        let mut map = serializer.serialize_map(Some(3))?;
        map.serialize_entry("code", gap.code().as_str())?;
        map.serialize_entry("followUp", gap.follow_up())?;
        map.serialize_entry("reason", gap.reason())?;
        map.end()
    }
}

/// The format's gaps that belong to no element entry: axes, operation
/// contexts and view modes, each with its subject.
struct FormatGapsJson(Schema);

impl Serialize for FormatGapsJson {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let mut seq = serializer.serialize_seq(None)?;
        for gap in self.0.gaps() {
            let subject: Vec<(&'static str, &'static str)> = match gap.subject() {
                // Reported with the element (`propertyGaps`).
                Subject::Property { .. } => continue,
                // An element's own gap is reported with the element (`gap`);
                // an element the table does not declare stays listed here.
                Subject::Element(name) => {
                    if self
                        .0
                        .element(name)
                        .is_some_and(|element| element.gap() == Some(gap))
                    {
                        continue;
                    }
                    vec![("subject", "element"), ("element", name)]
                },
                Subject::Axis(axis) => vec![("subject", "axis"), ("axis", axis)],
                Subject::Operation {
                    verb,
                    element,
                    context,
                } => vec![
                    ("subject", "operation"),
                    ("verb", verb),
                    ("element", element),
                    ("context", context),
                ],
                Subject::View(mode) => vec![("subject", "view"), ("mode", mode)],
            };
            let mut entry = subject;
            entry.push(("code", gap.code().as_str()));
            entry.push(("followUp", gap.follow_up()));
            entry.push(("reason", gap.reason()));
            seq.serialize_element(&Pairs(entry))?;
        }
        seq.end()
    }
}

#[cfg(test)]
mod tests;
