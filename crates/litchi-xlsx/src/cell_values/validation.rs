//! Closure checks for value-only cell publication.
//!
//! # The dependency rule
//!
//! Change [0602]'s D4, authorized by decision 7 of change 0652, replaced the
//! name allow-list this module used to carry with the property that
//! allow-list was a proxy for. A value-only rewrite composes exactly three
//! spans: the worksheet's `<dimension>` `ref` attribute, the `<sheetData>`
//! element, and — on the owner — the workbook's `<calcPr>`. Every other byte
//! of either part is copied from the source, unread and unaltered.
//!
//! So an element is admitted when one of two things is true:
//!
//! * the editor **models** it — it lies in a composed span, and its name,
//!   its place in the tree and the value it carries are ones the raw
//!   worksheet parser and the value-only writer both understand; or
//! * the editor **copies** it — it lies outside every composed span, so the
//!   rewrite reproduces its bytes exactly and never interprets them. Such an
//!   element is *unfamiliar*: it is admitted whatever it is called and
//!   whatever namespace it belongs to, together with its whole subtree.
//!
//! Attributes need no allow-list at all. The writer re-emits the source tag
//! of every element it touches and replaces only the attributes it owns
//! (`dimension/@ref`, and `@r`, `@t` and `@s` of an edited `<c>`), so every
//! other attribute survives an edit byte for byte. The one exception is a
//! relationship reference inside `<sheetData>`: a removed cell record takes
//! its attributes with it, so an `r:id` there could orphan a relationship the
//! rewrite is contracted to preserve. That is refused by name.
//!
//! Constructs whose *meaning* depends on the edited cell — pivot caches,
//! tables, cell metadata, shared and array formulas, merged ranges — are not
//! decided here. Shared strings are retained by the snapshot and their target
//! cells are guarded by the edit stage. The vocabulary cannot see the other
//! constructs; they are decided by the relationship and edit gates in
//! [`super::snapshot`](super::snapshot) and by the raw editor's own guards.
//!
//! [0602]: ../../../../docs/performance/0602-xlsx-real-producer-admission-design.md

use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result, allocation, invalid};
use crate::raw;

const TRANSITIONAL_SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";

#[cfg(test)]
#[path = "validation_borrow_tests.rs"]
mod borrow_tests;
#[cfg(test)]
#[path = "shared_traversal_tests.rs"]
mod shared_traversal_tests;

pub(super) fn workbook_xml(content: &[u8]) -> Result<()> {
    validate_xml(content, XmlOwner::Workbook)
}

pub(super) fn worksheet_xml(content: &[u8]) -> Result<()> {
    validate_xml(content, XmlOwner::Worksheet)
}

#[derive(Clone, Copy)]
enum XmlOwner {
    Workbook,
    Worksheet,
}

/// Validation state for the shared source traversal.
///
/// The state owns only the local element names required by the existing
/// closure checks. Event payloads remain borrowed for the duration of the
/// callback. On an observer failure the driver stops immediately and the
/// retained error is discarded before authoritative fallback validation.
struct Validator {
    owner: XmlOwner,
    depth: usize,
    elements: Vec<Box<[u8]>>,
    dialect: Option<Box<[u8]>>,
    /// Depth of the outermost element the rewrite copies verbatim, while one
    /// is open. Inside such a subtree nothing is modelled, so every element,
    /// namespace, attribute and text node is admitted.
    copied_from: Option<usize>,
    saw_root: bool,
    first_error: Option<Error>,
}

impl Validator {
    fn new(owner: XmlOwner) -> Self {
        Self {
            owner,
            depth: 0,
            elements: Vec::new(),
            dialect: None,
            copied_from: None,
            saw_root: false,
            first_error: None,
        }
    }

    fn observe(&mut self, namespace: &ResolveResult<'_>, event: &Event<'_>) -> bool {
        if self.first_error.is_some() {
            return false;
        }
        if let Err(error) = self.observe_inner(namespace, event) {
            self.first_error = Some(error);
            return false;
        }
        if matches!(event, Event::Eof) && (!self.saw_root || self.depth != 0) {
            self.first_error = Some(invalid("value-only XML has no complete root element"));
            return false;
        }
        true
    }

    /// Whether the element about to be entered lies inside a subtree the
    /// rewrite copies verbatim.
    const fn inside_copied(&self) -> bool {
        self.copied_from.is_some()
    }

    fn observe_inner(&mut self, namespace: &ResolveResult<'_>, event: &Event<'_>) -> Result<()> {
        match event {
            Event::Start(element) => {
                let inside_copied = self.inside_copied();
                let local = element
                    .name()
                    .local_name()
                    .as_ref()
                    .to_vec()
                    .into_boxed_slice();
                let admission = validate_element(
                    self.owner,
                    namespace,
                    element,
                    &local,
                    self.elements.last().map(AsRef::as_ref),
                    self.depth,
                    &mut self.dialect,
                    inside_copied,
                )?;
                self.saw_root = true;
                if admission == Admission::Copied && self.copied_from.is_none() {
                    self.copied_from = Some(self.depth);
                }
                // A copied subtree may nest as deeply as the input says, so
                // the validator now carries the parser's own depth bound
                // rather than inheriting one from a shallow vocabulary. The
                // raw parser refuses anything past it, so this cannot refuse
                // an input the pipeline would otherwise accept.
                if self.depth >= raw::worksheet::MAX_XML_DEPTH {
                    return Err(invalid("value-only XML nesting is too deep"));
                }
                self.depth = self
                    .depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("value-only XML depth overflow"))?;
                self.elements
                    .try_reserve(1)
                    .map_err(|source| allocation("value-only XML element stack", source))?;
                self.elements.push(local);
            },
            Event::Empty(element) => {
                let inside_copied = self.inside_copied();
                let local = element.name().local_name().as_ref().to_vec();
                validate_element(
                    self.owner,
                    namespace,
                    element,
                    &local,
                    self.elements.last().map(AsRef::as_ref),
                    self.depth,
                    &mut self.dialect,
                    inside_copied,
                )?;
                self.saw_root = true;
            },
            Event::End(element) => {
                self.depth = self
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("value-only XML has an unmatched closing element"))?;
                let expected = self
                    .elements
                    .pop()
                    .ok_or_else(|| invalid("value-only XML has no open element to close"))?;
                let closes_copied_root = self.copied_from == Some(self.depth);
                let modeled = !self.inside_copied();
                if closes_copied_root {
                    self.copied_from = None;
                }
                validate_close(
                    namespace,
                    element,
                    &expected,
                    self.dialect.as_deref(),
                    modeled,
                )?;
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "value-only edits refuse XML document type declarations",
                ));
            },
            Event::Eof => {},
            Event::Text(value) => {
                let decoded = value
                    .decode()
                    .map_err(|error| invalid(format!("invalid value-only XML text: {error}")))?;
                if !text_allowed(
                    self.owner,
                    self.elements.last().map(AsRef::as_ref),
                    self.inside_copied(),
                    &decoded,
                ) {
                    return Err(invalid(
                        "value-only XML has text outside a scalar value element",
                    ));
                }
            },
            Event::CData(value) => {
                let decoded = value
                    .decode()
                    .map_err(|error| invalid(format!("invalid value-only XML text: {error}")))?;
                if !text_allowed(
                    self.owner,
                    self.elements.last().map(AsRef::as_ref),
                    self.inside_copied(),
                    &decoded,
                ) {
                    return Err(invalid(
                        "value-only XML has text outside a scalar value element",
                    ));
                }
            },
            Event::GeneralRef(_) => {
                if !text_context_allowed(
                    self.owner,
                    self.elements.last().map(AsRef::as_ref),
                    self.inside_copied(),
                ) {
                    return Err(invalid(
                        "value-only XML has a reference outside a scalar value element",
                    ));
                }
            },
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
        Ok(())
    }

    fn finish(self) -> Result<()> {
        if let Some(error) = self.first_error {
            return Err(error);
        }
        if !self.saw_root || self.depth != 0 {
            return Err(invalid("value-only XML has no complete root element"));
        }
        Ok(())
    }
}

/// Validate and parse an eligible source using one borrowed `NsReader`.
///
/// Early reader, parser-transition, or provisional-cap uncertainty is
/// converted into a provisional failure. After the observer accepts the
/// complete source, parser materialization is carried as a `Complete` result;
/// the caller finishes validation before forwarding that result through the
/// raw facade. No speculative diagnostic escapes the established error order.
pub(super) fn worksheet_xml_and_parse_source<'a, F>(
    content: &[u8],
    admission: raw::worksheet::SourceAdmission,
    strings: F,
) -> Result<(crate::cell::Store, Option<raw::worksheet::SourceFacts>)>
where
    F: FnOnce() -> Result<Option<&'a [crate::cell::Text]>> + Copy,
{
    let mut validator = Validator::new(XmlOwner::Worksheet);
    let mut builder = raw::worksheet::FactsBuilder::new(content);
    let attempt = raw::worksheet::parse_source_with_observer(
        content,
        admission,
        strings,
        |namespace, event, span| {
            // The validator owns the first error and its precedence. The
            // fact builder only ever observes an event the validator has
            // already accepted, and it never stops the traversal: a builder
            // refusal drops the facts and nothing else.
            if !validator.observe(namespace, event) {
                return false;
            }
            builder.observe(namespace, event, span, content);
            true
        },
    );
    match attempt {
        raw::worksheet::SourceParseAttempt::Complete(parsed) => {
            validator.finish()?;
            let cells = raw::worksheet::complete_source_parse(content, parsed)?;
            // Facts are published only after validator EOF and a successful
            // raw finalization, so no provisional state can outlive a refusal.
            Ok((cells, builder.finish(content)))
        },
        raw::worksheet::SourceParseAttempt::ProvisionalFailed => {
            drop(validator);
            drop(builder);
            worksheet_xml(content)?;
            Ok((raw::worksheet::parse(content, strings)?, None))
        },
        raw::worksheet::SourceParseAttempt::ReaderFailed => {
            drop(validator);
            drop(builder);
            worksheet_xml(content)?;
            Ok((raw::worksheet::parse(content, strings)?, None))
        },
    }
}

/// Drive [`Validator`] over a complete part with an owning reader.
///
/// This is the authoritative pass. It shares every policy helper with the
/// borrowed traversal above, so the two cannot disagree about what is
/// admitted or about which refusal wins.
fn validate_xml(content: &[u8], owner: XmlOwner) -> Result<()> {
    let mut reader = NsReader::from_reader(content);
    let mut validator = Validator::new(owner);
    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("value-only XML scan failed: {error}")))?;
        // Events borrow the input slice; namespaces are inspected before the
        // next read advances the resolver's scope.
        let (namespace, event) = reader.resolver().resolve_event(event);
        let eof = matches!(event, Event::Eof);
        if !validator.observe(&namespace, &event) {
            break;
        }
        if eof {
            break;
        }
    }
    validator.finish()
}

/// How the value-only rewrite treats an element.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum Admission {
    /// The rewrite composes this element's span, or the editor reads a
    /// decision out of it, so the editor must model it.
    Modeled,
    /// The rewrite copies this element and its subtree verbatim and never
    /// reads them, so the editor need not know what it is.
    Copied,
}

/// The verdict of the dependency rule on one element.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Verdict {
    Modeled,
    Copied,
    /// A composed span may not carry a name the editor does not model.
    Unknown,
    /// A modelled name, but not where the editor models it.
    Misplaced,
}

/// Decide what the rewrite does with one element of a composed span.
///
/// The composed spans are the worksheet root tag, `<dimension>` and
/// `<sheetData>`, and the workbook root tag and its `<sheets>` catalog.
/// Everything reachable only through some other child of the root is copied,
/// whatever it is called.
fn classify(owner: XmlOwner, local: &[u8], parent: Option<&[u8]>) -> Verdict {
    let modeled = match owner {
        XmlOwner::Workbook => matches!(
            (parent, local),
            (None, b"workbook") | (Some(b"workbook"), b"sheets") | (Some(b"sheets"), b"sheet")
        ),
        XmlOwner::Worksheet => matches!(
            (parent, local),
            (None, b"worksheet")
                | (Some(b"worksheet"), b"dimension" | b"sheetData")
                | (Some(b"sheetData"), b"row")
                | (Some(b"row"), b"c")
                | (Some(b"c"), b"f" | b"v" | b"is")
                | (Some(b"is"), b"t" | b"r")
                | (Some(b"r"), b"t")
        ),
    };
    if modeled {
        return Verdict::Modeled;
    }
    // `rPr`, `rPh` and `phoneticPr` are the rich-text run's formatting and
    // phonetic children. The worksheet parser collects inline text only from
    // `is > t` and `is > r > t`, so nothing under them can change a cell's
    // value, and the rewrite copies an unedited cell record byte for byte.
    if matches!(owner, XmlOwner::Worksheet)
        && matches!(
            (parent, local),
            (Some(b"is"), b"rPh" | b"phoneticPr") | (Some(b"r"), b"rPr")
        )
    {
        return Verdict::Copied;
    }
    if in_composed_span(owner, parent) {
        return if modeled_name(owner, local) {
            Verdict::Misplaced
        } else {
            Verdict::Unknown
        };
    }
    Verdict::Copied
}

/// Whether every child of this parent lies in a span the rewrite composes.
///
/// `None` is the document root, which is composed on both owners: the editor
/// refuses a part whose root is not the one it was handed. A child of
/// `<worksheet>` or of `<workbook>` is deliberately absent: those are exactly
/// the places where an unfamiliar element is copied instead of refused.
const fn in_composed_span(owner: XmlOwner, parent: Option<&[u8]>) -> bool {
    match owner {
        XmlOwner::Workbook => matches!(parent, None | Some(b"sheets" | b"sheet")),
        XmlOwner::Worksheet => matches!(
            parent,
            None | Some(
                b"dimension" | b"sheetData" | b"row" | b"c" | b"is" | b"r" | b"f" | b"v" | b"t"
            )
        ),
    }
}

/// Whether this element lies inside the composed `<sheetData>` span.
const fn in_sheet_data(local: &[u8], parent: Option<&[u8]>) -> bool {
    matches!(local, b"sheetData")
        || matches!(parent, Some(b"sheetData" | b"row" | b"c" | b"is" | b"r"))
}

/// The names the editor models, used only to pick between the two refusals.
const fn modeled_name(owner: XmlOwner, local: &[u8]) -> bool {
    match owner {
        XmlOwner::Workbook => matches!(local, b"workbook" | b"sheets" | b"sheet"),
        XmlOwner::Worksheet => matches!(
            local,
            b"worksheet"
                | b"dimension"
                | b"sheetData"
                | b"row"
                | b"c"
                | b"f"
                | b"v"
                | b"is"
                | b"t"
                | b"r"
        ),
    }
}

/// Reject a malformed attribute list, and a relationship reference inside
/// the composed `<sheetData>` span.
///
/// `with_checks` is what makes a duplicated or unquoted attribute a refusal
/// here rather than a surprise in the parser. Every attribute that survives
/// this scan is admitted: the writer re-emits the source tag of each element
/// it touches and replaces only the attributes it owns, so an attribute it
/// does not know is preserved exactly.
///
/// The one exception is `refuse_relationships`, set inside `<sheetData>`.
/// Outside it a relationship reference is copied with its element and cannot
/// move; inside it a removed cell record would take the reference with it and
/// orphan the relationship the rewrite is contracted to preserve. `CT_Row`
/// and `CT_Cell` have no `r:id` attribute in either dialect, so no admissible
/// worksheet loses anything by this, and the check costs one `memchr` per
/// attribute of the span the editor composes.
fn scan_attributes(
    element: &BytesStart<'_>,
    local: &[u8],
    refuse_relationships: bool,
) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid value-only XML attribute: {error}")))?;
        if !refuse_relationships {
            continue;
        }
        let name = attribute.key.as_ref();
        let Some(position) = memchr::memchr(b':', name) else {
            continue;
        };
        let (prefix, suffix) = name.split_at(position);
        if suffix == b":id" && prefix != b"xml" {
            return Err(invalid(format!(
                "value-only edits refuse relationship reference '{}' on '{}'",
                String::from_utf8_lossy(name),
                String::from_utf8_lossy(local)
            )));
        }
    }
    Ok(())
}

/// Check one closing tag against the element it closes.
///
/// A modelled element must also close in the dialect its owner bound; a
/// copied one is only required to be balanced, which is what the source it is
/// copied from already guarantees.
fn validate_close(
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesEnd<'_>,
    expected: &[u8],
    dialect: Option<&[u8]>,
    modeled: bool,
) -> Result<()> {
    let names_match = element.local_name().as_ref() == expected;
    let dialect_matches = !modeled
        || matches!((namespace, dialect), (ResolveResult::Bound(Namespace(value)), Some(expected)) if *value == expected);
    if !names_match || !dialect_matches {
        return Err(invalid(
            "value-only XML has a mismatched or foreign closing element",
        ));
    }
    Ok(())
}

/// Refuse an element whose prefix no declaration binds.
///
/// A copied element may belong to any namespace, but a prefix with no binding
/// is not namespace-well-formed, and this module does not admit malformed
/// input just because it would be copied. The unprefixed, undeclared case
/// stays admitted inside a copied subtree: it is well-formed XML.
fn refuse_unbound_prefix(namespace: &ResolveResult<'_>) -> Result<()> {
    if matches!(namespace, ResolveResult::Unknown(_)) {
        return Err(invalid("value-only XML has an unbound element namespace"));
    }
    Ok(())
}

fn bind_dialect(namespace: &ResolveResult<'_>, dialect: &mut Option<Box<[u8]>>) -> Result<()> {
    let ResolveResult::Bound(Namespace(value)) = namespace else {
        return Err(invalid("value-only XML has an unbound element namespace"));
    };
    if *value != TRANSITIONAL_SML && *value != STRICT_SML {
        return Err(invalid("value-only XML has a foreign element namespace"));
    }
    match dialect {
        Some(expected) if expected.as_ref() != *value => {
            Err(invalid("value-only XML mixes SpreadsheetML dialects"))
        },
        Some(_) => Ok(()),
        None => {
            *dialect = Some(value.to_vec().into_boxed_slice());
            Ok(())
        },
    }
}

fn text_allowed(owner: XmlOwner, context: Option<&[u8]>, copied: bool, value: &str) -> bool {
    value.trim().is_empty() || text_context_allowed(owner, context, copied)
}

fn text_context_allowed(owner: XmlOwner, context: Option<&[u8]>, copied: bool) -> bool {
    // Character data the rewrite copies verbatim needs no context: it cannot
    // reach a composed span. Inside one, only the three scalar leaves the
    // parser reads may carry text.
    copied || (matches!(owner, XmlOwner::Worksheet) && matches!(context, Some(b"f" | b"v" | b"t")))
}

/// Refuse an element the editor neither models nor copies.
///
/// A copied element is admitted before its namespace is looked at, because
/// the rewrite reproduces its bytes without resolving a single name. A
/// modelled element is held to the established rules: one SpreadsheetML
/// dialect throughout, the modelled name in the modelled place, and no
/// relationship reference inside the composed `<sheetData>` span.
fn validate_element(
    owner: XmlOwner,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    local: &[u8],
    parent: Option<&[u8]>,
    depth: usize,
    dialect: &mut Option<Box<[u8]>>,
    inside_copied: bool,
) -> Result<Admission> {
    if inside_copied {
        refuse_unbound_prefix(namespace)?;
        scan_attributes(element, local, false)?;
        return Ok(Admission::Copied);
    }
    match classify(owner, local, parent) {
        Verdict::Copied => {
            refuse_unbound_prefix(namespace)?;
            scan_attributes(element, local, false)?;
            return Ok(Admission::Copied);
        },
        Verdict::Modeled => bind_dialect(namespace, dialect)?,
        Verdict::Unknown => {
            bind_dialect(namespace, dialect)?;
            return Err(invalid(format!(
                "value-only edits refuse dependency-bearing or unknown element '{}'",
                String::from_utf8_lossy(local)
            )));
        },
        Verdict::Misplaced => {
            bind_dialect(namespace, dialect)?;
            return Err(invalid(format!(
                "value-only edits refuse element '{}' in this XML context",
                String::from_utf8_lossy(local)
            )));
        },
    }
    scan_attributes(
        element,
        local,
        matches!(owner, XmlOwner::Worksheet) && in_sheet_data(local, parent),
    )?;
    let expected_root = match owner {
        XmlOwner::Workbook => b"workbook".as_slice(),
        XmlOwner::Worksheet => b"worksheet".as_slice(),
    };
    if depth == 0 && local != expected_root {
        return Err(invalid("value-only XML has the wrong root element"));
    }
    Ok(Admission::Modeled)
}
