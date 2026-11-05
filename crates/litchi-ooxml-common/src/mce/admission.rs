//! Proving, without processing a part, that markup-compatibility processing
//! leaves everything a reader sees unchanged.
//!
//! [`process_markup_compatibility`](super::process_markup_compatibility)
//! rewrites every part that binds the markup-compatibility (MCE) namespace,
//! even when the rewrite only changes how the part is spelled. An editor that
//! works on the source bytes of such a part while readers see the processed
//! projection is safe only where the two agree on everything either of them
//! reads. [`admit`] proves that agreement with one bounded pass that builds
//! no output, for the shapes producers actually write:
//!
//! * a part that never binds the namespace, which processing returns as is
//!   ([`Admission::Unbound`]);
//! * a part that binds it and uses nothing in it, which processing only
//!   respells: whitespace outside the document element goes, empty-element
//!   tags become start and end tags, and a rebuilt start tag is written with
//!   single spaces and double quotes ([`Admission::DeclaredOnly`]);
//! * a part whose only constructs are valid `Ignorable` directives, one on
//!   the document element and any below it listing only understood
//!   namespaces, and no element or attribute is in an ignorable namespace
//!   that the processing profile does not understand
//!   ([`Admission::Ignorable`]): processing also removes those attributes,
//!   nothing else;
//! * as above, except that processing also drops attributes, and elements
//!   with their subtrees, in ignorable namespaces the profile does not
//!   understand, none of which holds an element or attribute of a namespace
//!   the profile understands ([`Admission::IgnorableOpaque`]).
//!
//! Everything else is a [`Refusal`]: an element in the namespace
//! (`AlternateContent`, `Choice`, `Fallback`), any other directive
//! (`ProcessContent`, `MustUnderstand`, `PreserveElements`,
//! `PreserveAttributes`), an `Ignorable` below the document element that
//! lists a namespace the profile does not understand, a dropped subtree that
//! holds understood content, and every input on which
//! processing would fail or that the proof does not model. A refusal is never
//! a verdict about the document; it only says that the proof does not hold,
//! so a caller that needs the processor's own answer can still run it.
//!
//! [`admit_held`] proves a weaker agreement for one more shape: a part whose
//! `mc:AlternateContent` wrappers all lie inside a [`Holder`] element (a
//! WordprocessingML run, for a main document). Each such wrapper is opaque:
//! processing replaces it, in place inside the holder, with the content of
//! the branch it selects, and the proof does not examine that content.
//! Everything outside the wrappers agrees as [`admit`] proves it.
//!
//! # Agreement with the processor
//!
//! The proof admits a part only when processing it would succeed. It repeats
//! every check the processor makes on the admitted shapes, with the
//! processor's own namespace scope, its name rules and its
//! limits: well-formed XML with one document element, UTF-8 names and
//! values, qualified names with bound prefixes, no two attributes with the
//! same expanded name, valid namespace declarations within the binding
//! limit, a valid `Ignorable` list, only predefined or character references
//! in visible text, no document type declaration or processing instruction,
//! no XML declaration after the document element starts, and the depth,
//! attribute, token, input and output bounds. Where it is stricter than the
//! processor it refuses (for example on a declaration that rebinds `xml`),
//! which costs a caller a fallback but never a wrong admission.
//!
//! Every decision is made on decoded names. A namespace declaration whose
//! value spells the namespace with character or entity references binds it
//! exactly as a literal one does, here and in the processor, whose fast path
//! uses [`mentions_namespace`].

use std::borrow::Cow;
use std::str;

use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::{Reader, XmlVersion, encoding::Decoder};

use super::model::{ATTRIBUTES_PER_ELEMENT_CEILING, Capabilities, Limits, NAMESPACE, XML_NS};
use super::scope::{NO_NAMESPACE, Scope, UriId, has_duplicate_prefix};
use crate::xml::attributes::BytesStartExt as _;
use crate::xml_name;

/// How markup-compatibility processing would treat an admitted part.
///
/// Every variant guarantees that processing the part would succeed, and that
/// the part and its projection agree on every element and attribute of every
/// namespace the processing profile understands, in document order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Admission {
    /// The part does not mention the markup-compatibility namespace, even
    /// through references ([`mentions_namespace`]), so processing returns the
    /// part itself.
    Unbound,
    /// The part mentions the namespace, and no element or attribute is in it:
    /// it is bound and unused, or only mentioned. Processing changes nothing
    /// but spelling.
    DeclaredOnly,
    /// The only constructs are valid `Ignorable` directives, and processing
    /// drops nothing but them: the document element's lists namespaces the
    /// profile understands or that the part never uses, and any below it
    /// lists only understood ones. Processing removes those attributes and
    /// changes nothing but spelling.
    Ignorable,
    /// As [`Self::Ignorable`], except that processing also drops attributes,
    /// and elements with their subtrees, in listed namespaces the profile
    /// does not understand. No dropped subtree holds an element or attribute
    /// of a namespace the profile understands, so a reader that skips
    /// foreign markup sees the same content in the part and its projection.
    IgnorableOpaque,
}

impl Admission {
    /// Whether the projection differs from the part only in spelling and in
    /// the document element's `Ignorable` attribute: no other element or
    /// attribute is dropped.
    #[must_use]
    pub const fn is_lexical(self) -> bool {
        !matches!(self, Self::IgnorableOpaque)
    }
}

/// Why [`admit`] could not prove a part admissible.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Refusal {
    /// An element is in the markup-compatibility namespace:
    /// `AlternateContent`, `Choice`, `Fallback`, or another name.
    Element,
    /// A directive other than `Ignorable`: `ProcessContent`,
    /// `MustUnderstand`, `PreserveElements` or `PreserveAttributes`.
    Directive,
    /// `Ignorable` below the document element listing a namespace the
    /// profile does not understand. It would make content ignorable inside
    /// one subtree only, so moving that content would change whether
    /// processing drops it.
    NestedIgnorable,
    /// Processing would drop an element or attribute of a namespace the
    /// profile understands, inside an ignorable subtree, or would drop the
    /// document element itself.
    HiddenContent,
    /// Processing would fail or exceed a limit, the input is malformed, or it
    /// uses something the proof does not model (such as extension elements
    /// in the profile).
    Unproven,
}

impl Refusal {
    /// A short stable description, for diagnostics.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Element => "markup-compatibility element",
            Self::Directive => "markup-compatibility directive other than Ignorable",
            Self::NestedIgnorable => {
                "Ignorable below the document element listing a namespace that is not understood"
            },
            Self::HiddenContent => "ignorable markup would hide understood content",
            Self::Unproven => "markup-compatibility processing is not proven to succeed unchanged",
        }
    }
}

/// An element whose subtree may hold opaque `mc:AlternateContent`
/// wrappers, for [`admit_held`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Holder<'h> {
    /// The namespaces the element may be in.
    pub namespaces: &'h [&'h str],
    /// The element's local name.
    pub local: &'h str,
}

/// What [`admit_held`] proved about a part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Held {
    /// The verdict on everything outside the held wrappers, with the same
    /// meaning as [`admit`]'s.
    pub outside: Admission,
    /// How many `mc:AlternateContent` wrappers lie inside holders, counting
    /// only the outermost ones. When it is zero the verdict is [`admit`]'s
    /// own.
    pub wrappers: usize,
}

/// Whether `xml` mentions the markup-compatibility namespace: its literal
/// bytes occur anywhere, or a namespace declaration spells it with
/// references.
///
/// Processing returns its input unchanged exactly when this is `false`. The
/// literal test is the processor's historical fast path, kept as it was: a
/// mention outside any declaration, or in malformed markup, still sends the
/// part to the parser, which rewrites only its spelling or refuses it. The
/// second test decides on decoded names. A binding needs a declaration, an
/// attribute named `xmlns` or `xmlns:` followed by a prefix, and names cannot
/// be written with references, so every declaration holds the bytes `xmlns`;
/// a declaration can spell the namespace without its literal bytes only with
/// a reference, so only a part holding `&` is searched for one. Each
/// occurrence of `xmlns` is read as the attribute it would start, and its
/// value counts when quick-xml decodes it to the namespace. An occurrence
/// that only reads like a declaration, in text or a comment, costs the
/// caller its slower path, never a missed binding.
#[must_use]
pub fn mentions_namespace(xml: &[u8]) -> bool {
    memchr::memmem::find(xml, NAMESPACE.as_bytes()).is_some() || binds_namespace_by_reference(xml)
}

/// Whether a namespace declaration in `xml` spells the markup-compatibility
/// namespace with references, so that its literal bytes never occur.
///
/// A reader that resolves names on the values as written (quick-xml's
/// namespace resolver does) does not see such a binding, while processing
/// applies it. A scanner that edits source bytes and recognizes the
/// namespace that way refuses a part for which this is `true`.
#[must_use]
pub fn binds_namespace_by_reference(xml: &[u8]) -> bool {
    memchr::memchr(b'&', xml).is_some() && escaped_binding(xml)
}

/// Whether a namespace declaration whose value holds a reference decodes to
/// the namespace.
fn escaped_binding(xml: &[u8]) -> bool {
    memchr::memmem::find_iter(xml, b"xmlns").any(|start| {
        declaration_value(xml, start + b"xmlns".len())
            .is_some_and(|value| value.contains(&b'&') && decodes_to_namespace(value))
    })
}

/// The value, as written between its quotes, of the declaration whose name
/// continues at `cursor`, right after `xmlns`; `None` when the bytes there
/// cannot be the rest of a declaration.
fn declaration_value(xml: &[u8], mut cursor: usize) -> Option<&[u8]> {
    if xml.get(cursor) == Some(&b':') {
        cursor += 1;
        while xml
            .get(cursor)
            .is_some_and(|byte| !ends_attribute_name(*byte))
        {
            cursor += 1;
        }
    }
    cursor = skip_xml_whitespace(xml, cursor);
    if xml.get(cursor) != Some(&b'=') {
        return None;
    }
    cursor = skip_xml_whitespace(xml, cursor + 1);
    let quote = *xml
        .get(cursor)
        .filter(|byte| matches!(**byte, b'"' | b'\''))?;
    let tail = xml.get(cursor + 1..)?;
    // An unterminated value binds nothing; the part is not XML.
    memchr::memchr(quote, tail).map(|length| &tail[..length])
}

fn ends_attribute_name(byte: u8) -> bool {
    matches!(
        byte,
        b'=' | b' ' | b'\t' | b'\r' | b'\n' | b'>' | b'/' | b'<' | b'"' | b'\''
    )
}

fn skip_xml_whitespace(xml: &[u8], mut cursor: usize) -> usize {
    while xml
        .get(cursor)
        .is_some_and(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
    {
        cursor += 1;
    }
    cursor
}

/// Whether an attribute value, as written between its quotes, decodes to
/// the namespace.
///
/// The value is decoded by quick-xml's own reference decoding, the one the
/// processor's attribute normalization uses, so the two can never disagree
/// about which spellings name the namespace. A value quick-xml cannot decode
/// binds nothing: the processor rejects it. Normalization only turns
/// whitespace into spaces, and the namespace has neither.
fn decodes_to_namespace(value: &[u8]) -> bool {
    if !value.contains(&b'&') {
        return value == NAMESPACE.as_bytes();
    }
    str::from_utf8(value)
        .ok()
        .and_then(|text| quick_xml::escape::unescape(text).ok())
        .is_some_and(|decoded| decoded == NAMESPACE)
}

/// Prove how processing with `capabilities` and `limits` would treat `xml`.
///
/// # Errors
///
/// Returns the [`Refusal`] that keeps the proof from holding.
pub fn admit(
    xml: &[u8],
    capabilities: &Capabilities,
    limits: &Limits,
) -> Result<Admission, Refusal> {
    match admit_with(xml, capabilities, limits, |_| {
        Ok::<(), core::convert::Infallible>(())
    }) {
        Ok(verdict) => verdict,
        Err(never) => match never {},
    }
}

/// [`admit`], calling `observe` with the byte length of each XML event before
/// the event is examined, so that a budgeted caller can charge the work or
/// stop it. An error from `observe` ends the proof and is returned as is.
///
/// The proof reads nothing past the event `observe` last accepted, and holds
/// at most the processor's namespace scope, one entry per open element and
/// the declarations of one start tag.
///
/// # Errors
///
/// Returns `observe`'s error; the inner result is the verdict.
pub fn admit_with<E>(
    xml: &[u8],
    capabilities: &Capabilities,
    limits: &Limits,
    observe: impl FnMut(usize) -> Result<(), E>,
) -> Result<Result<Admission, Refusal>, E> {
    if xml.len() > limits.max_input_bytes || xml.len() > limits.max_output_bytes {
        return Ok(Err(Refusal::Unproven));
    }
    if !mentions_namespace(xml) {
        return Ok(Ok(Admission::Unbound));
    }
    // Extension elements are copied opaquely by the processor; the proof does
    // not model them. The processor's output is bounded by its input with
    // every value byte re-escaped and every empty tag expanded: eight bytes
    // per input byte bounds both.
    if !capabilities.extensions.is_empty()
        || xml.len().saturating_mul(MAX_OUTPUT_GROWTH) > limits.max_output_bytes
    {
        return Ok(Err(Refusal::Unproven));
    }
    Prover::new(xml, capabilities, limits, None)
        .run(observe)
        .map(|verdict| verdict.map(|held| held.outside))
}

/// Prove how processing with `capabilities` and `limits` would treat `xml`
/// outside the `mc:AlternateContent` wrappers that lie inside a `holder`
/// element.
///
/// A wrapper with a holder ancestor, outside any subtree processing drops,
/// is opaque: the declarations on its own start tag are read, and nothing in
/// it is examined beyond well-formedness, the depth bound and the reference
/// check [`admit`] makes on visible text. Every other markup-compatibility
/// element is refused as [`admit`] refuses it, and so is everything else
/// [`admit`] refuses.
///
/// An admitted part agrees with its projection, as [`Admission`] states, on
/// everything outside the wrappers. Processing replaces each wrapper, in
/// place inside its holder, with the content of the branch it selects (or
/// with nothing), so every element outside the holders keeps its parent and
/// its order. The proof does **not** establish that processing succeeds
/// inside a wrapper (a `Requires` naming an unbound prefix, for example, is
/// a processing error): when [`Held::wrappers`] is not zero, a caller runs
/// the processor, or has run it, before it relies on the verdict.
///
/// # Errors
///
/// Returns the [`Refusal`] that keeps the proof from holding.
pub fn admit_held(
    xml: &[u8],
    capabilities: &Capabilities,
    limits: &Limits,
    holder: Holder<'_>,
) -> Result<Held, Refusal> {
    if xml.len() > limits.max_input_bytes || xml.len() > limits.max_output_bytes {
        return Err(Refusal::Unproven);
    }
    if !mentions_namespace(xml) {
        return Ok(Held {
            outside: Admission::Unbound,
            wrappers: 0,
        });
    }
    if !capabilities.extensions.is_empty()
        || xml.len().saturating_mul(MAX_OUTPUT_GROWTH) > limits.max_output_bytes
    {
        return Err(Refusal::Unproven);
    }
    match Prover::new(xml, capabilities, limits, Some(holder))
        .run(|_| Ok::<(), core::convert::Infallible>(()))
    {
        Ok(verdict) => verdict,
        Err(never) => match never {},
    }
}

/// The most output bytes processing can write for one input byte of an
/// admitted part.
const MAX_OUTPUT_GROWTH: usize = 8;

/// One open element.
#[derive(Clone, Copy)]
struct Frame {
    /// The processor's count of namespace bindings in scope.
    bindings: usize,
    /// Whether the element declared namespaces, so its end changes the scope.
    declares: bool,
    /// Whether processing drops this element (and so its whole subtree).
    dropped: bool,
    /// Whether the element is a holder or lies inside one.
    held: bool,
}

/// What the proof asks about the namespace one name resolves to.
#[derive(Clone, Copy)]
struct Class {
    id: UriId,
    /// The markup-compatibility namespace.
    mce: bool,
    /// Understood by the profile: processing keeps it wherever it is visible.
    understood: bool,
    /// The `xml` namespace, under any prefix.
    xml: bool,
    /// Made ignorable by the document element and not understood: dropped.
    dropping: bool,
}

/// The classes of recently resolved short prefixes, cleared whenever the
/// scope or the ignorable set changes. Most names in a part share a handful
/// of prefixes, so this spares nearly every walk over the declarations in
/// scope.
struct PrefixCache {
    keys: [u64; 16],
    classes: [Class; 16],
}

/// No prefix packs to this key; long prefixes are not cached.
const UNCACHED: u64 = u64::MAX;

impl PrefixCache {
    const fn new(none: Class) -> Self {
        Self {
            keys: [UNCACHED; 16],
            classes: [none; 16],
        }
    }

    /// A prefix of at most seven bytes as one integer, its length in the top
    /// byte, or [`UNCACHED`] for a longer one.
    fn key(prefix: &[u8]) -> u64 {
        if prefix.len() > 7 {
            return UNCACHED;
        }
        let mut key = (prefix.len() as u64) << 56;
        for (index, byte) in prefix.iter().enumerate() {
            key |= u64::from(*byte) << (8 * index);
        }
        key
    }

    fn slot(key: u64) -> usize {
        ((key ^ (key >> 17) ^ (key >> 41)) & 15) as usize
    }

    fn get(&self, key: u64) -> Option<Class> {
        let slot = Self::slot(key);
        (key != UNCACHED && self.keys[slot] == key).then_some(self.classes[slot])
    }

    fn put(&mut self, key: u64, class: Class) {
        if key != UNCACHED {
            let slot = Self::slot(key);
            self.keys[slot] = key;
            self.classes[slot] = class;
        }
    }

    fn clear(&mut self) {
        self.keys = [UNCACHED; 16];
    }
}

/// One attribute as written: its name and its raw value.
type Raw<'t> = (&'t [u8], &'t [u8]);

/// The state of one proof.
struct Prover<'a> {
    xml: &'a [u8],
    capabilities: &'a Capabilities,
    limits: &'a Limits,
    scope: Scope<String>,
    frames: Vec<Frame>,
    cache: PrefixCache,
    /// The class of no namespace: unprefixed attributes, and unprefixed
    /// elements outside any default namespace.
    none: Class,
    /// The namespaces the document element makes ignorable that the profile
    /// does not understand: the only ones processing drops.
    dropping: Vec<UriId>,
    /// Whether the document element carries a valid `Ignorable`.
    ignorable_directive: bool,
    /// Whether processing drops anything besides that attribute.
    drops: bool,
    /// Whether the document element has started.
    root_seen: bool,
    /// The declarations of the start tag being examined.
    local: Vec<(String, String)>,
    /// The element whose subtree may hold opaque wrappers.
    holder: Option<Holder<'a>>,
    /// The depth of the opaque wrapper being skipped.
    opaque: Option<usize>,
    /// The opaque wrappers seen.
    wrappers: usize,
}

impl<'a> Prover<'a> {
    fn new(
        xml: &'a [u8],
        capabilities: &'a Capabilities,
        limits: &'a Limits,
        holder: Option<Holder<'a>>,
    ) -> Self {
        let scope = Scope::new(capabilities);
        let none = Class {
            id: NO_NAMESPACE,
            mce: scope.is_mce(NO_NAMESPACE),
            understood: scope.understood(NO_NAMESPACE),
            xml: scope.is_xml(NO_NAMESPACE),
            dropping: false,
        };
        Self {
            xml,
            capabilities,
            limits,
            scope,
            frames: Vec::new(),
            cache: PrefixCache::new(none),
            none,
            dropping: Vec::new(),
            ignorable_directive: false,
            drops: false,
            root_seen: false,
            local: Vec::new(),
            holder,
            opaque: None,
            wrappers: 0,
        }
    }

    fn run<E>(
        mut self,
        mut observe: impl FnMut(usize) -> Result<(), E>,
    ) -> Result<Result<Held, Refusal>, E> {
        let mut reader = Reader::from_reader(self.xml);
        reader.config_mut().trim_text(false);
        loop {
            let before = reader.buffer_position();
            let event = reader.read_event();
            let consumed = usize::try_from(reader.buffer_position().saturating_sub(before))
                .unwrap_or(usize::MAX);
            observe(consumed.max(1))?;
            let decoder = reader.decoder();
            let step = match event {
                Ok(Event::Start(element)) => self.start(&element, decoder, false),
                Ok(Event::Empty(element)) => self.start(&element, decoder, true),
                Ok(Event::End(_)) => self.end(),
                Ok(Event::GeneralRef(reference)) => {
                    if self.frames.last().is_some_and(|frame| !frame.dropped) {
                        // The processor copies a visible reference only when
                        // it is a character reference or a predefined entity.
                        let known = match reference.resolve_char_ref() {
                            Ok(Some(_)) => reference.decode().is_ok(),
                            Ok(None) => reference.decode().is_ok_and(|name| {
                                matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot")
                            }),
                            Err(_) => false,
                        };
                        if known {
                            Ok(())
                        } else {
                            Err(Refusal::Unproven)
                        }
                    } else {
                        Ok(())
                    }
                },
                Ok(Event::Decl(_)) => {
                    if self.frames.is_empty() && !self.root_seen {
                        Ok(())
                    } else {
                        Err(Refusal::Unproven)
                    }
                },
                Ok(Event::DocType(_) | Event::PI(_)) | Err(_) => Err(Refusal::Unproven),
                Ok(Event::Eof) => {
                    return Ok(if self.frames.is_empty() {
                        Ok(Held {
                            outside: self.verdict(),
                            wrappers: self.wrappers,
                        })
                    } else {
                        Err(Refusal::Unproven)
                    });
                },
                Ok(Event::Text(_) | Event::CData(_) | Event::Comment(_)) => Ok(()),
            };
            if let Err(refusal) = step {
                return Ok(Err(refusal));
            }
        }
    }

    fn verdict(&self) -> Admission {
        if self.drops {
            Admission::IgnorableOpaque
        } else if self.ignorable_directive {
            Admission::Ignorable
        } else {
            Admission::DeclaredOnly
        }
    }

    fn end(&mut self) -> Result<(), Refusal> {
        let frame = self.frames.pop().ok_or(Refusal::Unproven)?;
        if self.opaque.is_some_and(|depth| self.frames.len() < depth) {
            self.opaque = None;
        }
        if frame.declares {
            self.scope.truncate(self.frames.len() + 1);
            self.cache.clear();
        }
        Ok(())
    }

    fn start(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        empty: bool,
    ) -> Result<(), Refusal> {
        if self.frames.len() >= self.limits.max_depth {
            return Err(Refusal::Unproven);
        }
        if self.opaque.is_some() {
            // Inside an opaque wrapper: only nesting is tracked.
            if !empty {
                self.frames.try_reserve(1).map_err(|_| Refusal::Unproven)?;
                self.frames.push(Frame {
                    bindings: 0,
                    declares: false,
                    dropped: false,
                    held: true,
                });
            }
            return Ok(());
        }
        let depth = self.frames.len() + 1;
        let root = depth == 1;
        if root {
            if self.root_seen {
                return Err(Refusal::Unproven);
            }
            self.root_seen = true;
        }
        let parent = self.frames.last().copied();
        let mut bindings = parent.map_or(1, |frame| frame.bindings);
        let parent_dropped = parent.is_some_and(|frame| frame.dropped);

        // Declarations apply to the whole tag wherever they stand in it, so a
        // tag that may hold one is read in full before any name on it is
        // resolved. Every declaration's name holds `xmlns`.
        let declares = if root || contains_xmlns(element.attributes_raw()) {
            let attributes = self.read_attributes(element)?;
            let declares = self.declare(&attributes, decoder, depth, &mut bindings)?;
            if root {
                self.root_directive(&attributes, decoder)?;
            }
            declares
        } else {
            false
        };

        let (prefix, local) = split_name(element.name().into_inner())?;
        let class = self.resolve(prefix, true)?;
        let parent_held = parent.is_some_and(|frame| frame.held);
        if class.mce {
            if parent_held && !parent_dropped && local == b"AlternateContent" {
                return self.open_wrapper(depth, bindings, declares, empty);
            }
            return Err(Refusal::Element);
        }
        let held = parent_held || self.is_holder(prefix, local);
        let dropped = if parent_dropped {
            if class.understood {
                return Err(Refusal::HiddenContent);
            }
            true
        } else if class.dropping {
            if root {
                return Err(Refusal::HiddenContent);
            }
            self.drops = true;
            true
        } else {
            false
        };
        self.attributes(element, decoder, root, dropped)?;

        if empty {
            if declares {
                self.scope.truncate(depth);
                self.cache.clear();
            }
        } else {
            self.frames.try_reserve(1).map_err(|_| Refusal::Unproven)?;
            self.frames.push(Frame {
                bindings,
                declares,
                dropped,
                held,
            });
        }
        Ok(())
    }

    /// Whether an element (not in the markup-compatibility namespace) is the
    /// holder.
    fn is_holder(&self, prefix: &[u8], local: &[u8]) -> bool {
        let Some(holder) = self.holder else {
            return false;
        };
        holder.local.as_bytes() == local
            && str::from_utf8(prefix)
                .ok()
                .and_then(|prefix| self.scope.resolve(prefix))
                .is_some_and(|uri| holder.namespaces.contains(&uri.text))
    }

    /// Enter an opaque `mc:AlternateContent` inside a holder, whose own
    /// declarations are already in scope.
    fn open_wrapper(
        &mut self,
        depth: usize,
        bindings: usize,
        declares: bool,
        empty: bool,
    ) -> Result<(), Refusal> {
        self.wrappers = self.wrappers.checked_add(1).ok_or(Refusal::Unproven)?;
        if empty {
            if declares {
                self.scope.truncate(depth);
                self.cache.clear();
            }
            return Ok(());
        }
        self.frames.try_reserve(1).map_err(|_| Refusal::Unproven)?;
        self.frames.push(Frame {
            bindings,
            declares,
            dropped: false,
            held: true,
        });
        self.opaque = Some(depth);
        Ok(())
    }

    /// Every attribute of a tag as written, refusing a malformed one and a
    /// tag over the attribute limit.
    fn read_attributes<'t>(&self, element: &'t BytesStart<'_>) -> Result<Vec<Raw<'t>>, Refusal> {
        let maximum = self
            .limits
            .max_attributes_per_element
            .min(ATTRIBUTES_PER_ELEMENT_CEILING);
        let mut attributes = Vec::new();
        for attribute in element.unchecked_attributes() {
            if attributes.len() == maximum {
                return Err(Refusal::Unproven);
            }
            let attribute = attribute.map_err(|_| Refusal::Unproven)?;
            let Cow::Borrowed(value) = attribute.value else {
                return Err(Refusal::Unproven);
            };
            attributes.try_reserve(1).map_err(|_| Refusal::Unproven)?;
            attributes.push((attribute.key.into_inner(), value));
        }
        Ok(attributes)
    }

    /// Validate and add one tag's declarations, as the processor does before
    /// it resolves any name on the tag. Returns whether there were any.
    fn declare(
        &mut self,
        attributes: &[Raw<'_>],
        decoder: Decoder,
        depth: usize,
        bindings: &mut usize,
    ) -> Result<bool, Refusal> {
        self.local.clear();
        for &(key, value) in attributes {
            let prefix = if key == b"xmlns" {
                ""
            } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                str::from_utf8(prefix).map_err(|_| Refusal::Unproven)?
            } else {
                continue;
            };
            let value = decode(key, value, decoder)?;
            if !prefix.is_empty() && (!xml_name::is_ncname(prefix) || value.is_empty()) {
                return Err(Refusal::Unproven);
            }
            // The processor resolves `xml` by definition and never resolves
            // `xmlns`; a declaration that rebinds either is outside the
            // proof, and quick-xml readers refuse it.
            if prefix == "xmlns" || (prefix == "xml" && value != XML_NS) {
                return Err(Refusal::Unproven);
            }
            self.local.try_reserve(1).map_err(|_| Refusal::Unproven)?;
            self.local.push((prefix.to_owned(), value.into_owned()));
        }
        if self.local.is_empty() {
            return Ok(false);
        }
        // Two declarations of one prefix, the only duplicate names a tag can
        // hold among its declarations.
        if has_duplicate_prefix(&self.local).map_err(|_| Refusal::Unproven)? {
            return Err(Refusal::Unproven);
        }
        for (prefix, _) in &self.local {
            if self.scope.get_outside(prefix, depth).is_none() {
                *bindings = bindings.checked_add(1).ok_or(Refusal::Unproven)?;
            }
        }
        if *bindings > self.limits.max_namespace_bindings {
            return Err(Refusal::Unproven);
        }
        self.scope
            .push(depth, &self.local, self.capabilities)
            .map_err(|_| Refusal::Unproven)?;
        self.cache.clear();
        Ok(true)
    }

    /// Find and validate the document element's `Ignorable`, before any name
    /// on it is classified: the directive applies to the element itself.
    fn root_directive(&mut self, attributes: &[Raw<'_>], decoder: Decoder) -> Result<(), Refusal> {
        let mut directive = None;
        for &(key, value) in attributes {
            let Some((prefix, local)) = split_prefixed(key) else {
                continue;
            };
            if prefix == b"xmlns" || !self.resolve(prefix, false)?.mce {
                continue;
            }
            if local != b"Ignorable" {
                return Err(directive_refusal(local));
            }
            if directive.is_some() {
                // Two prefixes for the namespace: a duplicate attribute.
                return Err(Refusal::Unproven);
            }
            directive = Some(decode(key, value, decoder)?.into_owned());
        }
        let Some(value) = directive else {
            return Ok(());
        };
        let mut seen: Vec<&str> = Vec::new();
        for prefix in value.split_whitespace() {
            if seen.len() >= self.limits.max_directive_tokens
                || !xml_name::is_ncname(prefix)
                || seen.contains(&prefix)
            {
                return Err(Refusal::Unproven);
            }
            seen.try_reserve(1).map_err(|_| Refusal::Unproven)?;
            seen.push(prefix);
            let id = self
                .scope
                .resolve(prefix)
                .map(|uri| uri.id)
                .ok_or(Refusal::Unproven)?;
            if self.scope.is_mce(id) {
                return Err(Refusal::Unproven);
            }
            if !self.scope.understood(id) && !self.dropping.contains(&id) {
                self.dropping
                    .try_reserve(1)
                    .map_err(|_| Refusal::Unproven)?;
                self.dropping.push(id);
            }
        }
        self.ignorable_directive = true;
        // Classes cached so far predate the ignorable set.
        self.cache.clear();
        Ok(())
    }

    /// Validate an `Ignorable` below the document element, as the processor
    /// does, and admit it only when every namespace it lists is understood
    /// by the profile: processing never drops an understood name, so the
    /// directive changes nothing but its own presence, wherever content
    /// moves. Litchi's own revision writer puts one listing `w16du` on each
    /// revision wrapper it authors.
    fn nested_directive(
        &mut self,
        key: &[u8],
        value: &[u8],
        decoder: Decoder,
    ) -> Result<(), Refusal> {
        let value = decode(key, value, decoder)?;
        let mut seen: Vec<&str> = Vec::new();
        for prefix in value.split_whitespace() {
            if seen.len() >= self.limits.max_directive_tokens
                || !xml_name::is_ncname(prefix)
                || seen.contains(&prefix)
            {
                return Err(Refusal::Unproven);
            }
            seen.try_reserve(1).map_err(|_| Refusal::Unproven)?;
            seen.push(prefix);
            let id = self
                .scope
                .resolve(prefix)
                .map(|uri| uri.id)
                .ok_or(Refusal::Unproven)?;
            if self.scope.is_mce(id) {
                return Err(Refusal::Unproven);
            }
            if !self.scope.understood(id) {
                return Err(Refusal::NestedIgnorable);
            }
        }
        self.ignorable_directive = true;
        Ok(())
    }

    /// Check every attribute of a tag whose declarations are in scope.
    fn attributes(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        root: bool,
        dropped: bool,
    ) -> Result<(), Refusal> {
        let maximum = self
            .limits
            .max_attributes_per_element
            .min(ATTRIBUTES_PER_ELEMENT_CEILING);
        let mut count = 0usize;
        // Hashes of the expanded names seen so far: an equal hash may be a
        // repeated name, which the exact check below then decides, as it
        // does for every tag wider than this.
        let mut hashes = [0u64; NAME_HASHES];
        let mut named = 0usize;
        let mut suspect = false;
        for attribute in element.unchecked_attributes() {
            count += 1;
            if count > maximum {
                return Err(Refusal::Unproven);
            }
            let attribute = attribute.map_err(|_| Refusal::Unproven)?;
            let key = attribute.key.into_inner();
            let value = attribute.value.as_ref();
            // Only a reference can make a UTF-8 value undecodable;
            // normalization cannot fail.
            match value
                .iter()
                .position(|byte| *byte == b'&' || !byte.is_ascii())
            {
                None => {},
                Some(_) if value.contains(&b'&') => {
                    decode(key, value, decoder)?;
                },
                Some(_) => {
                    str::from_utf8(value).map_err(|_| Refusal::Unproven)?;
                },
            }
            if is_declaration(key) {
                continue;
            }
            let (prefix, local) = split_name(key)?;
            let class = self.resolve(prefix, false)?;
            if class.mce {
                if local != b"Ignorable" {
                    return Err(directive_refusal(local));
                }
                if !root {
                    self.nested_directive(key, value, decoder)?;
                }
            } else if dropped {
                if class.understood && !class.xml {
                    return Err(Refusal::HiddenContent);
                }
            } else if class.dropping {
                self.drops = true;
            }
            let hash = name_hash(local, class.id);
            if named < NAME_HASHES {
                suspect |= hashes[..named].contains(&hash);
                hashes[named] = hash;
                named += 1;
            } else {
                suspect = true;
            }
        }
        if suspect {
            self.exact_duplicates(element)?;
        }
        Ok(())
    }

    /// Refuse a tag that names one ordinary attribute twice, under one prefix
    /// or under two bound to one namespace: the processor's lexical and
    /// expanded duplicate checks together. Runs only when the filter in
    /// [`Self::attributes`] saw a possible repeat; sorting keeps it
    /// `O(n log n)` on a tag of any width.
    fn exact_duplicates(&mut self, element: &BytesStart<'_>) -> Result<(), Refusal> {
        let mut names: Vec<(&[u8], UriId)> = Vec::new();
        for attribute in element.unchecked_attributes() {
            let key = attribute.map_err(|_| Refusal::Unproven)?.key.into_inner();
            if is_declaration(key) {
                continue;
            }
            let (prefix, local) = split_name(key)?;
            let id = self.resolve(prefix, false)?.id;
            names.try_reserve(1).map_err(|_| Refusal::Unproven)?;
            names.push((local, id));
        }
        names.sort_unstable();
        if names.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(Refusal::Unproven);
        }
        Ok(())
    }

    /// The class of the namespace `prefix` is bound to: an element's
    /// unprefixed name takes the default namespace, an attribute's none.
    fn resolve(&mut self, prefix: &[u8], element: bool) -> Result<Class, Refusal> {
        if prefix.is_empty() && !element {
            return Ok(self.none);
        }
        let key = PrefixCache::key(prefix);
        if let Some(class) = self.cache.get(key) {
            return Ok(class);
        }
        let text = str::from_utf8(prefix).map_err(|_| Refusal::Unproven)?;
        let id = match self.scope.resolve(text) {
            Some(uri) => uri.id,
            None if prefix.is_empty() => NO_NAMESPACE,
            None => return Err(Refusal::Unproven),
        };
        let class = Class {
            id,
            mce: self.scope.is_mce(id),
            understood: self.scope.understood(id),
            xml: self.scope.is_xml(id),
            dropping: !self.scope.understood(id) && self.dropping.contains(&id),
        };
        self.cache.put(key, class);
        Ok(class)
    }
}

/// Whether an attribute name declares a namespace, as the processor reads it.
fn is_declaration(key: &[u8]) -> bool {
    key == b"xmlns" || key.starts_with(b"xmlns:")
}

/// Whether a tag's attribute bytes hold `xmlns`, as every declaration does.
fn contains_xmlns(raw: &[u8]) -> bool {
    memchr::memchr_iter(b'x', raw).any(|at| raw[at..].starts_with(b"xmlns"))
}

/// The ordinary attributes one tag compares by hash before the duplicate
/// check falls back to sorting their names.
const NAME_HASHES: usize = 16;

/// A hash of an expanded attribute name, for the duplicate check.
fn name_hash(local: &[u8], id: UriId) -> u64 {
    let mut hash = (id as u64).wrapping_mul(0x9E37_79B9_7F4A_7C15) ^ 0xCBF2_9CE4_8422_2325;
    for byte in local {
        hash = (hash ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01B3);
    }
    hash
}

/// Decode and normalize one attribute value exactly as the processor does.
fn decode<'t>(key: &'t [u8], value: &'t [u8], decoder: Decoder) -> Result<Cow<'t, str>, Refusal> {
    Attribute {
        key: QName(key),
        value: Cow::Borrowed(value),
    }
    .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
    .map_err(|_| Refusal::Unproven)
}

/// The refusal for a markup-compatibility attribute other than `Ignorable`:
/// a known directive, or a name the processor rejects.
fn directive_refusal(local: &[u8]) -> Refusal {
    if matches!(
        local,
        b"ProcessContent" | b"MustUnderstand" | b"PreserveElements" | b"PreserveAttributes"
    ) {
        Refusal::Directive
    } else {
        Refusal::Unproven
    }
}

/// Split a qualified name at its colon, refusing anything that is not one
/// by the processor's rule: an `NCName`, optionally a colon and another.
fn split_name(name: &[u8]) -> Result<(&[u8], &[u8]), Refusal> {
    let mut colon = None;
    let mut at_start = true;
    for (index, &byte) in name.iter().enumerate() {
        match byte {
            b'A'..=b'Z' | b'a'..=b'z' | b'_' => {},
            b'0'..=b'9' | b'-' | b'.' if !at_start => {},
            b':' if colon.is_none() && !at_start => {
                colon = Some(index);
                at_start = true;
                continue;
            },
            byte if !byte.is_ascii() => return split_unicode_name(name),
            _ => return Err(Refusal::Unproven),
        }
        at_start = false;
    }
    if at_start {
        return Err(Refusal::Unproven);
    }
    Ok(match colon {
        Some(colon) => (&name[..colon], &name[colon + 1..]),
        None => (&[], name),
    })
}

#[cold]
fn split_unicode_name(name: &[u8]) -> Result<(&[u8], &[u8]), Refusal> {
    let text = str::from_utf8(name).map_err(|_| Refusal::Unproven)?;
    if !xml_name::is_qualified_name(text) {
        return Err(Refusal::Unproven);
    }
    Ok(split_prefixed(name).unwrap_or((&[], name)))
}

/// The prefix and local part of a name with a colon.
fn split_prefixed(name: &[u8]) -> Option<(&[u8], &[u8])> {
    let colon = memchr::memchr(b':', name)?;
    Some((&name[..colon], &name[colon + 1..]))
}

#[cfg(test)]
#[path = "admission_tests.rs"]
mod tests;
