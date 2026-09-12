//! Bounded, source-borrowing namespace resolution for the OTH structure codec.
//!
//! `quick_xml::reader::NsReader` keeps a second namespace stack internally and
//! stores declaration values exactly as they appeared in the input.  OTH needs
//! both a normalized namespace value (character references are legal in an
//! `xmlns` value) and one budget that covers the semantic projection.  This
//! module is the small resolver used by a plain `quick_xml::Reader` pass.
//!
//! The resolver borrows element and local names from the reader's source and
//! retains only bounded owned prefix index keys.  A URI is represented by a
//! small known-namespace enum when it is one of the OTH namespaces.  Other
//! URIs are interned once and shared through an `Arc<str>`, so aliases do not
//! copy a URI for every attribute or element.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    encoding::Decoder,
    events::BytesStart,
    name::{LocalName, QName},
};
use std::{
    cmp::Ordering,
    collections::{HashMap, hash_map::DefaultHasher},
    hash::BuildHasherDefault,
    marker::PhantomData,
    sync::Arc,
};

use std::mem::size_of;

const ARC_HEADER_BYTES: usize = size_of::<usize>() * 2;
const HASH_BUCKET_BYTES: usize = size_of::<usize>();

/// The largest context-local accounting total accepted by this helper.
///
/// The structure projection's existing aggregate limit is 16 MiB.  Keeping a
/// matching local ceiling means a caller that uses the helper without a
/// parent budget still cannot grow the namespace index without a bound.
pub(crate) const MAX_NAMESPACE_BYTES: usize = 16 * 1024 * 1024;

/// The maximum number of open namespace scopes retained by one reader pass.
pub(crate) const MAX_NAMESPACE_DEPTH: usize = 256;

/// The maximum number of namespace declarations on one start tag.
pub(crate) const MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT: usize = 256;

const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const DR3D_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const SVG_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const DC_NAMESPACE: &str = "http://purl.org/dc/elements/1.1/";
const META_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const XHTML_NAMESPACE: &str = "http://www.w3.org/1999/xhtml";
const FO_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const STYLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";

type PrefixMap = HashMap<Arc<[u8]>, Binding, BuildHasherDefault<DefaultHasher>>;
type UnknownUriIndex = HashMap<Arc<str>, (), BuildHasherDefault<DefaultHasher>>;

/// Known OTH namespace identities.  Keeping these as enum values avoids an
/// owned URI for every known element and attribute.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[non_exhaustive]
pub(crate) enum Ns {
    Office,
    Text,
    Table,
    Draw,
    Dr3d,
    Xlink,
    Svg,
    Xml,
    Dc,
    Meta,
    Xhtml,
    Fo,
    Style,
}

/// A normalized namespace identity.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) enum NamespaceId {
    Known(Ns),
    Unknown(Arc<str>),
}

impl NamespaceId {
    pub(crate) fn known(&self) -> Option<Ns> {
        match self {
            Self::Known(namespace) => Some(*namespace),
            Self::Unknown(_) => None,
        }
    }

    pub(crate) fn unknown_uri(&self) -> Option<&str> {
        match self {
            Self::Known(_) => None,
            Self::Unknown(uri) => Some(uri),
        }
    }

    pub(crate) fn is_unknown(&self) -> bool {
        matches!(self, Self::Unknown(_))
    }
}

/// The result of resolving one source QName.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct ResolvedQName<'a> {
    pub(crate) namespace: Option<NamespaceId>,
    pub(crate) local: LocalName<'a>,
    pub(crate) prefix: Option<&'a [u8]>,
}

impl<'a> ResolvedQName<'a> {
    pub(crate) fn local_bytes(&self) -> &'a [u8] {
        self.local.into_inner()
    }
}

/// `xmlns` declaration spelling extracted without allocating.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum NamespaceDeclaration<'a> {
    Default,
    Prefix(&'a [u8]),
}

impl<'a> NamespaceDeclaration<'a> {
    pub(crate) fn prefix(self) -> &'a [u8] {
        match self {
            Self::Default => b"",
            Self::Prefix(prefix) => prefix,
        }
    }
}

/// Extracts a namespace declaration from an attribute key.
pub(crate) fn namespace_declaration<'a>(key: QName<'a>) -> Option<NamespaceDeclaration<'a>> {
    let key = key.into_inner();
    if key == b"xmlns" {
        return Some(NamespaceDeclaration::Default);
    }
    key.strip_prefix(b"xmlns:")
        .map(NamespaceDeclaration::Prefix)
}

#[derive(Clone, Debug)]
struct Binding {
    /// `None` means this prefix is explicitly unbound (`xmlns=""` for the
    /// default prefix).  The distinction from a missing map entry matters for
    /// prefixed names, which must still fail as unknown.
    namespace: Option<NamespaceId>,
}

#[derive(Clone, Debug)]
struct Undo {
    prefix: Arc<[u8]>,
    previous: Option<Binding>,
    changed: bool,
}

#[derive(Clone, Copy, Debug)]
struct Frame {
    undo_start: usize,
}

/// A scoped semantic namespace resolver for one borrowed XML input.
pub(crate) struct SemanticNamespaceContext<'xml> {
    prefixes: PrefixMap,
    undo: Vec<Undo>,
    frames: Vec<Frame>,
    unknown_uris: UnknownUriIndex,
    charged_bytes: usize,
    max_depth: usize,
    max_declarations_per_element: usize,
    poisoned: bool,
    marker: PhantomData<&'xml ()>,
}

impl<'xml> SemanticNamespaceContext<'xml> {
    /// Creates a context with the OTH reader bounds.
    pub(crate) fn new() -> Self {
        Self::with_limits(MAX_NAMESPACE_DEPTH, MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT)
    }

    /// Creates a context with explicit depth and per-tag declaration bounds.
    /// The aggregate byte bound remains [`MAX_NAMESPACE_BYTES`].
    pub(crate) fn with_limits(max_depth: usize, max_declarations_per_element: usize) -> Self {
        Self {
            prefixes: PrefixMap::default(),
            undo: Vec::new(),
            frames: Vec::new(),
            unknown_uris: UnknownUriIndex::default(),
            charged_bytes: 0,
            max_depth: max_depth.min(MAX_NAMESPACE_DEPTH),
            max_declarations_per_element: max_declarations_per_element
                .min(MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT),
            poisoned: false,
            marker: PhantomData,
        }
    }

    pub(crate) fn depth(&self) -> usize {
        self.frames.len()
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.frames.is_empty()
    }

    fn ensure_live(&self) -> Result<()> {
        if self.poisoned {
            return invalid("OTH namespace context is unusable after a failed scope");
        }
        Ok(())
    }

    /// Begins a start/empty-element scope and validates its names and
    /// expanded attributes.  Namespace declarations are normalized and
    /// inserted before the element or ordinary attributes are resolved.
    ///
    /// The callback is called before every context allocation and before a
    /// declaration value is decoded.  A parent projection can pass its
    /// temporary `Budget::storage` reservation here.
    pub(crate) fn start<F>(
        &mut self,
        start: &BytesStart<'_>,
        decoder: Decoder,
        charge: &mut F,
    ) -> Result<()>
    where
        F: FnMut(usize) -> Result<()>,
    {
        self.begin_scope(start, decoder, charge)
    }

    /// Begins a start/empty-element scope.  This is the implementation seam
    /// kept separate from `start` so callers can name the operation `push`.
    pub(crate) fn push<F>(
        &mut self,
        start: &BytesStart<'_>,
        decoder: Decoder,
        charge: &mut F,
    ) -> Result<()>
    where
        F: FnMut(usize) -> Result<()>,
    {
        self.begin_scope(start, decoder, charge)
    }

    fn begin_scope<F>(
        &mut self,
        start: &BytesStart<'_>,
        decoder: Decoder,
        charge: &mut F,
    ) -> Result<()>
    where
        F: FnMut(usize) -> Result<()>,
    {
        self.ensure_live()?;
        let result = self.push_inner(start, decoder, charge);
        if result.is_err() {
            // Declarations are rolled back by `push_inner`; marking the
            // context terminal prevents callers from accidentally continuing
            // with a parent budget that may already have consumed a partial
            // reservation after a failed scope.
            self.poisoned = true;
        }
        result
    }

    fn push_inner<F>(
        &mut self,
        start: &BytesStart<'_>,
        decoder: Decoder,
        charge: &mut F,
    ) -> Result<()>
    where
        F: FnMut(usize) -> Result<()>,
    {
        if self.frames.len() >= self.max_depth {
            return invalid(format!("OTH namespace depth exceeds {}", self.max_depth));
        }

        validate_qname(start.name(), "element")?;

        let mut declarations = 0usize;
        let mut ordinary_attributes = 0usize;
        let mut prefix_bytes = 0usize;

        // This pass only inspects borrowed names and raw values.  It lets us
        // charge and reserve the frame, undo records, expanded-name slots,
        // and map buckets before any corresponding allocation.
        let mut attributes = start.attributes();
        attributes.with_checks(false);
        for raw in attributes {
            let attribute = raw.map_err(attribute_error)?;
            validate_qname(attribute.key, "attribute")?;
            if let Some(declaration) = namespace_declaration(attribute.key) {
                declarations = declarations
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("OTH namespace declaration count overflow"))?;
                if declarations > self.max_declarations_per_element {
                    return invalid(format!(
                        "OTH element declares more than {} namespaces",
                        self.max_declarations_per_element
                    ));
                }
                let prefix = declaration.prefix();
                prefix_bytes = prefix_bytes
                    .checked_add(prefix.len())
                    .ok_or_else(|| invalid_error("OTH namespace prefix storage overflow"))?;
            } else {
                ordinary_attributes = ordinary_attributes
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("OTH attribute count overflow"))?;
            }
        }

        // Charge the permanent frame/index state before reserving it.
        // `size_of::<&[u8]>()` deliberately
        // accounts for the borrowed slice slot even though the source bytes
        // themselves stay in the reader input.
        let mut context_bytes = size_of::<Frame>();
        context_bytes = context_bytes
            .checked_add(
                size_of::<Undo>()
                    .checked_mul(declarations)
                    .ok_or_else(|| invalid_error("OTH namespace undo storage overflow"))?,
            )
            .ok_or_else(|| invalid_error("OTH namespace context storage overflow"))?;
        context_bytes = context_bytes
            .checked_add(
                (size_of::<(Arc<[u8]>, Binding)>() + HASH_BUCKET_BYTES)
                    .checked_mul(declarations)
                    .ok_or_else(|| invalid_error("OTH namespace index slot overflow"))?,
            )
            .ok_or_else(|| invalid_error("OTH namespace context storage overflow"))?;
        context_bytes = context_bytes
            .checked_add(prefix_bytes)
            .ok_or_else(|| invalid_error("OTH namespace context storage overflow"))?;
        context_bytes = context_bytes
            .checked_add(
                ARC_HEADER_BYTES
                    .checked_mul(declarations)
                    .ok_or_else(|| invalid_error("OTH namespace prefix Arc storage overflow"))?,
            )
            .ok_or_else(|| invalid_error("OTH namespace context storage overflow"))?;
        self.charge(context_bytes, charge)?;

        // Check declaration-prefix uniqueness with one bounded pass.  The
        // borrowed slice slots are charged before this temporary vector is
        // reserved; at most the per-tag declaration bound can be retained.
        let prefix_slots = size_of::<&[u8]>()
            .checked_mul(declarations)
            .ok_or_else(|| invalid_error("OTH namespace prefix index overflow"))?;
        self.charge(prefix_slots, charge)?;
        let mut declaration_prefixes = Vec::<&[u8]>::new();
        declaration_prefixes
            .try_reserve_exact(declarations)
            .map_err(|source| allocation("OTH namespace declaration prefixes", source))?;
        let mut declaration_attributes = start.attributes();
        declaration_attributes.with_checks(false);
        for raw in declaration_attributes {
            let attribute = raw.map_err(attribute_error)?;
            let Some(declaration) = namespace_declaration(attribute.key) else {
                continue;
            };
            let prefix = declaration.prefix();
            if declaration_prefixes.contains(&prefix) {
                return invalid("duplicate OTH namespace declaration");
            }
            declaration_prefixes.push(prefix);
        }
        drop(declaration_prefixes);

        // Each reservation follows its charge.  This keeps allocator failure
        // deterministic and leaves a caller budget as the first guard.
        self.frames
            .try_reserve_exact(1)
            .map_err(|source| allocation("OTH namespace frames", source))?;
        self.undo
            .try_reserve_exact(declarations)
            .map_err(|source| allocation("OTH namespace undo log", source))?;
        self.prefixes
            .try_reserve(declarations)
            .map_err(|source| allocation("OTH namespace prefix index", source))?;

        let undo_start = self.undo.len();
        let frame_count = self.frames.len();
        let result = self.apply_declarations(start, decoder, declarations, charge);
        if let Err(error) = result {
            self.rollback_to(undo_start);
            debug_assert_eq!(self.frames.len(), frame_count);
            return Err(error);
        }

        // Declarations are now in scope, so an alias such as `a:x` and `b:x`
        // can be checked against one normalized expanded namespace identity.
        let result = self.validate_attributes(start, ordinary_attributes, charge);
        if let Err(error) = result {
            self.rollback_to(undo_start);
            debug_assert_eq!(self.frames.len(), frame_count);
            return Err(error);
        }

        self.frames.push(Frame { undo_start });
        Ok(())
    }

    fn apply_declarations<F>(
        &mut self,
        start: &BytesStart<'_>,
        decoder: Decoder,
        declarations: usize,
        charge: &mut F,
    ) -> Result<()>
    where
        F: FnMut(usize) -> Result<()>,
    {
        let mut seen = 0usize;
        let mut attributes = start.attributes();
        attributes.with_checks(false);
        for raw in attributes {
            let attribute = raw.map_err(attribute_error)?;
            let key = attribute.key.into_inner();
            let Some(prefix) = (if key == b"xmlns" {
                Some(b"".as_slice())
            } else {
                key.strip_prefix(b"xmlns:")
            }) else {
                continue;
            };
            seen += 1;
            debug_assert!(seen <= declarations);
            let raw_value_len = attribute.value.len();

            // `decoded_and_normalized_value` uses a Cow and allocates when a
            // reference or an XML 1.0 whitespace character needs replacing.
            // The raw value is a safe upper bound for that temporary output;
            // charge it before invoking quick-xml's decoder.
            self.charge(raw_value_len, charge)?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "invalid OTH namespace declaration value: {error}"
                    ))
                })?;
            let value = value.as_ref();
            let namespace = self.namespace_id(value, prefix, charge)?;

            // The special `xml` binding is implicit.  Retaining an undo record
            // for it still makes duplicate-prefix detection and frame costs
            // uniform, while avoiding an index entry that could be shadowed.
            let owned_prefix: Arc<[u8]> = Arc::from(prefix);
            let special_xml = prefix == b"xml";
            let previous = if special_xml {
                None
            } else {
                self.prefixes.insert(
                    Arc::clone(&owned_prefix),
                    Binding {
                        namespace: namespace.clone(),
                    },
                )
            };
            self.undo.push(Undo {
                prefix: owned_prefix,
                previous,
                changed: !special_xml,
            });
        }
        debug_assert_eq!(seen, declarations);
        Ok(())
    }

    fn namespace_id<F>(
        &mut self,
        value: &str,
        prefix: &[u8],
        charge: &mut F,
    ) -> Result<Option<NamespaceId>>
    where
        F: FnMut(usize) -> Result<()>,
    {
        if value.is_empty() {
            if !prefix.is_empty() {
                return invalid("OTH prefixed namespace declarations cannot be empty");
            }
            return Ok(None);
        }

        if value == XMLNS_NAMESPACE {
            return invalid("OTH namespace declarations cannot bind the xmlns URI");
        }
        if value == XML_NAMESPACE {
            if prefix != b"xml" {
                return invalid("OTH XML namespace URI requires the xml prefix");
            }
            return Ok(Some(NamespaceId::Known(Ns::Xml)));
        }
        if prefix == b"xml" {
            return invalid("OTH xml prefix must bind the XML namespace URI");
        }
        if prefix == b"xmlns" {
            return invalid("OTH xmlns prefix is reserved");
        }

        if let Some(namespace) = known_namespace(value) {
            return Ok(Some(NamespaceId::Known(namespace)));
        }

        if let Some((uri, ())) = self.unknown_uris.get_key_value(value) {
            return Ok(Some(NamespaceId::Unknown(Arc::clone(uri))));
        }

        let bytes = ARC_HEADER_BYTES
            .checked_add(value.len())
            .and_then(|bytes| bytes.checked_add(size_of::<(Arc<str>, ())>() + HASH_BUCKET_BYTES))
            .ok_or_else(|| invalid_error("OTH namespace URI storage overflow"))?;
        self.charge(bytes, charge)?;
        self.unknown_uris
            .try_reserve(1)
            .map_err(|source| allocation("OTH unknown namespace URIs", source))?;
        let uri: Arc<str> = Arc::from(value);
        self.unknown_uris.insert(Arc::clone(&uri), ());
        Ok(Some(NamespaceId::Unknown(uri)))
    }

    fn validate_attributes<F>(
        &mut self,
        start: &BytesStart<'_>,
        ordinary_attributes: usize,
        charge: &mut F,
    ) -> Result<()>
    where
        F: FnMut(usize) -> Result<()>,
    {
        // This method is logically read-only, but the charge callback is
        // intentionally accepted so callers can account duplicate-key slots
        // before the temporary vector is allocated.
        let bytes = size_of::<ExpandedAttribute<'static>>()
            .checked_mul(ordinary_attributes)
            .ok_or_else(|| invalid_error("OTH expanded attribute storage overflow"))?;
        // The callback is the parent budget reservation and must see the
        // temporary vector allocation immediately before it is made.
        self.charge(bytes, charge)?;
        let mut seen = Vec::<ExpandedAttribute<'_>>::new();
        seen.try_reserve_exact(ordinary_attributes)
            .map_err(|source| allocation("OTH expanded attributes", source))?;

        let mut attributes = start.attributes();
        attributes.with_checks(false);
        for raw in attributes {
            let attribute = raw.map_err(attribute_error)?;
            if namespace_declaration(attribute.key).is_some() {
                continue;
            }
            let resolved = self.resolve_attribute(attribute.key)?;
            let key = ExpandedAttribute {
                namespace: resolved.namespace,
                local: resolved.local,
            };
            seen.push(key);
        }
        // Sorting keeps duplicate detection bounded by O(A log A) while
        // retaining borrowed local names and avoiding a second hash index.
        seen.sort_unstable_by(compare_expanded_attributes);
        if seen.windows(2).any(|pair| pair[0] == pair[1]) {
            return invalid("duplicate OTH expanded attribute");
        }
        Ok(())
    }

    /// Resolves an element QName using the current default namespace.
    pub(crate) fn resolve_element<'name>(
        &self,
        name: QName<'name>,
    ) -> Result<ResolvedQName<'name>> {
        self.resolve(name, true)
    }

    /// Resolves an attribute QName.  Unprefixed attributes intentionally do
    /// not use the current default namespace.
    pub(crate) fn resolve_attribute<'name>(
        &self,
        name: QName<'name>,
    ) -> Result<ResolvedQName<'name>> {
        self.resolve(name, false)
    }

    fn resolve<'name>(
        &self,
        name: QName<'name>,
        use_default: bool,
    ) -> Result<ResolvedQName<'name>> {
        self.ensure_live()?;
        validate_qname(name, if use_default { "element" } else { "attribute" })?;
        let (local, prefix) = name.decompose();
        let prefix_bytes = prefix.map(|prefix| prefix.into_inner());
        let namespace = match prefix_bytes {
            None if !use_default => None,
            None => self
                .prefixes
                .get(b"".as_slice())
                .and_then(|binding| binding.namespace.clone()),
            Some(b"xml") => Some(NamespaceId::Known(Ns::Xml)),
            Some(b"xmlns") => return invalid("OTH xmlns prefix is reserved for declarations"),
            Some(prefix) => Some(match self.prefixes.get(prefix) {
                Some(binding) => binding.namespace.clone().ok_or_else(|| {
                    invalid_error("OTH QName uses an explicitly unbound namespace prefix")
                })?,
                None => return invalid("OTH QName uses an unbound namespace prefix"),
            }),
        };
        Ok(ResolvedQName {
            namespace,
            local,
            prefix: prefix_bytes,
        })
    }

    /// Ends the current scope and restores all shadowed prefix bindings.
    pub(crate) fn end(&mut self) -> Result<()> {
        self.pop()
    }

    /// Alias for [`Self::end`].
    pub(crate) fn pop(&mut self) -> Result<()> {
        self.ensure_live()?;
        let frame = self
            .frames
            .pop()
            .ok_or_else(|| invalid_error("OTH namespace scope underflow"))?;
        self.rollback_to(frame.undo_start);
        Ok(())
    }

    fn rollback_to(&mut self, undo_start: usize) {
        while self.undo.len() > undo_start {
            let undo = self.undo.pop().expect("undo length checked");
            if !undo.changed {
                continue;
            }
            match undo.previous {
                Some(previous) => {
                    self.prefixes.insert(undo.prefix, previous);
                },
                None => {
                    self.prefixes.remove(&undo.prefix);
                },
            }
        }
    }

    fn charge<F>(&mut self, bytes: usize, charge: &mut F) -> Result<()>
    where
        F: FnMut(usize) -> Result<()>,
    {
        let next = self
            .charged_bytes
            .checked_add(bytes)
            .ok_or_else(|| invalid_error("OTH namespace budget overflow"))?;
        if next > MAX_NAMESPACE_BYTES {
            return invalid("OTH namespace context exceeds the aggregate limit");
        }
        charge(bytes)?;
        self.charged_bytes = next;
        Ok(())
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct ExpandedAttribute<'a> {
    namespace: Option<NamespaceId>,
    local: LocalName<'a>,
}

fn compare_expanded_attributes(
    left: &ExpandedAttribute<'_>,
    right: &ExpandedAttribute<'_>,
) -> Ordering {
    compare_namespaces(&left.namespace, &right.namespace)
        .then_with(|| left.local.into_inner().cmp(right.local.into_inner()))
}

fn compare_namespaces(left: &Option<NamespaceId>, right: &Option<NamespaceId>) -> Ordering {
    let (left_kind, left_namespace, left_uri) = namespace_sort_key(left);
    let (right_kind, right_namespace, right_uri) = namespace_sort_key(right);
    left_kind
        .cmp(&right_kind)
        .then_with(|| match (left_namespace, right_namespace) {
            (Some(left), Some(right)) => {
                known_namespace_rank(left).cmp(&known_namespace_rank(right))
            },
            (None, None) | (Some(_), None) | (None, Some(_)) => Ordering::Equal,
        })
        .then_with(|| match (left_uri, right_uri) {
            (None, None) => Ordering::Equal,
            (None, Some(_)) => Ordering::Less,
            (Some(_), None) => Ordering::Greater,
            (Some(left), Some(right)) => left.cmp(right),
        })
}

fn namespace_sort_key(namespace: &Option<NamespaceId>) -> (u8, Option<Ns>, Option<&str>) {
    match namespace {
        None => (0, None, None),
        Some(NamespaceId::Known(namespace)) => (1, Some(*namespace), None),
        Some(NamespaceId::Unknown(uri)) => (2, None, Some(uri.as_ref())),
    }
}

fn known_namespace_rank(namespace: Ns) -> u8 {
    match namespace {
        Ns::Office => 0,
        Ns::Text => 1,
        Ns::Table => 2,
        Ns::Draw => 3,
        Ns::Dr3d => 4,
        Ns::Xlink => 5,
        Ns::Svg => 6,
        Ns::Xml => 7,
        Ns::Dc => 8,
        Ns::Meta => 9,
        Ns::Xhtml => 10,
        Ns::Fo => 11,
        Ns::Style => 12,
    }
}

fn known_namespace(uri: &str) -> Option<Ns> {
    Some(match uri {
        OFFICE_NAMESPACE => Ns::Office,
        TEXT_NAMESPACE => Ns::Text,
        TABLE_NAMESPACE => Ns::Table,
        DRAW_NAMESPACE => Ns::Draw,
        DR3D_NAMESPACE => Ns::Dr3d,
        XLINK_NAMESPACE => Ns::Xlink,
        SVG_NAMESPACE => Ns::Svg,
        XML_NAMESPACE => Ns::Xml,
        DC_NAMESPACE => Ns::Dc,
        META_NAMESPACE => Ns::Meta,
        XHTML_NAMESPACE => Ns::Xhtml,
        FO_NAMESPACE => Ns::Fo,
        STYLE_NAMESPACE => Ns::Style,
        _ => return None,
    })
}

fn validate_qname(name: QName<'_>, field: &str) -> Result<()> {
    let raw = std::str::from_utf8(name.as_ref()).map_err(|error| {
        Error::InvalidFormat(format!("invalid OTH {field} QName UTF-8: {error}"))
    })?;
    let mut pieces = raw.split(':');
    let first = pieces.next().unwrap_or_default();
    let second = pieces.next();
    if first.is_empty() || !is_ncname(first) {
        return invalid(format!("invalid OTH {field} QName"));
    }
    match second {
        None => {
            if raw == "xmlns" && field == "element" {
                // An unprefixed element named `xmlns` is legal as an NCName;
                // only the `xmlns` prefix is reserved.
            }
        },
        Some(local) if !local.is_empty() && is_ncname(local) && pieces.next().is_none() => {},
        Some(_) => return invalid(format!("invalid OTH {field} QName")),
    }
    if raw.starts_with("xmlns:") && field != "attribute" {
        return invalid("OTH element names cannot use the xmlns prefix");
    }
    if raw.starts_with("xmlns:") && raw == "xmlns:" {
        return invalid("OTH namespace declaration has an empty prefix");
    }
    Ok(())
}

fn is_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    is_ncname_start(first) && chars.all(is_ncname_character)
}

fn is_ncname_start(character: char) -> bool {
    matches!(
        character,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{00c0}'..='\u{00d6}'
            | '\u{00d8}'..='\u{00f6}'
            | '\u{00f8}'..='\u{02ff}'
            | '\u{0370}'..='\u{037d}'
            | '\u{037f}'..='\u{1fff}'
            | '\u{200c}'..='\u{200d}'
            | '\u{2070}'..='\u{218f}'
            | '\u{2c00}'..='\u{2fef}'
            | '\u{3001}'..='\u{d7ff}'
            | '\u{f900}'..='\u{fdcf}'
            | '\u{fdf0}'..='\u{fffd}'
            | '\u{10000}'..='\u{effff}'
    )
}

fn is_ncname_character(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00b7}' | '\u{0300}'..='\u{036f}' | '\u{203f}'..='\u{2040}'
        )
}

fn attribute_error(error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!("invalid OTH XML attribute: {error}"))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::{events::Event, reader::Reader};
    use std::sync::Arc;

    fn parse_start<'a>(reader: &mut Reader<&'a [u8]>) -> BytesStart<'a> {
        match reader.read_event().expect("XML event") {
            Event::Start(start) | Event::Empty(start) => start,
            event => panic!("expected start event, got {event:?}"),
        }
    }

    fn no_charge(_: usize) -> Result<()> {
        Ok(())
    }

    #[test]
    fn shadows_and_restores_default_and_prefixed_bindings() {
        let xml = br#"<r xmlns="urn:outer" xmlns:p="urn:p1"><p:x><y xmlns="urn:inner" xmlns:p="urn:p2"/></p:x></r>"#;
        let mut reader = Reader::from_reader(xml.as_slice());
        let mut context = SemanticNamespaceContext::new();

        let root = parse_start(&mut reader);
        context
            .start(&root, reader.decoder(), &mut no_charge)
            .expect("root scope");
        assert_eq!(
            context.resolve_element(root.name()).unwrap().namespace,
            Some(NamespaceId::Unknown(Arc::from("urn:outer")))
        );
        assert_eq!(
            context
                .resolve_attribute(QName(b"plain"))
                .unwrap()
                .namespace,
            None
        );

        let child = parse_start(&mut reader);
        context
            .start(&child, reader.decoder(), &mut no_charge)
            .expect("child scope");
        assert_eq!(
            context.resolve_element(child.name()).unwrap().namespace,
            Some(NamespaceId::Unknown(Arc::from("urn:p1")))
        );
        context.pop().expect("child end");

        let empty = parse_start(&mut reader);
        context
            .start(&empty, reader.decoder(), &mut no_charge)
            .expect("empty scope");
        assert_eq!(
            context.resolve_element(QName(b"y")).unwrap().namespace,
            Some(NamespaceId::Unknown(Arc::from("urn:inner")))
        );
        assert_eq!(
            context.resolve_element(QName(b"p:x")).unwrap().namespace,
            Some(NamespaceId::Unknown(Arc::from("urn:p2")))
        );
        context.pop().expect("empty end");
        assert_eq!(
            context.resolve_element(QName(b"y")).unwrap().namespace,
            Some(NamespaceId::Unknown(Arc::from("urn:outer")))
        );
    }

    #[test]
    fn default_namespace_does_not_apply_to_unprefixed_attributes() {
        let xml = br#"<r xmlns="urn:outer" value="x"/>"#;
        let mut reader = Reader::from_reader(xml.as_slice());
        let start = parse_start(&mut reader);
        let mut context = SemanticNamespaceContext::new();
        context
            .start(&start, reader.decoder(), &mut no_charge)
            .expect("scope");
        assert_eq!(
            context.resolve_element(QName(b"r")).unwrap().namespace,
            Some(NamespaceId::Unknown(Arc::from("urn:outer")))
        );
        assert_eq!(
            context
                .resolve_attribute(QName(b"value"))
                .unwrap()
                .namespace,
            None
        );
    }

    #[test]
    fn start_and_push_aliases_construct_and_pop_scopes() {
        let xml = br#"<a/><b/>"#;
        let mut reader = Reader::from_reader(xml.as_slice());
        let mut context = SemanticNamespaceContext::new();

        let first = parse_start(&mut reader);
        context
            .start(&first, reader.decoder(), &mut no_charge)
            .expect("start alias");
        assert_eq!(context.depth(), 1);
        context.end().expect("end alias");

        let second = parse_start(&mut reader);
        context
            .push(&second, reader.decoder(), &mut no_charge)
            .expect("push alias");
        assert_eq!(context.depth(), 1);
        context.pop().expect("pop alias");
        assert!(context.is_empty());
    }

    #[test]
    fn caller_limits_cannot_raise_hard_bounds() {
        let context = SemanticNamespaceContext::with_limits(usize::MAX, usize::MAX);
        assert_eq!(context.max_depth, MAX_NAMESPACE_DEPTH);
        assert_eq!(
            context.max_declarations_per_element,
            MAX_NAMESPACE_DECLARATIONS_PER_ELEMENT
        );
    }

    #[test]
    fn normalizes_known_and_unknown_escaped_uris_and_shares_unknown_arc() {
        let xml = br#"<r xmlns:t="urn:oasis:names&#x3A;tc&#x3A;opendocument&#x3A;xmlns&#x3A;table&#x3A;1.0" xmlns:a="urn:opaque&#x3A;one" xmlns:b="urn:opaque:one"/>"#;
        let mut reader = Reader::from_reader(xml.as_slice());
        let start = parse_start(&mut reader);
        let mut context = SemanticNamespaceContext::new();
        context
            .start(&start, reader.decoder(), &mut no_charge)
            .expect("scope");
        assert_eq!(
            context.resolve_element(QName(b"t:cell")).unwrap().namespace,
            Some(NamespaceId::Known(Ns::Table))
        );
        let first = context.resolve_element(QName(b"a:x")).unwrap();
        let second = context.resolve_element(QName(b"b:x")).unwrap();
        let (Some(NamespaceId::Unknown(left)), Some(NamespaceId::Unknown(right))) =
            (first.namespace, second.namespace)
        else {
            panic!("unknown namespace expected");
        };
        assert_eq!(left.as_ref(), "urn:opaque:one");
        assert!(Arc::ptr_eq(&left, &right));
    }

    #[test]
    fn rejects_duplicate_expanded_attributes_even_through_aliases() {
        let xml = br#"<r xmlns:a="urn:x" xmlns:b="urn:x" a:value="1" b:value="2"/>"#;
        let mut reader = Reader::from_reader(xml.as_slice());
        let start = parse_start(&mut reader);
        let mut context = SemanticNamespaceContext::new();
        assert!(
            context
                .start(&start, reader.decoder(), &mut no_charge)
                .is_err()
        );
    }

    #[test]
    fn enforces_reserved_bindings_and_empty_prefixed_namespace() {
        for xml in [
            br#"<r xmlns="http://www.w3.org/XML/1998/namespace"/>"#.as_slice(),
            br#"<r xmlns:xml="urn:wrong"/>"#.as_slice(),
            br#"<r xmlns:xmlns="urn:wrong"/>"#.as_slice(),
            br#"<r xmlns:p=""/>"#.as_slice(),
            br#"<xmlns:r/>"#.as_slice(),
            br#"<r xmlns:p="urn:x" p:a="1" p:a="2"/>"#.as_slice(),
        ] {
            let mut reader = Reader::from_reader(xml);
            let start = parse_start(&mut reader);
            let mut context = SemanticNamespaceContext::new();
            assert!(
                context
                    .start(&start, reader.decoder(), &mut no_charge)
                    .is_err(),
                "{xml:?}"
            );
        }
    }

    #[test]
    fn rejects_unbound_prefixes_and_malformed_qnames() {
        let xml = br#"<r/>"#;
        let mut reader = Reader::from_reader(xml.as_slice());
        let start = parse_start(&mut reader);
        let mut context = SemanticNamespaceContext::new();
        context
            .start(&start, reader.decoder(), &mut no_charge)
            .expect("scope");
        assert!(context.resolve_element(QName(b"missing:item")).is_err());
        assert!(context.resolve_attribute(QName(b"missing:item")).is_err());
        assert!(validate_qname(QName(b"a:b:c"), "element").is_err());
        assert!(validate_qname(QName(b":local"), "element").is_err());
        assert!(validate_qname(QName(b"prefix:"), "element").is_err());
    }

    #[test]
    fn checks_depth_and_charge_before_namespace_value_decode() {
        let root_xml = br#"<r/>"#;
        let mut reader = Reader::from_reader(root_xml.as_slice());
        let root = parse_start(&mut reader);
        let mut context = SemanticNamespaceContext::with_limits(1, 256);
        context
            .start(&root, reader.decoder(), &mut no_charge)
            .expect("root scope");
        let child_xml = br#"<c/>"#;
        let mut child_reader = Reader::from_reader(child_xml.as_slice());
        let child = parse_start(&mut child_reader);
        assert!(
            context
                .start(&child, child_reader.decoder(), &mut no_charge)
                .is_err()
        );
        assert!(context.resolve_element(QName(b"r")).is_err());

        let escaped = br#"<r xmlns:p="urn:opaque&#x3A;value"/>"#;
        let mut escaped_reader = Reader::from_reader(escaped.as_slice());
        let escaped_start = parse_start(&mut escaped_reader);
        let mut escaped_context = SemanticNamespaceContext::new();
        let mut charged = false;
        let result =
            escaped_context.start(&escaped_start, escaped_reader.decoder(), &mut |bytes| {
                charged = true;
                if bytes > 0 {
                    return Err(Error::InvalidFormat("budget before decode".into()));
                }
                Ok(())
            });
        assert!(result.is_err());
        assert!(charged);
    }

    #[test]
    fn empty_default_binding_removes_inherited_default() {
        let xml = br#"<r xmlns="urn:outer"><c xmlns=""/></r>"#;
        let mut reader = Reader::from_reader(xml.as_slice());
        let mut context = SemanticNamespaceContext::new();
        let root = parse_start(&mut reader);
        context
            .start(&root, reader.decoder(), &mut no_charge)
            .expect("root");
        let child = parse_start(&mut reader);
        context
            .start(&child, reader.decoder(), &mut no_charge)
            .expect("child");
        assert_eq!(
            context.resolve_element(QName(b"x")).unwrap().namespace,
            None
        );
        context.pop().expect("child end");
        assert_eq!(
            context.resolve_element(QName(b"x")).unwrap().namespace,
            Some(NamespaceId::Unknown(Arc::from("urn:outer")))
        );
    }
}
