//! Namespace prefix bindings in scope, for the namespace constraints of
//! Namespaces in XML 1.0 (Third Edition): reserved prefixes and names, no
//! prefix undeclaring, every prefix declared, and attributes unique by
//! expanded name.
//!
//! Only prefixed bindings are kept. A default-namespace declaration is checked
//! where it stands and then forgotten: no well-formedness rule reads the
//! default namespace, because an unprefixed element name is always bound and
//! an unprefixed attribute is in no namespace.

use core::hash::{BuildHasher, BuildHasherDefault, Hasher};
use std::collections::HashMap;
use std::hash::RandomState;

use super::{Error, wellformed};

/// The namespace name the `xml` prefix is bound to by definition.
pub(super) const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
/// The namespace name of namespace declarations, bound to `xmlns`.
pub(super) const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";

/// No binding.
const NONE: usize = usize::MAX;
/// Slots of the direct-mapped lookup cache.
const RECENT: usize = 32;
/// An empty slot of the lookup cache: no packed prefix has this key.
const EMPTY: (u64, usize) = (0, NONE);

/// One prefix binding in scope.
#[derive(Clone, Copy, Debug)]
struct Binding {
    /// `text[start..split]` is the prefix.
    start: usize,
    /// `text[split..end]` is the normalized namespace name.
    split: usize,
    end: usize,
    /// Depth of the element whose start tag declared the binding.
    depth: usize,
    /// The binding the index held for the same prefix hash before this one:
    /// an outer binding of the same prefix or a hash collision, or [`NONE`].
    hidden: usize,
}

/// The prefixed bindings in scope, innermost last.
///
/// Every table holds only bindings of open elements: the bindings of an
/// element are dropped when it closes. Their bytes therefore come from at most
/// one start tag per open element and never exceed the input, and their number
/// is bounded by the attribute budget, since each is an attribute. Lookups go
/// through a keyed hash index, so an input cannot make them slower by piling
/// up bindings.
#[derive(Debug)]
pub(super) struct Namespaces {
    text: Vec<u8>,
    bindings: Vec<Binding>,
    /// Prefix hash to the innermost binding with that hash.
    index: HashMap<u64, usize, BuildHasherDefault<PassThrough>>,
    /// Keys the prefix hash, so that its collisions cannot be chosen.
    keys: RandomState,
    /// Bindings recently resolved, keyed by their [`packed`] prefix in the
    /// slot the key selects; cleared whenever a binding is added or removed.
    recent: [(u64, usize); RECENT],
}

impl Namespaces {
    pub(super) fn new() -> Self {
        Self {
            text: Vec::new(),
            bindings: Vec::new(),
            index: HashMap::default(),
            keys: RandomState::new(),
            recent: [EMPTY; RECENT],
        }
    }

    /// An independent copy, or `None` if a table cannot be allocated.
    pub(super) fn try_clone(&self) -> Option<Self> {
        let mut text = Vec::new();
        text.try_reserve_exact(self.text.len()).ok()?;
        text.extend_from_slice(&self.text);
        let mut bindings = Vec::new();
        bindings.try_reserve_exact(self.bindings.len()).ok()?;
        bindings.extend_from_slice(&self.bindings);
        let mut index = HashMap::default();
        index.try_reserve(self.index.len()).ok()?;
        index.extend(self.index.iter().map(|(hash, binding)| (*hash, *binding)));
        Some(Self {
            text,
            bindings,
            index,
            keys: self.keys.clone(),
            recent: [EMPTY; RECENT],
        })
    }

    fn prefix(&self, binding: usize) -> &[u8] {
        let binding = &self.bindings[binding];
        &self.text[binding.start..binding.split]
    }

    fn name(&self, binding: usize) -> &[u8] {
        let binding = &self.bindings[binding];
        &self.text[binding.split..binding.end]
    }

    fn hash(&self, prefix: &[u8]) -> u64 {
        self.keys.hash_one(prefix)
    }

    /// Checks one namespace declaration of the element at `depth` and keeps
    /// the binding it makes. `prefix` is `None` for `xmlns` and the declared
    /// prefix for `xmlns:prefix`; `value` is the attribute value as written,
    /// which [`wellformed::check_attribute_value`] has accepted. `offset` is
    /// where the declaration's name starts.
    ///
    /// # Errors
    ///
    /// [`Error::Malformed`] for a declaration Namespaces in XML 1.0 forbids,
    /// and [`Error::Allocation`].
    pub(super) fn declare(
        &mut self,
        prefix: Option<&[u8]>,
        value: &[u8],
        depth: usize,
        offset: usize,
    ) -> Result<(), Error> {
        if prefix == Some(b"xmlns") {
            return Err(Error::malformed(
                offset,
                "the xmlns prefix must not be declared",
            ));
        }
        let start = self.text.len();
        // A normalized value is never longer than the value as written.
        let bytes = prefix.map_or(0, <[u8]>::len).saturating_add(value.len());
        self.text
            .try_reserve(bytes)
            .map_err(|_allocation| Error::Allocation)?;
        if let Some(prefix) = prefix {
            self.text.extend_from_slice(prefix);
        }
        let split = self.text.len();
        wellformed::push_normalized_value(value, &mut self.text);
        let end = self.text.len();
        let result = self.bind(prefix, start, split, end, depth, offset);
        if !matches!(result, Ok(true)) {
            self.text.truncate(start);
        }
        result.map(|_kept| ())
    }

    /// Checks a declaration whose bytes `declare` has just appended, and
    /// keeps it when it binds an ordinary prefix. Returns whether it was kept.
    fn bind(
        &mut self,
        prefix: Option<&[u8]>,
        start: usize,
        split: usize,
        end: usize,
        depth: usize,
        offset: usize,
    ) -> Result<bool, Error> {
        let name = &self.text[split..end];
        let Some(prefix) = prefix else {
            if name == XML_NAMESPACE || name == XMLNS_NAMESPACE {
                return Err(Error::malformed(
                    offset,
                    "the default namespace must not be a reserved namespace name",
                ));
            }
            return Ok(false);
        };
        if prefix == b"xml" {
            if name != XML_NAMESPACE {
                return Err(Error::malformed(
                    offset,
                    "the xml prefix must not be bound to another namespace name",
                ));
            }
            return Ok(false);
        }
        if name.is_empty() {
            return Err(Error::malformed(
                offset,
                "a namespace prefix must not be undeclared",
            ));
        }
        if name == XML_NAMESPACE || name == XMLNS_NAMESPACE {
            return Err(Error::malformed(
                offset,
                "only a reserved prefix may be bound to a reserved namespace name",
            ));
        }
        self.bindings
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation)?;
        self.index
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation)?;
        let hash = self.hash(&self.text[start..split]);
        let binding = self.bindings.len();
        let hidden = self.index.insert(hash, binding).unwrap_or(NONE);
        self.bindings.push(Binding {
            start,
            split,
            end,
            depth,
            hidden,
        });
        self.recent = [EMPTY; RECENT];
        Ok(true)
    }

    /// Ends the scope of the bindings declared by the element at `depth`.
    #[inline]
    pub(super) fn close(&mut self, depth: usize) {
        if self
            .bindings
            .last()
            .is_some_and(|binding| binding.depth >= depth)
        {
            self.close_scope(depth);
        }
    }

    #[cold]
    fn close_scope(&mut self, depth: usize) {
        let mut closed = false;
        while let Some(binding) = self.bindings.last().copied()
            && binding.depth >= depth
        {
            let hash = self.hash(&self.text[binding.start..binding.split]);
            if binding.hidden == NONE {
                self.index.remove(&hash);
            } else {
                // The key is present, so this replaces a value in place.
                self.index.insert(hash, binding.hidden);
            }
            self.bindings.pop();
            self.text.truncate(binding.start);
            closed = true;
        }
        if closed {
            self.recent = [EMPTY; RECENT];
        }
    }

    /// The innermost binding of `prefix` in scope.
    #[inline]
    fn resolve(&mut self, prefix: &[u8]) -> Option<usize> {
        let Some(key) = packed(prefix) else {
            return self.resolve_indexed(prefix);
        };
        let slot = slot(key);
        let (cached, binding) = self.recent[slot];
        if cached == key {
            return Some(binding);
        }
        let binding = self.resolve_indexed(prefix)?;
        self.recent[slot] = (key, binding);
        Some(binding)
    }

    /// [`Self::resolve`] through the prefix index, remembering the answer.
    fn resolve_indexed(&mut self, prefix: &[u8]) -> Option<usize> {
        let mut binding = *self.index.get(&self.hash(prefix))?;
        while binding != NONE {
            if self.prefix(binding) == prefix {
                return Some(binding);
            }
            binding = self.bindings[binding].hidden;
        }
        None
    }

    /// Checks the names of one start tag once its declarations are bound:
    /// the element's prefix is not `xmlns` and is declared, each attribute
    /// prefix in `prefixed` is declared, and no two attributes share an
    /// expanded name. `inner` is the tag between `<` and its close, starting
    /// at byte `offset`; `colon` is the element name's colon, if any.
    ///
    /// # Errors
    ///
    /// [`Error::Malformed`] at the offending name.
    #[inline]
    pub(super) fn check_tag(
        &mut self,
        inner: &[u8],
        colon: Option<usize>,
        prefixed: &mut [Prefixed],
        offset: usize,
    ) -> Result<(), Error> {
        match colon {
            None if prefixed.is_empty() => return Ok(()),
            // The common case: a prefix already resolved, which every
            // prefixed attribute shares. One binding cannot give two
            // attributes one expanded name without one qualified name, which
            // the tokenizer refuses.
            Some(colon) if self.cached(&inner[..colon]) => {
                let prefix = &inner[..colon];
                if prefixed
                    .iter()
                    .all(|attribute| same_bytes(&inner[attribute.start..attribute.colon], prefix))
                {
                    return Ok(());
                }
            },
            _ => {},
        }
        self.check_prefixed_tag(inner, colon, prefixed, offset)
    }

    /// Whether the lookup cache holds a binding for `prefix`.
    #[inline]
    fn cached(&self, prefix: &[u8]) -> bool {
        packed(prefix).is_some_and(|key| self.recent[slot(key)].0 == key)
    }

    fn check_prefixed_tag(
        &mut self,
        inner: &[u8],
        colon: Option<usize>,
        prefixed: &mut [Prefixed],
        offset: usize,
    ) -> Result<(), Error> {
        // The last prefix resolved and its binding: an attribute usually
        // shares its element's prefix, or the attribute's before it.
        let mut last: Option<(&[u8], usize)> = None;
        if let Some(colon) = colon {
            let prefix = &inner[..colon];
            if prefix == b"xmlns" {
                return Err(Error::malformed(
                    offset,
                    "an element name must not use the xmlns prefix",
                ));
            }
            if prefix != b"xml" {
                let binding = self
                    .resolve(prefix)
                    .ok_or_else(|| Error::malformed(offset, "undeclared namespace prefix"))?;
                last = Some((prefix, binding));
            }
        }
        for attribute in prefixed.iter_mut() {
            let prefix = &inner[attribute.start..attribute.colon];
            attribute.binding = match last {
                Some((known, binding)) if same_bytes(known, prefix) => binding,
                _ => {
                    let binding = self.resolve(prefix).ok_or_else(|| {
                        Error::malformed(offset + attribute.start, "undeclared namespace prefix")
                    })?;
                    last = Some((prefix, binding));
                    binding
                },
            };
        }
        if prefixed.len() < 2 {
            return Ok(());
        }
        self.check_expanded_names(inner, prefixed, offset)
    }

    /// Refuses two attributes with the same local name whose prefixes are
    /// bound to the same namespace name (Namespaces in XML 1.0, constraint
    /// "Attributes Unique"). Equal qualified names are refused earlier.
    fn check_expanded_names(
        &self,
        inner: &[u8],
        prefixed: &mut [Prefixed],
        offset: usize,
    ) -> Result<(), Error> {
        let Some(first) = prefixed.first().map(|attribute| attribute.binding) else {
            return Ok(());
        };
        let Some(second) = prefixed
            .iter()
            .map(|attribute| attribute.binding)
            .find(|binding| *binding != first)
        else {
            // One prefix: equal expanded names would be equal qualified names.
            return Ok(());
        };
        let two_prefixes = prefixed
            .iter()
            .all(|attribute| attribute.binding == first || attribute.binding == second);
        if two_prefixes && self.name(first) != self.name(second) {
            return Ok(());
        }
        let local = |attribute: &Prefixed| &inner[attribute.colon + 1..attribute.end];
        prefixed.sort_unstable_by(|left, right| {
            self.name(left.binding)
                .cmp(self.name(right.binding))
                .then_with(|| local(left).cmp(local(right)))
        });
        for pair in prefixed.windows(2) {
            if self.name(pair[0].binding) == self.name(pair[1].binding)
                && local(&pair[0]) == local(&pair[1])
            {
                return Err(Error::malformed(
                    offset + pair[0].start.max(pair[1].start),
                    "two attributes have the same namespace name and local name",
                ));
            }
        }
        Ok(())
    }
}

/// The lookup-cache slot of a [`packed`] key.
#[inline]
const fn slot(key: u64) -> usize {
    (key ^ (key >> 29)) as usize % RECENT
}

/// A prefix of one to seven bytes as a lookup-cache key: its bytes and its
/// length, which is never zero, so no key equals an empty slot's. `None` for
/// a longer prefix, which the index resolves directly.
#[inline]
fn packed(prefix: &[u8]) -> Option<u64> {
    if prefix.is_empty() || prefix.len() > 7 {
        return None;
    }
    let mut key = prefix.len() as u64;
    for &byte in prefix {
        key = (key << 8) | u64::from(byte);
    }
    Some(key)
}

/// Byte equality without a library call, for the short prefixes compared on
/// every prefixed name.
#[inline]
fn same_bytes(left: &[u8], right: &[u8]) -> bool {
    left.len() == right.len() && left.iter().zip(right).all(|(left, right)| left == right)
}

/// A prefixed attribute of the start tag being checked, other than a
/// namespace declaration or an `xml:` attribute. Offsets are within the tag's
/// text between `<` and its close.
#[derive(Clone, Copy, Debug)]
pub(super) struct Prefixed {
    pub(super) start: usize,
    pub(super) colon: usize,
    pub(super) end: usize,
    /// The binding its prefix resolves to, once resolved.
    pub(super) binding: usize,
}

impl Prefixed {
    pub(super) const fn new(start: usize, colon: usize, end: usize) -> Self {
        Self {
            start,
            colon,
            end,
            binding: NONE,
        }
    }
}

/// The hasher of the prefix index, whose keys are already keyed hashes.
#[derive(Default)]
struct PassThrough(u64);

impl Hasher for PassThrough {
    fn finish(&self) -> u64 {
        self.0
    }

    fn write(&mut self, bytes: &[u8]) {
        for byte in bytes {
            self.0 = self.0.rotate_left(8) ^ u64::from(*byte);
        }
    }

    fn write_u64(&mut self, value: u64) {
        self.0 = value;
    }
}
