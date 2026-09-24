//! Namespace prefix bindings in scope, for the namespace constraints of
//! Namespaces in XML 1.0 (Third Edition): reserved prefixes and names, no
//! prefix undeclaring, every prefix declared, and attributes unique by
//! expanded name.
//!
//! Only prefixed bindings are kept. A default-namespace declaration is checked
//! where it stands and then forgotten: no well-formedness rule reads the
//! default namespace, because an unprefixed element name is always bound and
//! an unprefixed attribute is in no namespace.
//!
//! # Costs
//!
//! No per-tag operation reads a namespace name, whose bytes come from an
//! ancestor's declaration and are not bounded by the tag being checked.
//!
//! * Declaring a binding reads its prefix and its namespace name a constant
//!   number of times: once to normalize, once to hash, and once to compare with
//!   the name of an earlier binding in scope with the same keyed hash, which
//!   gives it that binding's identity (see [`Namespaces::names`]).
//! * Resolving a prefix hashes it and compares it with the prefix of the
//!   innermost binding with the same keyed hash, so it reads only bytes of the
//!   name being resolved; a prefix of at most seven bytes is usually answered
//!   by a direct-mapped cache instead.
//! * Checking that a tag's `k` prefixed attributes have distinct expanded names
//!   is `O(k)`: it compares namespace-name identities, not bytes, and only a
//!   tag in which two prefixes share a namespace name hashes its local names.
//! * Ending a scope is constant work per binding it removes.
//!
//! The keys of every hash are random per audit, so an input cannot choose the
//! collisions that would lengthen a comparison chain.

use core::hash::{BuildHasher, BuildHasherDefault, Hasher};
use std::collections::{HashMap, HashSet};
use std::hash::RandomState;

use super::{Error, wellformed};

/// The namespace name the `xml` prefix is bound to by definition.
pub(super) const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
/// The namespace name of namespace declarations, bound to `xmlns`.
pub(super) const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";

/// No binding and no name.
const NONE: usize = usize::MAX;
/// Slots of the direct-mapped lookup cache.
const RECENT: usize = 32;
/// An empty slot of the lookup cache: no packed prefix has this key.
const EMPTY: (u64, usize) = (0, NONE);

/// A table keyed by hashes that are already keyed.
type Index = HashMap<u64, usize, BuildHasherDefault<PassThrough>>;

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
    /// The keyed hash of the prefix: the binding's key in the prefix index.
    key: u64,
    /// The binding the prefix index held for the same key before this one:
    /// an outer binding of the same prefix or a hash collision, or [`NONE`].
    hidden: usize,
    /// The identity of the namespace name: an index into
    /// [`Namespaces::names`], shared by every binding in scope whose
    /// namespace name has the same bytes.
    name: usize,
}

/// One distinct namespace name in scope.
#[derive(Clone, Copy, Debug)]
struct Name {
    /// The keyed hash of the name: its key in the name index.
    key: u64,
    /// The binding that declared the name first, which holds its bytes.
    binding: usize,
    /// The name the name index held for the same key before this one: a hash
    /// collision, or [`NONE`].
    hidden: usize,
    /// The last start tag [`Namespaces::check_expanded_names`] saw the name
    /// in, as a count of the tags it has examined, and the binding it saw
    /// the name through there.
    seen: u64,
    seen_through: usize,
}

/// The prefixed bindings in scope, innermost last.
///
/// Every table holds only bindings of open elements: the bindings of an
/// element are dropped when it closes. Their bytes therefore come from at most
/// one start tag per open element and never exceed the input, and their number
/// is bounded by the attribute budget, since each is an attribute. The costs of
/// each operation are listed in the module documentation.
#[derive(Debug)]
pub(super) struct Namespaces {
    text: Vec<u8>,
    bindings: Vec<Binding>,
    /// Prefix hash to the innermost binding with that hash.
    prefixes: Index,
    /// The distinct namespace names of the bindings in scope, in the order
    /// they were first declared. A name is removed with the binding that
    /// declared it first; every later binding of the same name, and every
    /// later name, was declared after that binding and has already been
    /// removed, so this is a stack.
    names: Vec<Name>,
    /// Name hash to the latest name with that hash.
    name_index: Index,
    /// Keys every hash, so that its collisions cannot be chosen.
    keys: RandomState,
    /// Bindings recently resolved, keyed by their [`packed`] prefix in the
    /// slot the key selects; cleared whenever a binding is added or removed.
    recent: [(u64, usize); RECENT],
    /// Start tags examined by [`Self::check_expanded_names`].
    tags: u64,
}

impl Namespaces {
    pub(super) fn new() -> Self {
        Self {
            text: Vec::new(),
            bindings: Vec::new(),
            prefixes: Index::default(),
            names: Vec::new(),
            name_index: Index::default(),
            keys: RandomState::new(),
            recent: [EMPTY; RECENT],
            tags: 0,
        }
    }

    /// An independent copy with the same keys and name identities, or `None`
    /// if a table cannot be allocated.
    pub(super) fn try_clone(&self) -> Option<Self> {
        Some(Self {
            text: try_clone_vec(&self.text)?,
            bindings: try_clone_vec(&self.bindings)?,
            prefixes: try_clone_index(&self.prefixes)?,
            names: try_clone_vec(&self.names)?,
            name_index: try_clone_index(&self.name_index)?,
            keys: self.keys.clone(),
            recent: [EMPTY; RECENT],
            tags: self.tags,
        })
    }

    fn prefix(&self, binding: usize) -> &[u8] {
        let binding = &self.bindings[binding];
        &self.text[binding.start..binding.split]
    }

    /// The namespace name of `binding`. Every read of a kept namespace name
    /// goes through here, so that tests can count them.
    fn name(&self, binding: usize) -> &[u8] {
        let binding = &self.bindings[binding];
        let name = &self.text[binding.split..binding.end];
        #[cfg(test)]
        tests::count_name_bytes(name.len());
        name
    }

    fn hash(&self, bytes: &[u8]) -> u64 {
        self.keys.hash_one(bytes)
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
        self.prefixes
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation)?;
        self.names
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation)?;
        self.name_index
            .try_reserve(1)
            .map_err(|_allocation| Error::Allocation)?;
        let binding = self.bindings.len();
        let key = self.hash(&self.text[start..split]);
        let name = self.identify(split, end, binding);
        let hidden = self.prefixes.insert(key, binding).unwrap_or(NONE);
        self.bindings.push(Binding {
            start,
            split,
            end,
            depth,
            key,
            hidden,
            name,
        });
        self.recent = [EMPTY; RECENT];
        Ok(true)
    }

    /// The identity of the namespace name `text[split..end]`, which the new
    /// binding `binding` declares: the identity of a name in scope with the
    /// same bytes, or a new one. The name is hashed once and compared only
    /// with names in scope that have the same keyed hash, which, collisions
    /// aside, is at most the one it equals. The tables have room for one more
    /// entry.
    fn identify(&mut self, split: usize, end: usize, binding: usize) -> usize {
        let bytes = &self.text[split..end];
        #[cfg(test)]
        tests::count_name_bytes(bytes.len());
        let key = self.hash(bytes);
        let mut candidate = self.name_index.get(&key).copied().unwrap_or(NONE);
        while candidate != NONE {
            let known = self.names[candidate];
            if self.name(known.binding) == &self.text[split..end] {
                return candidate;
            }
            candidate = known.hidden;
        }
        let name = self.names.len();
        let hidden = self.name_index.insert(key, name).unwrap_or(NONE);
        self.names.push(Name {
            key,
            binding,
            hidden,
            seen: 0,
            seen_through: NONE,
        });
        name
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
            let index = self.bindings.len() - 1;
            // The keys are present, so these replace values in place.
            if binding.hidden == NONE {
                self.prefixes.remove(&binding.key);
            } else {
                self.prefixes.insert(binding.key, binding.hidden);
            }
            let name = self.names[binding.name];
            if name.binding == index {
                debug_assert_eq!(binding.name + 1, self.names.len());
                if name.hidden == NONE {
                    self.name_index.remove(&name.key);
                } else {
                    self.name_index.insert(name.key, name.hidden);
                }
                self.names.pop();
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

    /// [`Self::resolve`] through the prefix index.
    fn resolve_indexed(&mut self, prefix: &[u8]) -> Option<usize> {
        let mut binding = *self.prefixes.get(&self.hash(prefix))?;
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
    /// [`Error::Malformed`] at the offending name, and [`Error::Allocation`].
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
    /// "Attributes Unique"). Equal qualified names are refused earlier, so
    /// such a pair needs two bindings in this tag with one namespace name.
    /// Finding whether there are any marks each name with the first binding
    /// the tag reaches it through: constant work per attribute, comparing
    /// name identities rather than name bytes.
    fn check_expanded_names(
        &mut self,
        inner: &[u8],
        prefixed: &[Prefixed],
        offset: usize,
    ) -> Result<(), Error> {
        self.tags += 1;
        let tag = self.tags;
        let mut shared = false;
        for attribute in prefixed {
            #[cfg(test)]
            tests::count_expanded_name_step();
            let identity = self.bindings[attribute.binding].name;
            let name = &mut self.names[identity];
            if name.seen != tag {
                name.seen = tag;
                name.seen_through = attribute.binding;
            } else if name.seen_through != attribute.binding {
                shared = true;
                break;
            }
        }
        if shared {
            self.find_duplicate_expanded_name(inner, prefixed, offset)
        } else {
            Ok(())
        }
    }

    /// For a tag in which two prefixes share a namespace name: refuses the
    /// first attribute, in document order, whose expanded name an earlier
    /// attribute has. Each attribute's name identity and local name are
    /// hashed once, with keys an input cannot predict.
    #[cold]
    fn find_duplicate_expanded_name(
        &self,
        inner: &[u8],
        prefixed: &[Prefixed],
        offset: usize,
    ) -> Result<(), Error> {
        let mut seen = HashSet::with_hasher(self.keys.clone());
        seen.try_reserve(prefixed.len())
            .map_err(|_allocation| Error::Allocation)?;
        for attribute in prefixed {
            #[cfg(test)]
            tests::count_expanded_name_step();
            let local = &inner[attribute.colon + 1..attribute.end];
            if !seen.insert(ExpandedName {
                name: self.bindings[attribute.binding].name,
                local,
            }) {
                return Err(Error::malformed(
                    offset + attribute.start,
                    "two attributes have the same namespace name and local name",
                ));
            }
        }
        Ok(())
    }
}

/// An attribute's expanded name within one start tag: the identity of its
/// namespace name and its local name.
#[derive(Eq, Hash, PartialEq)]
struct ExpandedName<'a> {
    name: usize,
    local: &'a [u8],
}

/// A copy of `items`, or `None` if it cannot be allocated.
fn try_clone_vec<T: Copy>(items: &[T]) -> Option<Vec<T>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(items.len()).ok()?;
    copy.extend_from_slice(items);
    Some(copy)
}

/// A copy of `index`, or `None` if it cannot be allocated.
fn try_clone_index(index: &Index) -> Option<Index> {
    let mut copy = Index::default();
    copy.try_reserve(index.len()).ok()?;
    copy.extend(index.iter().map(|(key, value)| (*key, *value)));
    Some(copy)
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

/// The hasher of the prefix and name indexes, whose keys are already keyed
/// hashes.
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

#[cfg(test)]
mod tests {
    //! Operation bounds. Namespace names in these inputs are megabytes long
    //! and declared by ancestors of every checked tag, so any per-tag work
    //! proportional to their length shows up in the counts.

    use core::cell::Cell;

    use crate::audit::{
        Error, Limits, ReplacementError, ReplacementProof, Resource, verify_source,
        verify_source_replacement,
    };

    thread_local! {
        static NAME_BYTES: Cell<usize> = const { Cell::new(0) };
        static EXPANDED_NAME_STEPS: Cell<usize> = const { Cell::new(0) };
    }

    pub(super) fn count_name_bytes(bytes: usize) {
        NAME_BYTES.with(|count| count.set(count.get() + bytes));
    }

    pub(super) fn count_expanded_name_step() {
        EXPANDED_NAME_STEPS.with(|count| count.set(count.get() + 1));
    }

    /// Runs `audit` and returns its result with the namespace-name bytes it
    /// read and its steps of expanded-name checking.
    fn counted<T>(audit: impl FnOnce() -> T) -> (T, usize, usize) {
        NAME_BYTES.with(|count| count.set(0));
        EXPANDED_NAME_STEPS.with(|count| count.set(0));
        let result = audit();
        (
            result,
            NAME_BYTES.with(Cell::get),
            EXPANDED_NAME_STEPS.with(Cell::get),
        )
    }

    /// Three nested declarations of namespace names of `length` bytes, the
    /// first two equal when `alias` is set, around `body`.
    fn nested(length: usize, alias: bool, body: &str) -> (Vec<u8>, usize) {
        let base = "u".repeat(length - 1);
        let second = if alias { '1' } else { '2' };
        let document = format!(
            r#"<r xmlns:a="{base}1"><s xmlns:b="{base}{second}"><t xmlns:c="{base}3">{body}</t></s></r>"#
        );
        (document.into_bytes(), 3 * length)
    }

    /// `count` attributes cycling through the prefixes `a`, `b` and `c`, each
    /// with its own local name.
    fn many_attributes(count: usize) -> String {
        let attributes: Vec<String> = (0..count)
            .map(|index| format!(r#"{}:k{index}="""#, ["a", "b", "c"][index % 3]))
            .collect();
        format!("<x {}/>", attributes.join(" "))
    }

    #[test]
    fn a_tag_never_reads_the_namespace_names_its_ancestors_declared() {
        // The reviewer's input was one tag with 249,990 attributes under
        // three 3.9 MB namespace names that differ only in their last byte.
        // The per-element attribute limit now refuses that tag at its first
        // surplus attribute, before any namespace work for it, so the cost
        // property is checked on the largest tag a configuration can admit.
        let length = 3_900_000;
        let reviewer = 249_990;
        let (document, declared) = nested(length, false, &many_attributes(reviewer));
        let (result, name_bytes, steps) = counted(|| verify_source(&document, Limits::default()));
        let limit = Limits::DEFAULT_ELEMENT_ATTRIBUTES;
        assert!(
            matches!(
                result,
                Err(Error::Limit {
                    resource: Resource::ElementAttributes,
                    limit: refused_at,
                    actual,
                    ..
                }) if refused_at == limit && actual == limit + 1
            ),
            "{result:?}"
        );
        assert!(name_bytes <= declared, "{name_bytes} name bytes read");
        assert!(steps <= limit, "{steps} expanded-name steps");

        let attributes = Limits::ELEMENT_ATTRIBUTE_CEILING;
        let widest = Limits::builder()
            .element_attributes(attributes)
            .expect("the ceiling is a valid per-element limit")
            .build();
        let (document, declared) = nested(length, false, &many_attributes(attributes));
        let (result, name_bytes, steps) = counted(|| verify_source(&document, widest));
        assert!(result.is_ok(), "{result:?}");
        // Each declaration is hashed once; none is compared, being new.
        assert!(name_bytes <= declared, "{name_bytes} name bytes read");
        assert!(steps <= attributes, "{steps} expanded-name steps");

        // The same with the first two names equal: the second declaration is
        // compared with the first once, and the tag, which now has two
        // prefixes bound to one name, hashes each attribute once more.
        let (document, declared) = nested(length, true, &many_attributes(attributes));
        let (result, name_bytes, steps) = counted(|| verify_source(&document, widest));
        assert!(result.is_ok(), "{result:?}");
        assert!(name_bytes <= 2 * declared, "{name_bytes} name bytes read");
        assert!(steps <= 2 * attributes, "{steps} expanded-name steps");
    }

    #[test]
    fn many_tags_under_long_namespace_names_cost_no_name_reads() {
        // 50,000 tags that each use two prefixes bound to 3.9 MB names.
        let tags = 50_000;
        for alias in [false, true] {
            let body = r#"<x a:k="" b:j=""/>"#.repeat(tags);
            let (document, declared) = nested(3_900_000, alias, &body);
            let (result, name_bytes, steps) =
                counted(|| verify_source(&document, Limits::default()));
            assert!(result.is_ok(), "alias {alias}: {result:?}");
            assert!(name_bytes <= 2 * declared, "alias {alias}: {name_bytes}");
            assert!(steps <= 4 * tags, "alias {alias}: {steps}");
        }
        // With equal local names, the aliased prefixes give the first tag a
        // duplicate expanded name, found without reading either name.
        let body = r#"<x a:k="" b:k=""/>"#.repeat(tags);
        let (document, declared) = nested(3_900_000, true, &body);
        let (result, name_bytes, _steps) = counted(|| verify_source(&document, Limits::default()));
        // Each name is 3,900,000 bytes where the pattern below has one.
        let opening =
            r#"<r xmlns:a="1"><s xmlns:b="1"><t xmlns:c="3">"#.len() + 3 * (3_900_000 - 1);
        assert!(
            matches!(
                &result,
                Err(Error::Malformed { offset, detail })
                    if *offset == opening + r#"<x a:k="" "#.len()
                        && detail.contains("same namespace name and local name")
            ),
            "{result:?}"
        );
        assert!(name_bytes <= 2 * declared, "{name_bytes} name bytes read");
    }

    #[test]
    fn a_window_replays_name_identities_without_reading_names() {
        // A one-byte edit inside one of 50,000 tags under two aliased 3.9 MB
        // names: the window proof starts from a copy of the bindings in
        // scope, whose identities it reuses.
        let tags = 50_000;
        let body = r#"<x a:k="" b:j=""/>"#.repeat(tags);
        let (original, declared) = nested(3_900_000, true, &body);
        let at = original.len() / 2;
        let edit = at
            + original[at..]
                .windows(2)
                .position(|pair| pair == b"j=")
                .unwrap();
        let mut replacement = original.clone();
        replacement[edit] = b'i';
        let (result, name_bytes, _steps) =
            counted(|| verify_source_replacement(&original, &replacement, Limits::default()));
        assert!(
            matches!(result, Ok(ReplacementProof::Window { .. })),
            "{result:?}"
        );
        // The original's complete audit, and in debug builds the complete
        // audit that re-derives the window proof, each read every name at
        // most twice; the window itself reads none.
        let complete_audits = if cfg!(debug_assertions) { 2 } else { 1 };
        assert!(
            name_bytes <= complete_audits * 2 * declared,
            "{name_bytes} name bytes read"
        );
        // Turning the edited attribute's local name into its neighbour's is
        // refused inside the window, as the complete audit refuses it.
        replacement[edit] = b'k';
        let (result, _name_bytes, _steps) =
            counted(|| verify_source_replacement(&original, &replacement, Limits::default()));
        assert!(
            matches!(
                &result,
                Err(ReplacementError::Replacement(Error::Malformed { detail, .. }))
                    if detail.contains("same namespace name")
            ),
            "{result:?}"
        );
    }
}
