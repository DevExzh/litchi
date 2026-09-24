//! Namespace bindings in scope, indexed by prefix, for the MCE processors.
//!
//! Both processors also keep a persistent chain of per-element declaration
//! layers. Resolving a prefix by walking that chain costs one comparison per
//! declaration in scope, and the input chooses that number: the declarations
//! of every open ancestor count, shadowed re-declarations included. [`Scope`]
//! walks at most [`WALKED_DECLARATIONS`] of them, innermost first, and answers
//! from an index past that, so a lookup never costs more.
//!
//! It is kept in step with the element stack: an element's declarations are
//! added with its depth once its start tag has been checked, and
//! [`Scope::truncate`] removes them when it closes. Depths are positive and
//! strictly increase from an element to its children, so the bindings of one
//! prefix are always ordered by depth, outermost first.
//!
//! # Namespace identities
//!
//! Every namespace URI in scope has an identity, a [`UriId`], shared by every
//! binding in scope whose URI has the same bytes. A URI is hashed once, with a
//! key random per scope, when it is bound, and compared only with URIs in
//! scope that have the same keyed hash; the facts the processors ask about a
//! namespace (whether the profile understands it, whether it is the markup
//! compatibility namespace, whether an extension element lives in it) are
//! computed then. Directive sets and patterns hold identities, so checking an
//! element or attribute name against them costs the same whatever the length
//! of its namespace URI, which the input chooses.
//!
//! An identity lives while a binding of its URI is in scope. An element binds
//! its URIs after its ancestors and releases them before they do, so the
//! identities form a stack; a directive layer holds only identities of URIs
//! bound by its element or an ancestor, and is dropped when that element
//! closes, so it never sees a released identity reused.

use core::hash::{BuildHasher, BuildHasherDefault, Hasher};
use std::collections::{BTreeMap, HashMap};
use std::hash::RandomState;

use super::model::{Capabilities, Error, NAMESPACE, XML_NS};

/// Declarations a lookup walks, innermost first, before it asks the index. It
/// covers every declaration in scope in the repository's real documents, whose
/// largest element declares 38.
pub(super) const WALKED_DECLARATIONS: usize = 64;

/// The identity of a namespace in scope: equal URIs in scope share one.
pub(super) type UriId = usize;

/// No namespace: an unprefixed attribute, or an unprefixed element with no
/// default namespace in scope.
pub(super) const NO_NAMESPACE: UriId = usize::MAX;

/// The namespace the `xml` prefix is bound to by definition.
pub(super) const XML_URI: UriId = usize::MAX - 1;

/// No entry, in the collision chains of [`Scope::uri_index`].
const NONE: usize = usize::MAX;

/// A table keyed by hashes that are already keyed.
type Index = HashMap<u64, usize, BuildHasherDefault<PassThrough>>;

/// A resolved namespace: its URI and its identity.
#[derive(Clone, Copy, Debug)]
pub(super) struct Uri<'a> {
    pub(super) text: &'a str,
    pub(super) id: UriId,
}

impl Uri<'_> {
    /// No namespace.
    pub(super) const NONE: Uri<'static> = Uri {
        text: "",
        id: NO_NAMESPACE,
    };
    /// The `xml` namespace.
    pub(super) const XML: Uri<'static> = Uri {
        text: XML_NS,
        id: XML_URI,
    };
}

/// One distinct namespace URI in scope and the facts about it.
#[derive(Debug)]
struct UriEntry {
    text: String,
    /// The keyed hash of `text`: the entry's key in the index.
    key: u64,
    /// The entry the index held for the same key before this one, or
    /// [`NONE`].
    hidden: usize,
    /// Bindings in scope with this URI.
    refs: usize,
    facts: Facts,
}

/// What the processors ask about one namespace, computed once.
#[derive(Clone, Copy, Debug, Default)]
struct Facts {
    understood: bool,
    mce: bool,
    xml: bool,
    extension: bool,
}

impl Facts {
    fn of(text: &str, capabilities: &Capabilities) -> Self {
        #[cfg(test)]
        counter::uri_bytes(text.len());
        Self {
            understood: capabilities.understands(text),
            mce: text == NAMESPACE,
            xml: text == XML_NS,
            extension: capabilities
                .extensions
                .iter()
                .any(|name| name.namespace == text),
        }
    }
}

/// One binding in scope, in the order bindings were added.
#[derive(Debug)]
struct Added {
    depth: usize,
    prefix: String,
    /// [`packed`] `prefix`, compared first so that the walk of a lookup
    /// compares integers rather than calling `memcmp` for each short prefix.
    key: u64,
    uri: UriId,
}

/// A prefix of at most seven bytes as one integer, its length in the top byte:
/// equal for two such prefixes exactly when they are equal. Longer prefixes
/// share [`LONG_PREFIX`] and are compared as strings.
fn packed(prefix: &str) -> u64 {
    let bytes = prefix.as_bytes();
    if bytes.len() > 7 {
        return LONG_PREFIX;
    }
    let mut key = (bytes.len() as u64) << 56;
    for (index, byte) in bytes.iter().enumerate() {
        key |= u64::from(*byte) << (8 * index);
    }
    key
}

/// The [`packed`] key of every prefix longer than seven bytes.
const LONG_PREFIX: u64 = u64::MAX;

/// The bindings in scope at the element being processed.
#[derive(Debug)]
pub(super) struct Scope {
    /// Each prefix's bindings, outermost first, with the depth of the element
    /// that declared each.
    prefixes: BTreeMap<String, Vec<(usize, UriId)>>,
    /// Every binding in `prefixes`, in the order it was added.
    added: Vec<Added>,
    /// The distinct URIs in scope, by identity; a stack.
    uris: Vec<UriEntry>,
    /// Keyed URI hash to the latest identity with that hash.
    uri_index: Index,
    /// Keys every hash, so that its collisions cannot be chosen.
    keys: RandomState,
    /// The facts of [`XML_URI`] and [`NO_NAMESPACE`].
    xml: Facts,
    none: Facts,
}

impl Scope {
    /// An empty scope for a processor with `capabilities`.
    pub(super) fn new(capabilities: &Capabilities) -> Self {
        Self {
            prefixes: BTreeMap::new(),
            added: Vec::new(),
            uris: Vec::new(),
            uri_index: Index::default(),
            keys: RandomState::new(),
            xml: Facts::of(XML_NS, capabilities),
            none: Facts::of("", capabilities),
        }
    }

    /// The namespace `prefix` is bound to in scope, with its identity.
    #[inline]
    pub(super) fn resolve(&self, prefix: &str) -> Option<Uri<'_>> {
        if prefix == "xml" {
            return Some(Uri::XML);
        }
        let key = packed(prefix);
        let mut walked = 0usize;
        let mut found = None;
        for binding in self.added.iter().rev().take(WALKED_DECLARATIONS) {
            walked += 1;
            if binding.key == key && (key != LONG_PREFIX || binding.prefix == prefix) {
                found = Some(binding.uri);
                break;
            }
        }
        #[cfg(test)]
        counter::steps(walked);
        if let Some(id) = found {
            return Some(self.uri(id));
        }
        if walked == self.added.len() {
            return None;
        }
        #[cfg(test)]
        counter::lookup();
        self.prefixes
            .get(prefix)
            .and_then(|bindings| bindings.last())
            .map(|(_, id)| self.uri(*id))
    }

    /// The namespace `prefix` is bound to in scope.
    pub(super) fn get(&self, prefix: &str) -> Option<&str> {
        self.resolve(prefix).map(|uri| uri.text)
    }

    /// The URI and identity of `id`.
    #[inline]
    fn uri(&self, id: UriId) -> Uri<'_> {
        match id {
            NO_NAMESPACE => Uri::NONE,
            XML_URI => Uri::XML,
            _ => Uri {
                text: &self.uris[id].text,
                id,
            },
        }
    }

    #[inline]
    fn facts(&self, id: UriId) -> Facts {
        match id {
            NO_NAMESPACE => self.none,
            XML_URI => self.xml,
            _ => self.uris[id].facts,
        }
    }

    /// Whether the processor's profile understands the namespace `id`.
    #[inline]
    pub(super) fn understood(&self, id: UriId) -> bool {
        self.facts(id).understood
    }

    /// Whether `id` is the markup compatibility namespace.
    #[inline]
    pub(super) fn is_mce(&self, id: UriId) -> bool {
        self.facts(id).mce
    }

    /// Whether `id` is the `xml` namespace, whatever prefix bound it.
    #[inline]
    pub(super) fn is_xml(&self, id: UriId) -> bool {
        self.facts(id).xml
    }

    /// Whether an extension element of the profile is in the namespace `id`.
    #[inline]
    pub(super) fn has_extensions(&self, id: UriId) -> bool {
        self.facts(id).extension
    }

    /// The namespace `prefix` is bound to by an element shallower than
    /// `depth`, that is in scope at the parent of the element at `depth`.
    pub(super) fn get_outside(&self, prefix: &str, depth: usize) -> Option<&str> {
        #[cfg(test)]
        counter::lookup();
        if prefix == "xml" {
            return Some(XML_NS);
        }
        // One element binds a prefix at most once, so the binding sought is
        // the last or the one before it.
        self.prefixes.get(prefix).and_then(|bindings| {
            bindings
                .iter()
                .rev()
                .take(2)
                .find(|(declared, _)| *declared < depth)
                .map(|(_, id)| self.uri(*id).text)
        })
    }

    /// Whether the element at `depth` itself declares `prefix`.
    pub(super) fn declared_at(&self, prefix: &str, depth: usize) -> bool {
        #[cfg(test)]
        counter::lookup();
        self.prefixes
            .get(prefix)
            .and_then(|bindings| bindings.last())
            .is_some_and(|(declared, _)| *declared == depth)
    }

    /// Add the declarations of the element at `depth`, which binds each
    /// prefix at most once and is deeper than every binding in scope, giving
    /// each URI its identity for a processor with `capabilities`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`] when a table cannot grow.
    pub(super) fn push(
        &mut self,
        depth: usize,
        local: &[(String, String)],
        capabilities: &Capabilities,
    ) -> Result<(), Error> {
        let allocation = |source| Error::Allocation {
            resource: "MCE namespace scope",
            source,
        };
        self.added.try_reserve(local.len()).map_err(allocation)?;
        for (prefix, namespace) in local {
            debug_assert!(
                self.added.last().is_none_or(|last| last.depth <= depth),
                "namespace scope depths must not decrease"
            );
            let uri = self.intern(namespace, capabilities)?;
            match self.prefixes.get_mut(prefix.as_str()) {
                Some(bindings) => bindings.push((depth, uri)),
                None => {
                    self.prefixes.insert(prefix.clone(), vec![(depth, uri)]);
                },
            }
            self.added.push(Added {
                depth,
                prefix: prefix.clone(),
                key: packed(prefix),
                uri,
            });
        }
        Ok(())
    }

    /// The identity of `text` in scope: an existing one with the same bytes,
    /// or a new one. The URI is hashed once and compared only with URIs in
    /// scope with the same keyed hash.
    fn intern(&mut self, text: &str, capabilities: &Capabilities) -> Result<UriId, Error> {
        #[cfg(test)]
        counter::uri_bytes(text.len());
        let key = self.keys.hash_one(text);
        let mut candidate = self.uri_index.get(&key).copied().unwrap_or(NONE);
        while candidate != NONE {
            let entry = &mut self.uris[candidate];
            #[cfg(test)]
            counter::uri_bytes(text.len());
            if entry.text == text {
                entry.refs += 1;
                return Ok(candidate);
            }
            candidate = entry.hidden;
        }
        let allocation = |source| Error::Allocation {
            resource: "MCE namespace identities",
            source,
        };
        self.uris.try_reserve(1).map_err(allocation)?;
        self.uri_index.try_reserve(1).map_err(allocation)?;
        let mut owned = String::new();
        owned.try_reserve_exact(text.len()).map_err(allocation)?;
        owned.push_str(text);
        let id = self.uris.len();
        let hidden = self.uri_index.insert(key, id).unwrap_or(NONE);
        self.uris.push(UriEntry {
            text: owned,
            key,
            hidden,
            refs: 1,
            facts: Facts::of(text, capabilities),
        });
        Ok(id)
    }

    /// Release one binding of `id`, and every identity at the top of the
    /// stack that no binding holds any more.
    fn release(&mut self, id: UriId) {
        if let Some(entry) = self.uris.get_mut(id) {
            entry.refs = entry.refs.saturating_sub(1);
        }
        debug_assert!(
            self.uris.get(id).is_none_or(|entry| entry.refs > 0) || id + 1 == self.uris.len(),
            "namespace identities must be released in stack order"
        );
        while let Some(top) = self.uris.last()
            && top.refs == 0
        {
            let Some(entry) = self.uris.pop() else {
                break;
            };
            if entry.hidden == NONE {
                self.uri_index.remove(&entry.key);
            } else {
                self.uri_index.insert(entry.key, entry.hidden);
            }
        }
    }

    /// Remove every binding declared at `depth` or deeper: those of an element
    /// that has closed and of anything it left open.
    #[inline]
    pub(super) fn truncate(&mut self, depth: usize) {
        if self.added.last().is_some_and(|last| last.depth >= depth) {
            self.remove_from(depth);
        }
    }

    #[cold]
    fn remove_from(&mut self, depth: usize) {
        while self.added.last().is_some_and(|last| last.depth >= depth) {
            let Some(binding) = self.added.pop() else {
                break;
            };
            if let Some(bindings) = self.prefixes.get_mut(binding.prefix.as_str()) {
                debug_assert!(
                    bindings
                        .last()
                        .is_some_and(|(last, _)| *last == binding.depth),
                    "a prefix's latest binding must be the latest one added"
                );
                bindings.pop();
                if bindings.is_empty() {
                    self.prefixes.remove(binding.prefix.as_str());
                }
            }
            self.release(binding.uri);
        }
    }
}

/// The identity hasher of [`Index`], whose keys are already keyed hashes.
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

/// Whether two of one element's declarations bind the same prefix.
///
/// Sorting borrowed prefixes keeps the check `O(n log n)` in the number of
/// declarations; a pairwise comparison would be quadratic in a count the input
/// chooses.
///
/// # Errors
///
/// [`Error::Allocation`] when the sort buffer cannot be reserved.
pub(super) fn has_duplicate_prefix(local: &[(String, String)]) -> Result<bool, Error> {
    if local.len() < 2 {
        return Ok(false);
    }
    let mut prefixes = sorted_prefixes(local)?;
    prefixes.dedup();
    Ok(prefixes.len() != local.len())
}

/// One element's declared prefixes, sorted for binary search.
///
/// # Errors
///
/// [`Error::Allocation`] when the buffer cannot be reserved.
pub(super) fn sorted_prefixes(local: &[(String, String)]) -> Result<Vec<&str>, Error> {
    let mut prefixes = Vec::new();
    prefixes
        .try_reserve_exact(local.len())
        .map_err(|source| Error::Allocation {
            resource: "MCE namespace declaration check",
            source,
        })?;
    prefixes.extend(local.iter().map(|(prefix, _)| prefix.as_str()));
    prefixes.sort_unstable();
    Ok(prefixes)
}

/// Index operations and walked declarations, counted in tests so that they
/// can bound the work a hostile document costs without measuring time.
#[cfg(test)]
pub(super) mod counter {
    use core::cell::Cell;

    thread_local! {
        static LOOKUPS: Cell<usize> = const { Cell::new(0) };
        static URI_BYTES: Cell<usize> = const { Cell::new(0) };
    }

    /// Count `bytes` of namespace URI hashed or compared.
    pub(super) fn uri_bytes(bytes: usize) {
        URI_BYTES.with(|count| count.set(count.get() + bytes));
    }

    /// Run `work` and return its result with the namespace URI bytes the
    /// scope hashed or compared.
    pub(in crate::mce) fn counted_uri_bytes<T>(work: impl FnOnce() -> T) -> (T, usize) {
        URI_BYTES.with(|count| count.set(0));
        let result = work();
        (result, URI_BYTES.with(Cell::get))
    }

    pub(super) fn lookup() {
        LOOKUPS.with(|count| count.set(count.get() + 1));
    }

    /// Count `steps` declarations a production lookup walked.
    pub(in crate::mce) fn steps(steps: usize) {
        LOOKUPS.with(|count| count.set(count.get() + steps));
    }

    /// Run `work` and return its result with the index lookups it made and
    /// the declarations its lookups walked.
    pub(in crate::mce) fn counted<T>(work: impl FnOnce() -> T) -> (T, usize) {
        LOOKUPS.with(|count| count.set(0));
        let result = work();
        (result, LOOKUPS.with(Cell::get))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn declarations(pairs: &[(&str, &str)]) -> Vec<(String, String)> {
        pairs
            .iter()
            .map(|(prefix, namespace)| ((*prefix).to_owned(), (*namespace).to_owned()))
            .collect()
    }

    fn scope() -> Scope {
        Scope::new(&Capabilities::ooxml_baseline())
    }

    #[test]
    fn nested_bindings_shadow_and_truncation_restores_them() {
        let mut scope = scope();
        let caps = Capabilities::ooxml_baseline();
        scope
            .push(1, &declarations(&[("a", "urn:a1"), ("", "urn:d1")]), &caps)
            .unwrap();
        scope
            .push(2, &declarations(&[("b", "urn:b2")]), &caps)
            .unwrap();
        scope
            .push(3, &declarations(&[("a", "urn:a3"), ("", "")]), &caps)
            .unwrap();
        assert_eq!(scope.get("a"), Some("urn:a3"));
        assert_eq!(scope.get(""), Some(""));
        assert_eq!(scope.get("b"), Some("urn:b2"));
        assert_eq!(scope.get("c"), None);
        assert_eq!(scope.get("xml"), Some(XML_NS));
        assert_eq!(scope.get_outside("a", 3), Some("urn:a1"));
        assert_eq!(scope.get_outside("b", 3), Some("urn:b2"));
        assert_eq!(scope.get_outside("b", 2), None);
        assert!(scope.declared_at("a", 3));
        assert!(!scope.declared_at("b", 3));

        scope.truncate(3);
        assert_eq!(scope.get("a"), Some("urn:a1"));
        assert_eq!(scope.get(""), Some("urn:d1"));
        scope.truncate(2);
        assert_eq!(scope.get("b"), None);
        scope.truncate(1);
        assert_eq!(scope.get("a"), None);
        assert!(scope.prefixes.is_empty());
        assert!(scope.added.is_empty());
        assert!(scope.uris.is_empty());
        assert!(scope.uri_index.is_empty());
    }

    #[test]
    fn truncating_an_element_without_declarations_removes_nothing() {
        let mut scope = scope();
        let caps = Capabilities::ooxml_baseline();
        scope
            .push(1, &declarations(&[("a", "urn:a")]), &caps)
            .unwrap();
        scope.truncate(2);
        assert_eq!(scope.get("a"), Some("urn:a"));
    }

    #[test]
    fn equal_uris_share_one_identity_while_any_binding_holds_it() {
        let mut scope = scope();
        let caps = Capabilities::ooxml_baseline();
        scope
            .push(
                1,
                &declarations(&[("a", "urn:same"), ("b", NAMESPACE)]),
                &caps,
            )
            .unwrap();
        scope
            .push(
                2,
                &declarations(&[("c", "urn:same"), ("d", "urn:other")]),
                &caps,
            )
            .unwrap();
        let a = scope.resolve("a").unwrap().id;
        let c = scope.resolve("c").unwrap().id;
        let d = scope.resolve("d").unwrap().id;
        assert_eq!(a, c);
        assert_ne!(a, d);
        assert!(scope.is_mce(scope.resolve("b").unwrap().id));
        assert!(!scope.is_mce(a));
        assert!(scope.understood(XML_URI));
        assert!(!scope.understood(a));
        assert_eq!(scope.uris.len(), 3);

        scope.truncate(2);
        // `urn:same` is still bound at depth 1; `urn:other` is released.
        assert_eq!(scope.uris.len(), 2);
        assert_eq!(scope.resolve("a").unwrap().id, a);
        scope
            .push(2, &declarations(&[("e", "urn:third")]), &caps)
            .unwrap();
        assert_eq!(scope.resolve("e").unwrap().text, "urn:third");
        scope.truncate(1);
        assert!(scope.uris.is_empty());
        assert!(scope.uri_index.is_empty());
    }

    #[test]
    fn packed_prefixes_are_equal_exactly_when_the_prefixes_are() {
        let prefixes = [
            "", "a", "w", "x14ac", "xr", "xr2", "abcdefg", "abcdefgh", "abcdefgi",
        ];
        for left in prefixes {
            for right in prefixes {
                let short = left.len() <= 7 && right.len() <= 7;
                if short {
                    assert_eq!(
                        packed(left) == packed(right),
                        left == right,
                        "{left} {right}"
                    );
                }
            }
        }
        assert_eq!(packed("abcdefgh"), LONG_PREFIX);
        let mut scope = scope();
        let caps = Capabilities::ooxml_baseline();
        scope
            .push(
                1,
                &declarations(&[
                    ("abcdefgh", "urn:long-1"),
                    ("abcdefgi", "urn:long-2"),
                    ("a", "urn:a"),
                ]),
                &caps,
            )
            .unwrap();
        assert_eq!(scope.get("abcdefgh"), Some("urn:long-1"));
        assert_eq!(scope.get("abcdefgi"), Some("urn:long-2"));
        assert_eq!(scope.get("a"), Some("urn:a"));
        assert_eq!(scope.get("abcdefgj"), None);
    }

    #[test]
    fn lookups_past_the_walked_window_use_the_index() {
        let mut scope = scope();
        let caps = Capabilities::ooxml_baseline();
        scope
            .push(1, &declarations(&[("far", "urn:far")]), &caps)
            .unwrap();
        let many: Vec<(String, String)> = (0..2 * WALKED_DECLARATIONS)
            .map(|index| (format!("p{index}"), format!("urn:{index}")))
            .collect();
        scope.push(2, &many, &caps).unwrap();
        let (found, operations) = counter::counted(|| scope.resolve("far").map(|uri| uri.text));
        assert_eq!(found, Some("urn:far"));
        assert_eq!(operations, WALKED_DECLARATIONS + 1);
        let (missing, operations) =
            counter::counted(|| scope.resolve("nowhere").map(|uri| uri.text));
        assert_eq!(missing, None);
        assert_eq!(operations, WALKED_DECLARATIONS + 1);
    }

    #[test]
    fn duplicate_prefixes_are_found_in_any_position() {
        assert!(!has_duplicate_prefix(&declarations(&[])).unwrap());
        assert!(!has_duplicate_prefix(&declarations(&[("a", "1")])).unwrap());
        assert!(!has_duplicate_prefix(&declarations(&[("a", "1"), ("b", "1")])).unwrap());
        assert!(has_duplicate_prefix(&declarations(&[("a", "1"), ("a", "2")])).unwrap());
        assert!(
            has_duplicate_prefix(&declarations(&[
                ("", "1"),
                ("b", "1"),
                ("c", "1"),
                ("", "2")
            ]))
            .unwrap()
        );
    }
}
