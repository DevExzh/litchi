//! Persistent, bounded namespace environments for source and typed XML trees.

use super::invalid;
use litchi_core::sheet::Result;
use std::{
    collections::{HashMap, HashSet},
    sync::Arc,
};

const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

/// Resource bounds for one namespace environment and its XML nesting path.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct NamespaceLimits {
    pub(super) max_depth: usize,
    pub(super) max_bindings: usize,
    pub(super) max_bytes: usize,
}

impl NamespaceLimits {
    pub(super) const fn new(max_depth: usize, max_bindings: usize, max_bytes: usize) -> Self {
        Self {
            max_depth,
            max_bindings,
            max_bytes,
        }
    }
}

/// One decoded namespace declaration on the next XML element.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct NamespaceDecl<'a> {
    pub(super) prefix: &'a str,
    pub(super) uri: &'a str,
}

impl<'a> NamespaceDecl<'a> {
    pub(super) const fn new(prefix: &'a str, uri: &'a str) -> Self {
        Self { prefix, uri }
    }
}

/// A persistent namespace environment with bounded local deltas.
///
/// Cloning a context shares its frame. A child with no declarations also
/// shares the frame and only advances the logical XML depth. A child with
/// declarations stores one local frame linked to the inherited environment;
/// inherited bindings are never copied into that frame.
#[derive(Clone, Debug)]
pub(super) struct NamespaceContext {
    frame: Arc<Frame>,
    depth: usize,
    limits: NamespaceLimits,
    xml_uri: Arc<str>,
    xmlns_uri: Arc<str>,
}

#[derive(Debug)]
struct Frame {
    parent: Option<Arc<Frame>>,
    local: Arc<[Binding]>,
    prefix_index: Arc<HashMap<Arc<str>, Arc<str>>>,
    layers: usize,
    bindings: usize,
    bytes: usize,
}

#[derive(Debug)]
struct Binding {
    prefix: Arc<str>,
    uri: Arc<str>,
}

impl NamespaceContext {
    pub(super) fn root(limits: NamespaceLimits) -> Self {
        Self {
            frame: Arc::new(Frame {
                parent: None,
                local: Arc::from([]),
                prefix_index: Arc::new(HashMap::new()),
                layers: 0,
                bindings: 0,
                bytes: 0,
            }),
            depth: 0,
            limits,
            xml_uri: Arc::from(XML_NAMESPACE),
            xmlns_uri: Arc::from(XMLNS_NAMESPACE),
        }
    }

    /// Return a child context after applying that element's decoded bindings.
    pub(super) fn child<'a, I>(&self, declarations: I) -> Result<Self>
    where
        I: IntoIterator<Item = NamespaceDecl<'a>>,
    {
        let depth = self
            .depth
            .checked_add(1)
            .ok_or_else(|| invalid("connections namespace depth overflowed"))?;
        if depth > self.limits.max_depth {
            return Err(invalid("connections namespace depth limit exceeded"));
        }

        let mut local = Vec::new();
        let mut local_prefixes = HashSet::new();
        let mut local_bytes = 0usize;
        for declaration in declarations {
            if (declaration.prefix == "xml" && declaration.uri != XML_NAMESPACE)
                || (declaration.prefix == "xmlns" && declaration.uri != XMLNS_NAMESPACE)
                || (declaration.uri == XML_NAMESPACE && declaration.prefix != "xml")
                || (declaration.uri == XMLNS_NAMESPACE && declaration.prefix != "xmlns")
            {
                return Err(invalid("invalid reserved XML namespace binding"));
            }
            let bindings = self
                .frame
                .bindings
                .checked_add(local.len())
                .and_then(|bindings| bindings.checked_add(1))
                .ok_or_else(|| invalid("connections namespace binding count overflowed"))?;
            if bindings > self.limits.max_bindings {
                return Err(invalid("connections namespace binding limit exceeded"));
            }
            local_prefixes
                .try_reserve(1)
                .map_err(|_| invalid("connections namespace binding allocation failed"))?;
            if !local_prefixes.insert(declaration.prefix) {
                return Err(invalid("duplicate namespace declaration"));
            }
            let declaration_bytes = declaration
                .prefix
                .len()
                .checked_add(declaration.uri.len())
                .ok_or_else(|| invalid("connections namespace metadata size overflowed"))?;
            local_bytes = local_bytes
                .checked_add(declaration_bytes)
                .ok_or_else(|| invalid("connections namespace metadata size overflowed"))?;
            let bytes = self
                .frame
                .bytes
                .checked_add(local_bytes)
                .ok_or_else(|| invalid("connections namespace metadata size overflowed"))?;
            if bytes > self.limits.max_bytes {
                return Err(invalid("connections namespace metadata limit exceeded"));
            }
            local
                .try_reserve(1)
                .map_err(|_| invalid("connections namespace binding allocation failed"))?;
            let uri = if declaration.prefix == "xml" {
                Arc::clone(&self.xml_uri)
            } else if declaration.prefix == "xmlns" {
                Arc::clone(&self.xmlns_uri)
            } else {
                Arc::from(declaration.uri)
            };
            let prefix = Arc::from(declaration.prefix);
            local.push(Binding { prefix, uri });
        }

        if local.is_empty() {
            return Ok(Self {
                frame: Arc::clone(&self.frame),
                depth,
                limits: self.limits,
                xml_uri: Arc::clone(&self.xml_uri),
                xmlns_uri: Arc::clone(&self.xmlns_uri),
            });
        }

        let layers = self
            .frame
            .layers
            .checked_add(1)
            .ok_or_else(|| invalid("connections namespace lookup depth overflowed"))?;
        if layers > self.limits.max_depth {
            return Err(invalid("connections namespace lookup depth limit exceeded"));
        }
        let local_bindings = local.len();
        let bindings = self
            .frame
            .bindings
            .checked_add(local_bindings)
            .ok_or_else(|| invalid("connections namespace binding count overflowed"))?;
        let bytes = self
            .frame
            .bytes
            .checked_add(local_bytes)
            .ok_or_else(|| invalid("connections namespace metadata size overflowed"))?;
        let mut prefix_index = HashMap::new();
        prefix_index
            .try_reserve(local.len())
            .map_err(|_| invalid("connections namespace binding allocation failed"))?;
        for binding in &local {
            prefix_index.insert(Arc::clone(&binding.prefix), Arc::clone(&binding.uri));
        }
        Ok(Self {
            frame: Arc::new(Frame {
                parent: Some(Arc::clone(&self.frame)),
                local: Arc::from(local.into_boxed_slice()),
                prefix_index: Arc::new(prefix_index),
                layers,
                bindings,
                bytes,
            }),
            depth,
            limits: self.limits,
            xml_uri: Arc::clone(&self.xml_uri),
            xmlns_uri: Arc::clone(&self.xmlns_uri),
        })
    }

    /// Resolve a prefix without allocating a URI string.
    ///
    /// The returned `Arc<str>` can be cloned into a node or attribute and
    /// continues to share the URI storage with every context that inherited
    /// the same declaration.
    pub(super) fn resolve_uri(&self, prefix: &str) -> Option<&Arc<str>> {
        if prefix == "xml" {
            return Some(&self.xml_uri);
        }
        let mut frame = Some(self.frame.as_ref());
        while let Some(current) = frame {
            if let Some(uri) = current.prefix_index.get(prefix) {
                return Some(uri);
            }
            frame = current.parent.as_deref();
        }
        None
    }

    /// Return effective declarations in source order, with the nearest
    /// declaration winning when a prefix is shadowed.
    pub(super) fn effective_bindings(&self) -> Result<Vec<(&str, &str)>> {
        let mut bindings = Vec::new();
        let mut seen = HashSet::new();
        let mut frame = Some(self.frame.as_ref());
        while let Some(current) = frame {
            for binding in current.local.iter().rev() {
                if seen.contains(binding.prefix.as_ref()) {
                    continue;
                }
                seen.try_reserve(1)
                    .map_err(|_| invalid("connections namespace binding allocation failed"))?;
                bindings
                    .try_reserve(1)
                    .map_err(|_| invalid("connections namespace binding allocation failed"))?;
                seen.insert(binding.prefix.as_ref());
                bindings.push((binding.prefix.as_ref(), binding.uri.as_ref()));
            }
            frame = current.parent.as_deref();
        }
        bindings.reverse();
        Ok(bindings)
    }

    #[cfg(test)]
    pub(super) fn depth(&self) -> usize {
        self.depth
    }

    pub(super) fn xmlns_uri(&self) -> &Arc<str> {
        &self.xmlns_uri
    }
}

impl Drop for Frame {
    fn drop(&mut self) {
        let mut parent = self.parent.take();
        while let Some(frame) = parent.take() {
            match Arc::into_inner(frame) {
                Some(mut frame) => parent = frame.parent.take(),
                None => break,
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;

    fn limits() -> NamespaceLimits {
        NamespaceLimits::new(32, 64, 4096)
    }

    #[test]
    fn empty_children_share_frames_and_uri_storage() {
        let root = NamespaceContext::root(limits());
        let parent = root
            .child([
                NamespaceDecl::new("p", "urn:shared"),
                NamespaceDecl::new("", "urn:default"),
            ])
            .unwrap();
        let child = parent.child([]).unwrap();

        assert!(Arc::ptr_eq(&parent.frame, &child.frame));
        assert!(Arc::ptr_eq(
            parent.resolve_uri("p").unwrap(),
            child.resolve_uri("p").unwrap()
        ));
        assert_eq!(child.resolve_uri("").unwrap().as_ref(), "urn:default");
        assert_eq!(child.depth(), 2);
    }

    #[test]
    fn many_children_share_large_inherited_uri_storage() {
        let uri = format!("urn:large:{}", "x".repeat(64 * 1024));
        let root = NamespaceContext::root(NamespaceLimits::new(2, 2, 128 * 1024));
        let parent = root.child([NamespaceDecl::new("p", &uri)]).unwrap();
        let inherited = Arc::clone(parent.resolve_uri("p").unwrap());
        let mut children = Vec::new();
        children.try_reserve(1024).unwrap();
        for _ in 0..1024 {
            let child = parent.child([]).unwrap();
            assert!(Arc::ptr_eq(
                &inherited,
                child.resolve_uri("p").expect("inherited URI")
            ));
            children.push(child);
        }
        assert_eq!(inherited.as_ref(), uri);
    }

    #[test]
    fn shadowing_keeps_outer_scope_and_inherited_uri_sharing() {
        let root = NamespaceContext::root(limits());
        let outer = root.child([NamespaceDecl::new("p", "urn:shared")]).unwrap();
        let inner = outer.child([NamespaceDecl::new("p", "urn:inner")]).unwrap();
        let repeated = inner
            .child([NamespaceDecl::new("q", "urn:shared")])
            .unwrap();

        assert_eq!(outer.resolve_uri("p").unwrap().as_ref(), "urn:shared");
        assert_eq!(inner.resolve_uri("p").unwrap().as_ref(), "urn:inner");
        assert_eq!(repeated.resolve_uri("q").unwrap().as_ref(), "urn:shared");
        assert!(Arc::ptr_eq(
            inner.resolve_uri("p").unwrap(),
            repeated.resolve_uri("p").unwrap()
        ));
    }

    #[test]
    fn default_empty_and_xml_namespaces_have_distinct_resolution() {
        let root = NamespaceContext::root(limits());
        let child = root
            .child([
                NamespaceDecl::new("", ""),
                NamespaceDecl::new("xml", XML_NAMESPACE),
            ])
            .unwrap();

        assert_eq!(child.resolve_uri("").unwrap().as_ref(), "");
        assert_eq!(child.resolve_uri("xml").unwrap().as_ref(), XML_NAMESPACE);
        assert!(child.resolve_uri("missing").is_none());
    }

    #[test]
    fn depth_binding_and_metadata_bounds_are_cumulative() {
        let root = NamespaceContext::root(NamespaceLimits::new(1, 1, 4));
        let child = root.child([NamespaceDecl::new("p", "u")]).unwrap();
        assert!(child.child([]).is_err());
        assert!(
            root.child([NamespaceDecl::new("p", "u"), NamespaceDecl::new("q", "v"),])
                .is_err()
        );
        assert!(root.child([NamespaceDecl::new("long", "uri")]).is_err());
    }

    #[test]
    fn large_distinct_declaration_list_is_bounded_without_uri_intern_scans() {
        let count = 16_384;
        let mut declarations = Vec::new();
        declarations.try_reserve(count).unwrap();
        let mut metadata_bytes = 0usize;
        for index in 0..count {
            let prefix = format!("p{index}");
            let uri = format!("urn:distinct:{index}");
            metadata_bytes += prefix.len() + uri.len();
            declarations.push((prefix, uri));
        }
        let root = NamespaceContext::root(NamespaceLimits::new(1, count, metadata_bytes));
        let context = root
            .child(
                declarations
                    .iter()
                    .map(|(prefix, uri)| NamespaceDecl::new(prefix, uri)),
            )
            .unwrap();
        assert_eq!(
            context.resolve_uri("p0").unwrap().as_ref(),
            "urn:distinct:0"
        );
        assert_eq!(
            context.resolve_uri("p16383").unwrap().as_ref(),
            "urn:distinct:16383"
        );
    }

    #[test]
    fn deeply_linked_contexts_drop_without_recursive_parent_unwinding() {
        let max_depth = 4096;
        let mut context =
            NamespaceContext::root(NamespaceLimits::new(max_depth, max_depth, max_depth * 32));
        for index in 0..max_depth {
            let uri = format!("urn:namespace:{index}");
            context = context.child([NamespaceDecl::new("p", &uri)]).unwrap();
        }
        assert_eq!(context.depth(), max_depth);
        drop(context);
    }
}
