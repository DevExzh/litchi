//! The checked walk shared by the direct-property readers.
//!
//! One pass over an element's XML finds its direct property containers (for
//! example `w:rPr` under `w:r`) and hands every WordprocessingML child of a
//! container, and every WordprocessingML child of such a child, to a
//! visitor. Everything else is skipped by depth: run content, nested runs in
//! text boxes, nested tables, and revision snapshots such as `w:rPrChange`,
//! whose own `w:rPr` sits one level deeper than the run's. A nested property
//! container therefore can never be mistaken for the element's own.
//!
//! Names are resolved as the other fragment readers of this crate resolve
//! them (`parse_run_underline`): a bound name counts when its namespace is the
//! transitional or strict WordprocessingML namespace; an undeclared prefix,
//! or no prefix without a default namespace, counts only when the fragment's
//! own root is written that way, because a fragment cut from a part does not
//! carry the declarations of the part's root. Names are resolved only where a
//! property can sit; an undeclared prefix is compared in place rather than
//! copied.

use std::borrow::Cow;

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, QName};
use quick_xml::reader::NsReader;
use smallvec::SmallVec;

use crate::error::{Error, Result};
use crate::namespace::{STRICT_WORDPROCESSINGML_NAMESPACE, WORDPROCESSINGML_NAMESPACE};

/// Deepest element nesting accepted in one fragment, as the shared element
/// scanner accepts (`namespace::MAX_SCAN_DEPTH`).
const MAX_DEPTH: usize = 128;
/// Most elements visited in one fragment, as the shared element scanner.
const MAX_ELEMENTS: usize = 1_000_000;

/// How a name resolved; an undeclared prefix is not copied.
#[derive(Clone, Copy)]
pub(super) enum Resolution<'r> {
    Bound(&'r [u8]),
    Undeclared,
    Unbound,
}

/// How the fragment's root is written.
pub(super) enum Root {
    /// Bound to a namespace: only bound WordprocessingML names count.
    Bound,
    /// An undeclared prefix: names with that prefix count.
    Prefix(SmallVec<[u8; 8]>),
    /// No prefix and no default namespace: such names count.
    Unprefixed,
}

impl Root {
    pub(super) fn is_word(&self, prefix: Option<&[u8]>, resolution: Resolution<'_>) -> bool {
        match resolution {
            Resolution::Bound(namespace) => {
                namespace == WORDPROCESSINGML_NAMESPACE
                    || namespace == STRICT_WORDPROCESSINGML_NAMESPACE
            },
            Resolution::Undeclared => {
                matches!((self, prefix), (Self::Prefix(root), Some(prefix)) if root.as_slice() == prefix)
            },
            Resolution::Unbound => matches!(self, Self::Unprefixed),
        }
    }
}

/// Resolve `prefix` against the declarations in scope. `element` names use
/// the default namespace for no prefix; attribute names do not.
pub(super) fn resolve<'r>(
    resolver: &'r NamespaceResolver,
    prefix: Option<&[u8]>,
    element: bool,
) -> Resolution<'r> {
    let wanted = match prefix {
        Some(prefix) => Some(prefix),
        None if element => None,
        None => return Resolution::Unbound,
    };
    let declared = resolver
        .bindings()
        .find_map(|(declaration, Namespace(namespace))| {
            let matches = match (declaration, wanted) {
                (PrefixDeclaration::Named(name), Some(prefix)) => name == prefix,
                (PrefixDeclaration::Default, None) => true,
                (PrefixDeclaration::Named(_), None) | (PrefixDeclaration::Default, Some(_)) => {
                    false
                },
            };
            matches.then_some(namespace)
        })
        .filter(|namespace| !namespace.is_empty());
    match (declared, prefix) {
        (Some(namespace), _) => Resolution::Bound(namespace),
        (None, Some(_)) => Resolution::Undeclared,
        (None, None) => Resolution::Unbound,
    }
}

/// The name, split into local name and prefix bytes.
pub(super) fn split<'n>(name: QName<'n>) -> (&'n [u8], Option<&'n [u8]>) {
    let (local, prefix) = name.decompose();
    (local.into_inner(), prefix.map(|prefix| prefix.into_inner()))
}

/// Where a visited element sits.
#[derive(Clone, Copy)]
pub(super) enum Site<'p> {
    /// A direct child of the container `containers[index]`.
    Property(usize),
    /// A direct child of a WordprocessingML property element; the value is
    /// that property element's local name.
    Nested(&'p [u8]),
}

/// Name resolution and decoding for the element being visited.
pub(super) struct Context<'a> {
    resolver: &'a NamespaceResolver,
    decoder: Decoder,
    root: &'a Root,
}

impl Context<'_> {
    /// Visit every WordprocessingML attribute of `element` with its local
    /// name and its decoded, normalized value.
    ///
    /// # Errors
    ///
    /// Malformed or duplicated attributes and undecodable values are XML
    /// errors.
    pub(super) fn attributes<'e>(
        &self,
        element: &'e BytesStart<'_>,
        mut visit: impl FnMut(&[u8], Cow<'e, str>),
    ) -> Result<()> {
        let mut seen = std::collections::BTreeSet::new();
        for attribute in element.checked_attributes() {
            let attribute =
                attribute.map_err(|_| Error::Xml("malformed direct-property XML".into()))?;
            let (local, prefix) = split(attribute.key);
            if !self
                .root
                .is_word(prefix, resolve(self.resolver, prefix, false))
            {
                continue;
            }
            if !seen.insert(local) {
                return Err(Error::Xml("duplicate Word property attribute".into()));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, self.decoder)
                .map_err(|_| Error::Xml("malformed direct-property XML".into()))?;
            visit(local, value);
        }
        Ok(())
    }
}

/// Walk one `root` element fragment and visit the WordprocessingML children
/// of its direct `containers` children, and their WordprocessingML children.
///
/// `what` names the element in error messages.
///
/// # Errors
///
/// XML that is not well-formed, a root that is not a WordprocessingML
/// `root`, a second root, a container that appears twice, nesting deeper
/// than 128 elements, more than a million elements, and every error the
/// visitor returns.
pub(super) fn walk<'i>(
    xml: &'i [u8],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
    root: &[u8],
    containers: &[&[u8]],
    what: &str,
    mut visit: impl FnMut(&Context<'_>, Site<'_>, &[u8], &BytesStart<'i>) -> Result<()>,
) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    for (prefix, namespace) in inherited {
        let prefix = prefix
            .as_deref()
            .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|_| Error::Xml("invalid inherited property namespace".into()))?;
    }
    let mut fragment = None::<Root>;
    let mut depth = 0usize;
    let mut elements = 0usize;
    let mut seen = 0u32;
    let mut container = None::<usize>;
    let mut property = None::<SmallVec<[u8; 32]>>;

    loop {
        let event = reader
            .read_event()
            .map_err(|_| Error::Xml("malformed direct-property XML".into()))?;
        let (element, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                match depth {
                    2 => container = None,
                    3 => property = None,
                    _ => {},
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat(format!("invalid Word {what} XML nesting"))
                })?;
                continue;
            },
            Event::Eof if depth != 0 => {
                return Err(Error::InvalidFormat(format!(
                    "unterminated Word {what} XML"
                )));
            },
            Event::Eof => break,
            Event::DocType(_) => {
                return Err(Error::Xml("DOCTYPE is forbidden in property XML".into()));
            },
            Event::Text(ref text)
                if depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Error::Xml("text outside the property XML root".into()));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::Xml("content outside the property XML root".into()));
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => continue,
        };

        audit_attributes(&element, reader.resolver(), reader.decoder())?;
        elements += 1;
        if elements > MAX_ELEMENTS {
            return Err(Error::InvalidFormat(format!(
                "Word {what} XML exceeds {MAX_ELEMENTS} elements"
            )));
        }
        let level = depth + 1;
        if level > MAX_DEPTH {
            return Err(Error::InvalidFormat(format!(
                "Word {what} XML nesting exceeds the {MAX_DEPTH} depth limit"
            )));
        }
        if !empty {
            depth = level;
        }

        let (local, prefix) = split(element.name());
        match level {
            1 => {
                if fragment.is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "Word {what} XML has more than one root"
                    )));
                }
                let resolution = resolve(reader.resolver(), prefix, true);
                let written = match resolution {
                    Resolution::Bound(_) => Root::Bound,
                    Resolution::Undeclared => {
                        Root::Prefix(SmallVec::from_slice(prefix.unwrap_or_default()))
                    },
                    Resolution::Unbound => Root::Unprefixed,
                };
                if local != root || !written.is_word(prefix, resolution) {
                    return Err(Error::InvalidFormat(format!(
                        "Word {what} XML has an invalid root"
                    )));
                }
                fragment = Some(written);
            },
            2 => {
                let Some(index) = containers.iter().position(|name| *name == local) else {
                    continue;
                };
                let Some(written) = fragment.as_ref() else {
                    continue;
                };
                if !written.is_word(prefix, resolve(reader.resolver(), prefix, true)) {
                    continue;
                }
                let bit = 1u32 << index;
                if seen & bit != 0 {
                    return Err(Error::InvalidFormat(format!(
                        "Word {what} has a duplicate {} container",
                        String::from_utf8_lossy(local)
                    )));
                }
                seen |= bit;
                if !empty {
                    container = Some(index);
                }
            },
            3 | 4 => {
                let (Some(index), Some(written)) = (container, fragment.as_ref()) else {
                    continue;
                };
                if level == 4 && property.is_none() {
                    continue;
                }
                let resolver = reader.resolver();
                if !written.is_word(prefix, resolve(resolver, prefix, true)) {
                    continue;
                }
                let context = Context {
                    resolver,
                    decoder: reader.decoder(),
                    root: written,
                };
                if level == 3 {
                    visit(&context, Site::Property(index), local, &element)?;
                    if !empty {
                        property = Some(SmallVec::from_slice(local));
                    }
                } else if let Some(parent) = property.as_deref() {
                    visit(&context, Site::Nested(parent), local, &element)?;
                }
            },
            _ => {},
        }
    }

    if fragment.is_none() {
        return Err(Error::InvalidFormat(format!("Word {what} XML has no root")));
    }
    Ok(())
}

pub(crate) fn audit_attributes(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
) -> Result<()> {
    let mut seen = std::collections::BTreeSet::new();
    for attribute in element.checked_attributes() {
        let attribute =
            attribute.map_err(|_| Error::Xml("malformed property XML attribute".into()))?;
        let (local, prefix) = split(attribute.key);
        let key = match resolve(resolver, prefix, false) {
            Resolution::Bound(namespace) => (0, namespace, local),
            Resolution::Unbound => (1, &b""[..], local),
            Resolution::Undeclared => (2, prefix.unwrap_or_default(), local),
        };
        if !seen.insert(key) {
            return Err(Error::Xml(
                "duplicate expanded property XML attribute".into(),
            ));
        }
        attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|_| Error::Xml("malformed property XML attribute value".into()))?;
    }
    Ok(())
}
