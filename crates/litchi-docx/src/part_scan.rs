//! A shallow, located element scan of one WordprocessingML part (record
//! 0319).
//!
//! The style, numbering, note and story catalogs read their parts through
//! this scan: every element down to a chosen depth, in document order, with
//! its byte range in the part, its local name and prefix, its attributes by
//! local name, and its parent. Deeper content is checked for well-formedness
//! and skipped. Ranges address the raw part, so a catalog edit splices the
//! part's own bytes and leaves everything else exactly as authored.

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::first_wins;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

use crate::error::{Error, Result};
use crate::properties;

/// The deepest nesting the scan accepts.
const MAX_DEPTH: usize = 256;
/// The most elements one scanned part may hold.
const MAX_ELEMENTS: usize = 1_000_000;

/// Namespace bindings in scope: `(prefix, namespace)`, `None` for the
/// default namespace.
pub(crate) type Bindings = Vec<(Option<Vec<u8>>, Vec<u8>)>;

/// A property container wrapped for the property readers, with its scope.
pub(crate) type Wrapped = (Vec<u8>, Bindings);

/// One recorded element.
#[derive(Debug, Clone)]
pub(crate) struct Element {
    /// Offset of the start tag's `<`.
    pub(crate) start: usize,
    /// Offset just past the start tag (`>` or `/>`).
    pub(crate) tag_end: usize,
    /// Offset of the end tag's `<`, or `end` for an empty element.
    pub(crate) close: usize,
    /// Offset just past the element.
    pub(crate) end: usize,
    /// Whether the element is written `<x/>`.
    pub(crate) empty: bool,
    /// Depth: the root is 0.
    pub(crate) depth: usize,
    /// The local name.
    pub(crate) local: Box<[u8]>,
    /// The prefix, empty when unprefixed.
    pub(crate) prefix: Box<[u8]>,
    /// The attributes by local name, decoded; `xmlns` declarations are kept
    /// separately in `declarations`.
    pub(crate) attributes: Vec<(Box<[u8]>, String)>,
    /// The namespace declarations of the start tag: `(prefix, namespace)`,
    /// `None` for the default namespace.
    pub(crate) declarations: Vec<(Option<Vec<u8>>, Vec<u8>)>,
    /// The parent's index, `None` for the root.
    pub(crate) parent: Option<usize>,
}

impl Element {
    /// The value of the attribute with local name `local`.
    pub(crate) fn attribute(&self, local: &[u8]) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(name, _)| name.as_ref() == local)
            .map(|(_, value)| value.as_str())
    }

    /// Whether the element has local name `local`.
    pub(crate) fn is(&self, local: &[u8]) -> bool {
        self.local.as_ref() == local
    }
}

/// The recorded elements of one part, in document order.
#[derive(Debug, Clone)]
pub(crate) struct Tree {
    pub(crate) elements: Vec<Element>,
}

impl Tree {
    /// The root element.
    pub(crate) fn root(&self) -> Result<&Element> {
        self.elements
            .first()
            .ok_or_else(|| Error::InvalidFormat("the part has no root element".into()))
    }

    /// The indices of the direct children of element `index`.
    pub(crate) fn children(&self, index: usize) -> impl Iterator<Item = usize> + '_ {
        self.elements
            .iter()
            .enumerate()
            .skip(index + 1)
            .take_while(move |(_, element)| element.depth > self.elements[index].depth)
            .filter(move |(_, element)| element.parent == Some(index))
            .map(|(child, _)| child)
    }

    /// The first direct child of element `index` with local name `local`.
    pub(crate) fn child(&self, index: usize, local: &[u8]) -> Option<usize> {
        self.children(index)
            .find(|child| self.elements[*child].is(local))
    }

    /// The namespace bindings in scope inside element `index`: the
    /// declarations of every ancestor and of the element itself, outermost
    /// first.
    pub(crate) fn bindings(&self, index: usize) -> Vec<(Option<Vec<u8>>, Vec<u8>)> {
        let mut chain = Vec::new();
        let mut current = Some(index);
        while let Some(at) = current {
            chain.push(at);
            current = self.elements[at].parent;
        }
        let mut bindings: Vec<(Option<Vec<u8>>, Vec<u8>)> = Vec::new();
        for at in chain.into_iter().rev() {
            for declaration in &self.elements[at].declarations {
                bindings.retain(|(prefix, _)| *prefix != declaration.0);
                bindings.push(declaration.clone());
            }
        }
        bindings
    }
}

/// Every `paraId` attribute value (any prefix) of `xml` that parses as
/// hexadecimal: the paragraph identifiers a fresh one must avoid (record
/// 0345). A plain byte search; values that do not parse are skipped.
pub(crate) fn paragraph_ids(xml: &[u8]) -> Vec<u32> {
    const KEY: &[u8] = b"paraId=";
    let mut ids = Vec::new();
    let mut at = 0usize;
    while let Some(found) = xml[at..]
        .windows(KEY.len())
        .position(|window| window == KEY)
    {
        let value = at + found + KEY.len();
        at = value;
        let Some(&quote) = xml.get(value) else { break };
        if quote != b'"' && quote != b'\'' {
            continue;
        }
        let Some(length) = xml[value + 1..].iter().position(|byte| *byte == quote) else {
            break;
        };
        if let Some(id) = std::str::from_utf8(&xml[value + 1..value + 1 + length])
            .ok()
            .and_then(|text| u32::from_str_radix(text.trim(), 16).ok())
        {
            ids.push(id);
        }
    }
    ids
}

/// Scan `xml`, recording every element whose depth is at most `depth`.
pub(crate) fn scan(xml: &[u8], depth: usize) -> Result<Tree> {
    scan_kept(xml, depth, |_| true)
}

/// [`scan`], recording a direct child of the root only when `keep` accepts
/// its local name; the subtree of a skipped child is checked and not
/// recorded (the styles catalog skips `w:latentStyles` and
/// `w:docDefaults`).
pub(crate) fn scan_kept(xml: &[u8], depth: usize, keep: impl Fn(&[u8]) -> bool) -> Result<Tree> {
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut elements: Vec<Element> = Vec::new();
    // The recorded element at each open depth, `None` beyond `depth`.
    let mut open: Vec<Option<usize>> = Vec::new();
    let mut count = 0usize;
    let mut seen_root = false;
    let offset = |position: u64| {
        origin
            .offset(position)
            .ok_or_else(|| Error::InvalidFormat("part offset overflow".into()))
    };
    loop {
        let start = offset(reader.buffer_position())?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = offset(reader.buffer_position())?;
        match event {
            event @ (Event::Start(_) | Event::Empty(_)) => {
                let empty = matches!(event, Event::Empty(_));
                let (Event::Start(element) | Event::Empty(element)) = event else {
                    continue;
                };
                if open.is_empty() {
                    if seen_root {
                        return Err(Error::InvalidFormat("the part has two roots".into()));
                    }
                    seen_root = true;
                }
                count += 1;
                if count > MAX_ELEMENTS {
                    return Err(Error::InvalidFormat(
                        "the part has too many elements".into(),
                    ));
                }
                let level = open.len();
                if level >= MAX_DEPTH {
                    return Err(Error::InvalidFormat("the part nests too deeply".into()));
                }
                let parent = open.last().copied().flatten();
                let kept = match level {
                    0 => true,
                    1 => keep(element.local_name().as_ref()),
                    _ => parent.is_some(),
                };
                let recorded = if level <= depth && kept {
                    elements.push(record(&reader, &element, start, end, empty, level, parent)?);
                    Some(elements.len() - 1)
                } else {
                    None
                };
                if !empty {
                    open.push(recorded);
                }
            },
            Event::End(_) => {
                let Some(recorded) = open.pop() else {
                    return Err(Error::InvalidFormat(
                        "the part has an unmatched end tag".into(),
                    ));
                };
                if let Some(index) = recorded {
                    elements[index].close = start;
                    elements[index].end = end;
                }
            },
            Event::DocType(_) => {
                return Err(Error::Xml("DOCTYPE is forbidden in a Word part".into()));
            },
            Event::Eof => {
                if !open.is_empty() || !seen_root {
                    return Err(Error::InvalidFormat("the part is unterminated".into()));
                }
                return Ok(Tree { elements });
            },
            _ => {},
        }
    }
}

fn record(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    start: usize,
    end: usize,
    empty: bool,
    depth: usize,
    parent: Option<usize>,
) -> Result<Element> {
    let name = element.name();
    let local = element.local_name().as_ref().to_vec().into_boxed_slice();
    let prefix = name
        .prefix()
        .map_or_else(Vec::new, |prefix| prefix.into_inner().to_vec())
        .into_boxed_slice();
    let mut attributes = Vec::new();
    let mut declarations = Vec::new();
    for attribute in first_wins(element) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        if key == b"xmlns" {
            declarations.push((None, value.into_bytes()));
        } else if let Some(declared) = key.strip_prefix(b"xmlns:") {
            declarations.push((Some(declared.to_vec()), value.into_bytes()));
        } else {
            attributes.push((
                attribute
                    .key
                    .local_name()
                    .as_ref()
                    .to_vec()
                    .into_boxed_slice(),
                value,
            ));
        }
    }
    Ok(Element {
        start,
        tag_end: end,
        close: end,
        end,
        empty,
        depth,
        local,
        prefix,
        attributes,
        declarations,
        parent,
    })
}

/// Whether an on/off value (`w:val` of a toggle, absent meaning on) is on.
pub(crate) fn on_off(value: Option<&str>) -> Option<bool> {
    match value {
        None => Some(true),
        Some("1" | "true" | "on") => Some(true),
        Some("0" | "false" | "off") => Some(false),
        Some(_) => None,
    }
}

/// `prefix:local`, or `local` for an empty prefix.
pub(crate) fn qualified(prefix: &[u8], local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{}:{local}", String::from_utf8_lossy(prefix))
    }
}

/// `text` escaped for an attribute value or character data.
pub(crate) fn escape(text: &str) -> String {
    quick_xml::escape::escape(text).into_owned()
}

/// Replace `start..end` of `xml` with `bytes`.
pub(crate) fn splice(xml: &[u8], start: usize, end: usize, bytes: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(xml.len() - (end - start) + bytes.len());
    output.extend_from_slice(&xml[..start]);
    output.extend_from_slice(bytes);
    output.extend_from_slice(&xml[end..]);
    output
}

/// One child element a rewrite sets (`Some`) or removes (`None`).
pub(crate) struct Child {
    pub(crate) local: &'static str,
    pub(crate) bytes: Option<Vec<u8>>,
}

/// Rewrite the direct children of element `index` (local name `name`) of
/// `xml`: each listed child replaces the existing child of its name, is
/// removed, or is inserted at its position in the schema child `order`.
pub(crate) fn rewrite_children(
    xml: &[u8],
    tree: &Tree,
    index: usize,
    name: &str,
    children: &[Child],
    order: &[&str],
) -> Result<Vec<u8>> {
    if children.is_empty() {
        return Ok(xml.to_vec());
    }
    let element = &tree.elements[index];
    let rank = |local: &[u8]| order.iter().position(|known| known.as_bytes() == local);
    // (start, end, rank, bytes): replacements keep the child's range;
    // insertions are empty ranges ordered by rank.
    let mut edits: Vec<(usize, usize, usize, Vec<u8>)> = Vec::new();
    let existing: Vec<usize> = tree.children(index).collect();
    let mut inserts: Vec<(usize, Vec<u8>)> = Vec::new();
    for child in children {
        let found = existing
            .iter()
            .copied()
            .find(|at| tree.elements[*at].is(child.local.as_bytes()));
        match (found, &child.bytes) {
            (Some(at), bytes) => {
                let old = &tree.elements[at];
                edits.push((old.start, old.end, 0, bytes.clone().unwrap_or_default()));
            },
            (None, Some(bytes)) => {
                let new_rank = rank(child.local.as_bytes()).ok_or_else(|| {
                    Error::InvalidFormat("the child has no schema position".into())
                })?;
                inserts.push((new_rank, bytes.clone()));
            },
            (None, None) => {},
        }
    }
    if element.empty {
        if inserts.is_empty() {
            return Ok(xml.to_vec());
        }
        inserts.sort_by_key(|(rank, _)| *rank);
        let mut replacement = xml[element.start..element.tag_end - 2].to_vec();
        while replacement.last().is_some_and(u8::is_ascii_whitespace) {
            replacement.pop();
        }
        replacement.push(b'>');
        for (_, bytes) in &inserts {
            replacement.extend_from_slice(bytes);
        }
        replacement
            .extend_from_slice(format!("</{}>", qualified(&element.prefix, name)).as_bytes());
        return Ok(splice(xml, element.start, element.end, &replacement));
    }
    for (new_rank, bytes) in inserts {
        let anchor = existing
            .iter()
            .map(|at| &tree.elements[*at])
            .find(|child| rank(&child.local).is_some_and(|known| known > new_rank))
            .map_or(element.close, |child| child.start);
        edits.push((anchor, anchor, new_rank, bytes));
    }
    edits.sort_by_key(|(start, end, rank, _)| (*start, *end != *start, *rank));
    let mut output = Vec::with_capacity(xml.len());
    let mut cursor = 0usize;
    for (start, end, _, bytes) in edits {
        if start < cursor {
            return Err(Error::InvalidFormat("part edits overlap".into()));
        }
        output.extend_from_slice(&xml[cursor..start]);
        output.extend_from_slice(&bytes);
        cursor = end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(output)
}

/// `bytes` wrapped in a `root` element of `prefix`.
pub(crate) fn wrap(prefix: &[u8], root: &str, bytes: &[u8]) -> Vec<u8> {
    let name = qualified(prefix, root);
    let mut wrapper = Vec::with_capacity(bytes.len() + 2 * name.len() + 5);
    wrapper.push(b'<');
    wrapper.extend_from_slice(name.as_bytes());
    wrapper.push(b'>');
    wrapper.extend_from_slice(bytes);
    wrapper.extend_from_slice(b"</");
    wrapper.extend_from_slice(name.as_bytes());
    wrapper.push(b'>');
    wrapper
}

/// The content of a wrapper made by [`wrap`].
fn unwrap<'a>(prefix: &[u8], root: &str, wrapper: &'a [u8]) -> Result<&'a [u8]> {
    let name = qualified(prefix, root);
    let open = name.len() + 2;
    let close = name.len() + 3;
    if wrapper.len() < open + close
        || !wrapper.starts_with(format!("<{name}>").as_bytes())
        || !wrapper.ends_with(format!("</{name}>").as_bytes())
    {
        return Err(Error::InvalidFormat(
            "a property splice rewrote its wrapper".into(),
        ));
    }
    Ok(&wrapper[open..wrapper.len() - close])
}

/// Whether `bytes` is an empty property container.
fn is_empty_container(prefix: &[u8], local: &str, bytes: &[u8]) -> bool {
    let name = qualified(prefix, local);
    bytes.is_empty()
        || bytes == format!("<{name}/>").as_bytes()
        || bytes == format!("<{name}></{name}>").as_bytes()
}

/// The new `w:pPr` (or its removal) that makes a container's paragraph
/// properties read as `target`.
pub(crate) fn respliced_paragraph(
    prefix: &[u8],
    before: Option<Wrapped>,
    current: &properties::Paragraph,
    target: &properties::Paragraph,
    strict: bool,
    bindings: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Option<Vec<u8>>> {
    let wrapper = before.map_or_else(|| wrap(prefix, "p", &[]), |(wrapper, _)| wrapper);
    let spliced = properties::splice_paragraph(&wrapper, current, target, strict, bindings)
        .map_err(splice_error)?
        .unwrap_or(wrapper);
    let inner = unwrap(prefix, "p", &spliced)?;
    Ok((!is_empty_container(prefix, "pPr", inner)).then(|| inner.to_vec()))
}

/// The new `w:rPr` (or its removal) that makes a container's run
/// properties read as `target`.
pub(crate) fn respliced_run(
    prefix: &[u8],
    before: Option<Wrapped>,
    current: &properties::Run,
    target: &properties::Run,
    bindings: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Option<Vec<u8>>> {
    let wrapper = before.map_or_else(|| wrap(prefix, "r", &[]), |(wrapper, _)| wrapper);
    let spliced = properties::splice_run(&wrapper, current, target, bindings)
        .map_err(splice_error)?
        .unwrap_or(wrapper);
    let inner = unwrap(prefix, "r", &spliced)?;
    Ok((!is_empty_container(prefix, "rPr", inner)).then(|| inner.to_vec()))
}

pub(crate) fn splice_error(error: properties::SpliceError) -> Error {
    match error {
        properties::SpliceError::Refused(_) => {
            Error::InvalidFormat("the property change cannot be written".into())
        },
        properties::SpliceError::Document(error) => error,
        properties::SpliceError::Transaction(error) => Error::InvalidFormat(error.to_string()),
    }
}

/// `<p:local p:val="value"/>`, or `<p:local/>` for `None`.
pub(crate) fn val_element(prefix: &[u8], local: &str, value: Option<&str>) -> Vec<u8> {
    let name = qualified(prefix, local);
    match value {
        Some(value) => format!(
            r#"<{name} {}="{}"/>"#,
            qualified(prefix, "val"),
            escape(value)
        ),
        None => format!("<{name}/>"),
    }
    .into_bytes()
}

/// An on/off element: `<p:local/>` for on, `<p:local p:val="0"/>` for off.
pub(crate) fn on_off_element(prefix: &[u8], local: &str, on: bool) -> Vec<u8> {
    val_element(prefix, local, (!on).then_some("0"))
}

/// Whether `xml` holds an element named in `locals`.
pub(crate) fn holds_any(xml: &[u8], locals: &[&[u8]]) -> Result<bool> {
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element)
                if locals.contains(&element.local_name().as_ref()) =>
            {
                return Ok(true);
            },
            Event::Eof => return Ok(false),
            _ => {},
        }
    }
}

/// The field markup a text set refuses rather than keep or drop (record
/// 0304): keeping a field adds its result to the new text, and replacing it
/// loses it.
pub(crate) const FIELD_MARKUP: &[&[u8]] = &[b"fldChar", b"instrText", b"fldSimple"];

/// The concatenated `w:t` text of the runs of `xml`, skipping runs that
/// hold an element named in `skip` and alternate-content fallbacks: the text
/// OfficeCLI reads from a header, footer or note (record 0319).
pub(crate) fn run_text(xml: &[u8], skip: &[&[u8]]) -> Result<String> {
    let mut reader = Reader::from_reader(xml);
    let mut text = String::new();
    let mut run = String::new();
    let mut in_run = 0usize;
    let mut skipped = false;
    let mut in_text = false;
    let mut fallback = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => match element.local_name().as_ref() {
                b"Fallback" => fallback += 1,
                b"r" if fallback == 0 => {
                    in_run += 1;
                    if in_run == 1 {
                        run.clear();
                        skipped = false;
                    }
                },
                b"t" if fallback == 0 && in_run > 0 => in_text = true,
                local if skip.contains(&local) => skipped = true,
                _ => {},
            },
            Event::Empty(element) if skip.contains(&element.local_name().as_ref()) => {
                skipped = true;
            },
            Event::End(element) => match element.local_name().as_ref() {
                b"Fallback" => fallback = fallback.saturating_sub(1),
                b"r" if fallback == 0 && in_run > 0 => {
                    in_run -= 1;
                    if in_run == 0 && !skipped {
                        text.push_str(&run);
                    }
                },
                b"t" => in_text = false,
                _ => {},
            },
            Event::Text(content) if in_text => {
                let decoded = content
                    .xml10_content()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                run.push_str(&decoded);
            },
            Event::GeneralRef(reference) if in_text => {
                if let Ok(Some(character)) = reference.resolve_char_ref() {
                    run.push(character);
                } else {
                    let name = reference
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    run.push_str(match name.as_ref() {
                        "amp" => "&",
                        "lt" => "<",
                        "gt" => ">",
                        "quot" => "\"",
                        "apos" => "'",
                        _ => "",
                    });
                }
            },
            Event::Eof => return Ok(text),
            _ => {},
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn records_shallow_elements_with_ranges_and_bindings() {
        let xml = br#"<?xml version="1.0"?><w:styles xmlns:w="urn:w"><w:style w:styleId="A"><w:name w:val="Alpha"/><w:rPr><w:b/></w:rPr></w:style></w:styles>"#;
        let tree = scan(xml, 2).unwrap();
        let root = tree.root().unwrap();
        assert!(root.is(b"styles"));
        let style = tree.child(0, b"style").unwrap();
        assert_eq!(tree.elements[style].attribute(b"styleId"), Some("A"));
        let element = &tree.elements[style];
        assert!(xml[element.start..element.end].starts_with(b"<w:style "));
        assert!(xml[element.start..element.end].ends_with(b"</w:style>"));
        let name = tree.child(style, b"name").unwrap();
        assert!(tree.elements[name].empty);
        assert_eq!(tree.elements[name].attribute(b"val"), Some("Alpha"));
        let run = tree.child(style, b"rPr").unwrap();
        // `w:b` is below the recorded depth.
        assert_eq!(tree.children(run).count(), 0);
        assert_eq!(
            tree.bindings(run),
            vec![(Some(b"w".to_vec()), b"urn:w".to_vec())]
        );
        assert_eq!(on_off(None), Some(true));
        assert_eq!(on_off(Some("0")), Some(false));
        assert_eq!(on_off(Some("maybe")), None);
    }

    #[test]
    fn run_text_concatenates_runs_and_skips_marks_and_fallbacks() {
        let xml = br#"<w:ftn xmlns:w="w" xmlns:mc="m"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>A &amp; </w:t></w:r><mc:AlternateContent><mc:Fallback><w:r><w:t>X</w:t></w:r></mc:Fallback></mc:AlternateContent><w:r><w:t>B&#65;</w:t></w:r></w:p></w:ftn>"#;
        assert_eq!(run_text(xml, &[b"footnoteRef"]).unwrap(), "A & BA");
    }

    #[test]
    fn refuses_malformed_parts() {
        assert!(scan(b"<a><b></a>", 3).is_err());
        assert!(scan(b"<a/><b/>", 3).is_err());
        assert!(scan(b"", 3).is_err());
    }
}
