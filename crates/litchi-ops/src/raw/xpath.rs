//! `raw-set`: an XPath location-path subset over one XML part, and the byte
//! splice that applies an [`Action`] at every match.
//!
//! The splice keeps every byte outside the edited ranges. The caller
//! publishes the result through the package's audited XML replacement, so
//! a fragment that breaks the part is refused before any handle sees it.

use std::borrow::Cow;

use quick_xml::events::{BytesStart, Event};

use super::Action;
use crate::{Code, Error, Result};

/// OfficeCLI's auto-registered prefixes (`RawXmlHelper.CommonNamespaces`).
const KNOWN: [(&str, &str); 12] = [
    (
        "w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    ),
    (
        "r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ),
    ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
    (
        "p",
        "http://schemas.openxmlformats.org/presentationml/2006/main",
    ),
    (
        "x",
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    ),
    (
        "wp",
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    ),
    (
        "mc",
        "http://schemas.openxmlformats.org/markup-compatibility/2006",
    ),
    (
        "c",
        "http://schemas.openxmlformats.org/drawingml/2006/chart",
    ),
    (
        "xdr",
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    ),
    (
        "wps",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    ),
    (
        "wp14",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    ),
    ("v", "urn:schemas-microsoft-com:vml"),
];

/// The most steps one XPath may have, and the most bytes.
const MAX_STEPS: usize = 32;
const MAX_XPATH_BYTES: usize = 4096;

/// Apply `action` at every element `xpath` matches in `xml`. Returns the
/// edited bytes and the match count.
pub(super) fn edit(
    xml: &[u8],
    xpath: &str,
    action: Action,
    fragment: Option<&str>,
) -> Result<(Vec<u8>, usize)> {
    let steps = parse(xpath)?;
    let document = Document::parse(xml)?;
    // Name tests resolve through OfficeCLI's prefixes, overridden by the
    // root's own declarations.
    let prefixes = Prefixes::of(&document);
    let tests = steps
        .iter()
        .map(|step| step.resolve(&prefixes))
        .collect::<Result<Vec<_>>>()?;
    let matches = document.select(&tests);
    if matches.is_empty() {
        return Err(Error::new(Code::ZeroMatches, "the XPath matched no element")
            .with_suggestion("known prefixes: w, r, a, p, x, wp, mc, c, xdr, wps, wp14, v and the part root's own"));
    }
    let structural = action != Action::SetAttr;
    if structural {
        for pair in matches.windows(2) {
            if document.elements[pair[1]].start < document.elements[pair[0]].end {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "the XPath matches elements nested in one another",
                )
                .with_suggestion("narrow the XPath so the matches are disjoint"));
            }
        }
    }
    let mut splices: Vec<Splice> = Vec::with_capacity(matches.len());
    for &index in &matches {
        splices.push(document.splice(index, action, fragment, &prefixes)?);
    }
    splices.sort_by_key(|splice| splice.start);
    let mut output =
        Vec::with_capacity(xml.len() + splices.iter().map(|s| s.bytes.len()).sum::<usize>());
    let mut cursor = 0;
    for splice in &splices {
        if splice.start < cursor {
            return Err(Error::new(
                Code::InvalidArgument,
                "the XPath matches overlapping edits",
            ));
        }
        output.extend_from_slice(&xml[cursor..splice.start]);
        output.extend_from_slice(&splice.bytes);
        cursor = splice.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok((output, matches.len()))
}

/// One replacement of `xml[start..end]`.
struct Splice {
    start: usize,
    end: usize,
    bytes: Vec<u8>,
}

// ---------------------------------------------------------------------------
// The XPath subset.

struct Step {
    /// Preceded by `//` (descendant-or-self) instead of `/` (child).
    descendant: bool,
    name: Option<(Option<String>, String)>,
    predicates: Vec<Predicate>,
}

enum Predicate {
    Position(usize),
    Last,
    Attribute { name: String, value: String },
}

struct Test<'s> {
    descendant: bool,
    /// `None` for `*`; else the namespace (`None` for no namespace) and
    /// local name.
    name: Option<(Option<String>, &'s str)>,
    predicates: &'s [Predicate],
}

fn invalid(message: &'static str) -> Error {
    Error::new(Code::InvalidXpath, message).with_suggestion(
        "use a location path of '/' and '//' steps with prefixed names or '*', and [n], [last()] or [@name='value'] predicates",
    )
}

fn parse(xpath: &str) -> Result<Vec<Step>> {
    if xpath.len() > MAX_XPATH_BYTES {
        return Err(invalid("the XPath is too long"));
    }
    let bytes = xpath.trim().as_bytes();
    if !bytes.starts_with(b"/") {
        return Err(invalid(
            "the XPath must be absolute: start with '/' or '//'",
        ));
    }
    let mut steps = Vec::new();
    let mut at = 0;
    while at < bytes.len() {
        let descendant = bytes[at..].starts_with(b"//");
        at += if descendant { 2 } else { 1 };
        let start = at;
        while at < bytes.len() && !matches!(bytes[at], b'/' | b'[') {
            at += 1;
        }
        let name = std::str::from_utf8(&bytes[start..at]).map_err(|_| invalid("bad step"))?;
        let name = if name == "*" {
            None
        } else {
            let (prefix, local) = match name.split_once(':') {
                Some((prefix, local)) => (Some(prefix), local),
                None => (None, name),
            };
            if !is_ncname(local) || prefix.is_some_and(|prefix| !is_ncname(prefix)) {
                return Err(invalid("a step must name an element or '*'"));
            }
            Some((prefix.map(str::to_owned), local.to_owned()))
        };
        let mut predicates = Vec::new();
        while at < bytes.len() && bytes[at] == b'[' {
            let close = bytes[at..]
                .iter()
                .position(|&byte| byte == b']')
                .ok_or_else(|| invalid("an unclosed predicate"))?;
            let body = std::str::from_utf8(&bytes[at + 1..at + close])
                .map_err(|_| invalid("bad predicate"))?
                .trim();
            predicates.push(predicate(body)?);
            at += close + 1;
        }
        if at < bytes.len() && bytes[at] != b'/' {
            return Err(invalid("unexpected characters after a step"));
        }
        steps.push(Step {
            descendant,
            name,
            predicates,
        });
        if steps.len() > MAX_STEPS {
            return Err(invalid("the XPath has too many steps"));
        }
    }
    if steps.is_empty() {
        return Err(invalid("the XPath has no step"));
    }
    Ok(steps)
}

fn predicate(body: &str) -> Result<Predicate> {
    if body == "last()" {
        return Ok(Predicate::Last);
    }
    if let Ok(position) = body.parse::<usize>() {
        if position == 0 {
            return Err(invalid("positions start at 1"));
        }
        return Ok(Predicate::Position(position));
    }
    if let Some(rest) = body.strip_prefix('@') {
        let (name, value) = rest
            .split_once('=')
            .ok_or_else(|| invalid("an attribute predicate is [@name='value']"))?;
        let name = name.trim();
        let value = value.trim();
        let quoted = value.len() >= 2
            && ((value.starts_with('\'') && value.ends_with('\''))
                || (value.starts_with('"') && value.ends_with('"')));
        if !quoted || name.split(':').any(|part| !is_ncname(part)) || name.matches(':').count() > 1
        {
            return Err(invalid("an attribute predicate is [@name='value']"));
        }
        return Ok(Predicate::Attribute {
            name: name.to_owned(),
            value: value[1..value.len() - 1].to_owned(),
        });
    }
    Err(invalid("unsupported predicate"))
}

fn is_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    chars
        .next()
        .is_some_and(|first| first.is_alphabetic() || first == '_')
        && chars.all(|c| c.is_alphanumeric() || matches!(c, '_' | '-' | '.'))
}

impl Step {
    fn resolve<'s>(&'s self, prefixes: &Prefixes) -> Result<Test<'s>> {
        let name = match &self.name {
            None => None,
            Some((None, local)) => Some((None, local.as_str())),
            Some((Some(prefix), local)) => {
                let uri = prefixes
                    .uri(prefix)
                    .ok_or_else(|| invalid("the XPath uses an unknown namespace prefix"))?;
                Some((Some(uri.to_owned()), local.as_str()))
            },
        };
        Ok(Test {
            descendant: self.descendant,
            name,
            predicates: &self.predicates,
        })
    }
}

// ---------------------------------------------------------------------------
// The part's element table.

struct Element {
    parent: Option<usize>,
    namespace: Option<String>,
    /// The qualified name as written (`w:p`).
    qname: Vec<u8>,
    local: Vec<u8>,
    /// 1-based position among element siblings of the same expanded name,
    /// and among all element siblings.
    position_name: usize,
    position_any: usize,
    /// `<` of the start tag, byte after `>` of the start tag, `<` of the end
    /// tag, and byte after the end tag. An empty element has
    /// `content_start == content_end == end`.
    start: usize,
    content_start: usize,
    content_end: usize,
    end: usize,
    empty: bool,
    /// The element's own namespace declarations (`""` is the default).
    declarations: Vec<(String, String)>,
}

struct Document<'x> {
    xml: &'x [u8],
    elements: Vec<Element>,
    /// Per element: (child expanded-name totals, total element children),
    /// for `last()`.
    totals: Vec<Vec<(Option<String>, Vec<u8>, usize)>>,
    totals_any: Vec<usize>,
}

fn malformed() -> Error {
    Error::new(
        Code::InvalidInput,
        "the part is not namespace-well-formed XML",
    )
}

impl<'x> Document<'x> {
    fn parse(xml: &'x [u8]) -> Result<Self> {
        let mut reader = quick_xml::Reader::from_reader(xml);
        reader.config_mut().check_end_names = true;
        let mut elements: Vec<Element> = Vec::new();
        let mut totals: Vec<Vec<(Option<String>, Vec<u8>, usize)>> = Vec::new();
        let mut totals_any: Vec<usize> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        let mut roots = 0_usize;
        loop {
            let before = usize::try_from(reader.buffer_position()).map_err(|_| malformed())?;
            let event = reader.read_event().map_err(|_| malformed())?;
            let after = usize::try_from(reader.buffer_position()).map_err(|_| malformed())?;
            let (event, empty) = match event {
                Event::Empty(start) => (Event::Start(start), true),
                other => (other, false),
            };
            match event {
                Event::Start(start) => {
                    let parent = stack.last().copied();
                    if parent.is_none() {
                        roots += 1;
                    }
                    let declarations = declarations(&start)?;
                    let qname = start.name().as_ref().to_vec();
                    let local = start.local_name().as_ref().to_vec();
                    let prefix = start.name().prefix().map(|p| p.as_ref().to_vec());
                    let index = elements.len();
                    // Resolve the element's namespace through its own and its
                    // ancestors' declarations.
                    let namespace = {
                        let key = prefix
                            .as_deref()
                            .map_or(Cow::Borrowed(""), String::from_utf8_lossy);
                        lookup(&declarations, &key)
                            .map(str::to_owned)
                            .or_else(|| {
                                stack.iter().rev().find_map(|&ancestor| {
                                    lookup(&elements[ancestor].declarations, &key)
                                        .map(str::to_owned)
                                })
                            })
                            .or_else(|| {
                                (key == "xml")
                                    .then(|| "http://www.w3.org/XML/1998/namespace".to_owned())
                            })
                    };
                    if prefix.is_some() && namespace.is_none() {
                        return Err(malformed());
                    }
                    let namespace = namespace.filter(|uri| !uri.is_empty());
                    let (position_name, position_any) = match parent {
                        Some(parent) => {
                            totals_any[parent] += 1;
                            let entry = totals[parent]
                                .iter_mut()
                                .find(|(ns, name, _)| *ns == namespace && *name == local);
                            let count = match entry {
                                Some(entry) => {
                                    entry.2 += 1;
                                    entry.2
                                },
                                None => {
                                    totals[parent].push((namespace.clone(), local.clone(), 1));
                                    1
                                },
                            };
                            (count, totals_any[parent])
                        },
                        None => (1, 1),
                    };
                    elements.push(Element {
                        parent,
                        namespace,
                        qname,
                        local,
                        position_name,
                        position_any,
                        start: before,
                        content_start: after,
                        content_end: after,
                        end: after,
                        empty,
                        declarations,
                    });
                    totals.push(Vec::new());
                    totals_any.push(0);
                    if !empty {
                        stack.push(index);
                    }
                },
                Event::End(_) => {
                    let index = stack.pop().ok_or_else(malformed)?;
                    elements[index].content_end = before;
                    elements[index].end = after;
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if !stack.is_empty() || roots != 1 {
            return Err(malformed());
        }
        Ok(Self {
            xml,
            elements,
            totals,
            totals_any,
        })
    }

    /// The matches of the location path, in document order.
    fn select(&self, tests: &[Test<'_>]) -> Vec<usize> {
        // memo[k][e]: whether element e matches the path prefix 0..=k.
        let mut memo: Vec<Vec<bool>> = Vec::with_capacity(tests.len());
        for (k, test) in tests.iter().enumerate() {
            let mut row = vec![false; self.elements.len()];
            for (index, element) in self.elements.iter().enumerate() {
                if !self.test(index, element, test) {
                    continue;
                }
                row[index] = if k == 0 {
                    test.descendant || element.parent.is_none()
                } else if test.descendant {
                    let mut ancestor = element.parent;
                    let mut found = false;
                    while let Some(current) = ancestor {
                        if memo[k - 1][current] {
                            found = true;
                            break;
                        }
                        ancestor = self.elements[current].parent;
                    }
                    found
                } else {
                    element.parent.is_some_and(|parent| memo[k - 1][parent])
                };
            }
            memo.push(row);
        }
        memo.last()
            .map(|row| {
                row.iter()
                    .enumerate()
                    .filter_map(|(index, &hit)| hit.then_some(index))
                    .collect()
            })
            .unwrap_or_default()
    }

    fn test(&self, index: usize, element: &Element, test: &Test<'_>) -> bool {
        if let Some((namespace, local)) = &test.name {
            if element.local != local.as_bytes() || element.namespace != *namespace {
                return false;
            }
        }
        test.predicates.iter().all(|predicate| match predicate {
            Predicate::Position(position) => {
                let actual = if test.name.is_some() {
                    element.position_name
                } else {
                    element.position_any
                };
                actual == *position
            },
            Predicate::Last => match element.parent {
                None => true,
                Some(parent) => {
                    if test.name.is_some() {
                        self.totals[parent]
                            .iter()
                            .find(|(ns, name, _)| {
                                *ns == element.namespace && *name == element.local
                            })
                            .is_some_and(|(_, _, total)| *total == element.position_name)
                    } else {
                        self.totals_any[parent] == element.position_any
                    }
                },
            },
            Predicate::Attribute { name, value } => self
                .attribute(index, name)
                .is_some_and(|actual| actual == *value),
        })
    }

    /// The unescaped value of the attribute written `name` on element
    /// `index`.
    fn attribute(&self, index: usize, name: &str) -> Option<String> {
        let element = &self.elements[index];
        let tag = self.start_tag(element)?;
        let mut attributes = tag.attributes();
        attributes.with_checks(false);
        for attribute in attributes.flatten() {
            if attribute.key.as_ref() == name.as_bytes() {
                return attribute
                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                    .ok()
                    .map(Cow::into_owned);
            }
        }
        None
    }

    fn start_tag(&self, element: &Element) -> Option<BytesStart<'x>> {
        let raw = &self.xml[element.start..element.content_start];
        // `<name attrs>` or `<name attrs/>`
        let inner = raw.strip_prefix(b"<")?;
        let inner = inner.strip_suffix(b">")?;
        let inner = inner.strip_suffix(b"/").unwrap_or(inner);
        let name_len = inner
            .iter()
            .position(|byte| byte.is_ascii_whitespace())
            .unwrap_or(inner.len());
        Some(BytesStart::from_content(
            std::str::from_utf8(inner).ok()?,
            name_len,
        ))
    }

    /// The namespace declarations in scope at the start of `index`'s
    /// content (its own included), innermost last.
    fn scope(&self, index: Option<usize>) -> Vec<(String, String)> {
        let mut chain = Vec::new();
        let mut current = index;
        while let Some(element) = current {
            chain.push(element);
            current = self.elements[element].parent;
        }
        let mut scope: Vec<(String, String)> = Vec::new();
        for &element in chain.iter().rev() {
            for (prefix, uri) in &self.elements[element].declarations {
                scope.retain(|(existing, _)| existing != prefix);
                scope.push((prefix.clone(), uri.clone()));
            }
        }
        scope
    }

    fn splice(
        &self,
        index: usize,
        action: Action,
        fragment: Option<&str>,
        prefixes: &Prefixes,
    ) -> Result<Splice> {
        let element = &self.elements[index];
        let needs_parent = matches!(
            action,
            Action::InsertBefore | Action::InsertAfter | Action::Remove
        );
        if needs_parent && element.parent.is_none() {
            return Err(Error::new(
                Code::InvalidArgument,
                "the document root element cannot be removed or have siblings",
            )
            .with_suggestion("target a child element with a more specific XPath"));
        }
        let required = || {
            fragment.ok_or_else(|| {
                Error::new(Code::MissingArgument, "this raw-set action requires xml")
            })
        };
        match action {
            Action::Remove => Ok(Splice {
                start: element.start,
                end: element.end,
                bytes: Vec::new(),
            }),
            Action::SetAttr => self.set_attribute(index, required()?, prefixes),
            Action::Append | Action::Prepend => {
                let scope = self.scope(Some(index));
                let bytes = prepare(required()?, &scope, prefixes)?;
                if element.empty {
                    // `<x a="1"/>` becomes `<x a="1">fragment</x>`.
                    let tag = &self.xml[element.start..element.content_start];
                    let open = tag
                        .strip_suffix(b"/>")
                        .ok_or_else(malformed)?
                        .trim_ascii_end();
                    let mut out = open.to_vec();
                    out.push(b'>');
                    out.extend_from_slice(&bytes);
                    out.extend_from_slice(b"</");
                    out.extend_from_slice(&element.qname);
                    out.push(b'>');
                    Ok(Splice {
                        start: element.start,
                        end: element.end,
                        bytes: out,
                    })
                } else {
                    let at = if action == Action::Append {
                        element.content_end
                    } else {
                        element.content_start
                    };
                    Ok(Splice {
                        start: at,
                        end: at,
                        bytes,
                    })
                }
            },
            Action::InsertBefore | Action::InsertAfter | Action::Replace => {
                let scope = self.scope(element.parent);
                let bytes = prepare(required()?, &scope, prefixes)?;
                let (start, end) = match action {
                    Action::InsertBefore => (element.start, element.start),
                    Action::InsertAfter => (element.end, element.end),
                    _ => (element.start, element.end),
                };
                Ok(Splice { start, end, bytes })
            },
        }
    }

    fn set_attribute(&self, index: usize, assignment: &str, prefixes: &Prefixes) -> Result<Splice> {
        let element = &self.elements[index];
        let (name, value) = assignment
            .split_once('=')
            .ok_or_else(|| Error::new(Code::InvalidArgument, "setattr takes name=value"))?;
        let name = name.trim();
        if name.is_empty()
            || name.split(':').any(|part| !is_ncname(part))
            || name.matches(':').count() > 1
            || name.starts_with("xmlns")
        {
            return Err(Error::new(
                Code::InvalidArgument,
                "setattr takes name=value with an XML attribute name",
            ));
        }
        let scope = self.scope(Some(index));
        let mut declaration = None;
        let target = match name.split_once(':') {
            Some((prefix, local)) => {
                let uri = lookup(&scope, prefix)
                    .map(str::to_owned)
                    .or_else(|| {
                        let uri = prefixes.uri(prefix)?.to_owned();
                        declaration = Some((prefix.to_owned(), uri.clone()));
                        Some(uri)
                    })
                    .ok_or_else(|| {
                        Error::new(
                            Code::InvalidArgument,
                            "setattr uses an unknown namespace prefix",
                        )
                    })?;
                (Some(uri), local)
            },
            None => (None, name),
        };
        let escaped = escape_attribute(value);
        let tag_start = element.start;
        let raw = &self.xml[element.start..element.content_start];
        // Find an existing attribute with the same expanded name.
        let tag = self.start_tag(element).ok_or_else(malformed)?;
        let mut attributes = tag.attributes();
        attributes.with_checks(false);
        for attribute in attributes.flatten() {
            let key = attribute.key;
            let local = key.local_name();
            let namespace = match key.prefix() {
                Some(prefix) => {
                    let prefix = String::from_utf8_lossy(prefix.as_ref());
                    if prefix == "xmlns" {
                        continue;
                    }
                    lookup(&scope, &prefix).map(str::to_owned)
                },
                None => {
                    if key.as_ref() == b"xmlns" {
                        continue;
                    }
                    None
                },
            };
            if local.as_ref() == target.1.as_bytes() && namespace == target.0 {
                // Replace the value between the quotes.
                let raw_value: &[u8] = &attribute.value;
                let key_bytes = key.as_ref();
                let key_at = find(raw, key_bytes, raw_value).ok_or_else(malformed)?;
                return Ok(Splice {
                    start: tag_start + key_at.0,
                    end: tag_start + key_at.1,
                    bytes: escaped.into_bytes(),
                });
            }
        }
        // Insert before `>` or `/>`.
        let close = if element.empty { 2 } else { 1 };
        let at = element.content_start - close;
        let mut bytes = Vec::new();
        if let Some((prefix, uri)) = declaration {
            bytes.extend_from_slice(
                format!(" xmlns:{prefix}=\"{}\"", escape_attribute(&uri)).as_bytes(),
            );
        }
        bytes.extend_from_slice(format!(" {name}=\"{escaped}\"").as_bytes());
        // Keep `<x a="1" />` spacing: insert right after the last
        // non-space byte before the close.
        let trimmed = self.xml[..at]
            .iter()
            .rposition(|byte| !byte.is_ascii_whitespace())
            .map_or(at, |position| position + 1)
            .max(element.start + 1 + element.qname.len());
        Ok(Splice {
            start: trimmed,
            end: trimmed,
            bytes,
        })
    }
}

/// The value range of attribute `key` (whose raw value is `value`) inside
/// the start tag `raw`.
fn find(raw: &[u8], key: &[u8], value: &[u8]) -> Option<(usize, usize)> {
    let mut from = 0;
    while let Some(offset) = memchr::memmem::find(&raw[from..], key) {
        let at = from + offset;
        let preceded = at > 0 && raw[at - 1].is_ascii_whitespace();
        let mut cursor = at + key.len();
        while raw.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if preceded && raw.get(cursor) == Some(&b'=') {
            cursor += 1;
            while raw.get(cursor).is_some_and(u8::is_ascii_whitespace) {
                cursor += 1;
            }
            if matches!(raw.get(cursor), Some(b'"' | b'\'')) {
                let start = cursor + 1;
                let end = start + value.len();
                if raw.get(start..end) == Some(value) {
                    return Some((start, end));
                }
            }
        }
        from = at + 1;
    }
    None
}

fn escape_attribute(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            other => out.push(other),
        }
    }
    out
}

fn lookup<'d>(declarations: &'d [(String, String)], prefix: &str) -> Option<&'d str> {
    declarations
        .iter()
        .rev()
        .find(|(existing, _)| existing == prefix)
        .map(|(_, uri)| uri.as_str())
}

/// The `xmlns` declarations of one start tag.
fn declarations(start: &BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut out = Vec::new();
    let mut attributes = start.attributes();
    attributes.with_checks(false);
    for attribute in attributes {
        let attribute = attribute.map_err(|_| malformed())?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            String::new()
        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            String::from_utf8_lossy(prefix).into_owned()
        } else {
            continue;
        };
        let uri = attribute
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .map_err(|_| malformed())?
            .into_owned();
        out.push((prefix, uri));
    }
    Ok(out)
}

/// The prefixes XPath name tests and fragments resolve through: OfficeCLI's,
/// overridden by the part root's declarations; plus the root's default
/// namespace, which unprefixed fragment elements take.
struct Prefixes {
    map: Vec<(String, String)>,
    default: Option<String>,
}

impl Prefixes {
    fn of(document: &Document<'_>) -> Self {
        let mut map: Vec<(String, String)> = KNOWN
            .iter()
            .map(|(prefix, uri)| ((*prefix).to_owned(), (*uri).to_owned()))
            .collect();
        let mut default = None;
        if let Some(root) = document.elements.first() {
            for (prefix, uri) in &root.declarations {
                if prefix.is_empty() {
                    default = Some(uri.clone()).filter(|uri| !uri.is_empty());
                } else {
                    map.retain(|(existing, _)| existing != prefix);
                    map.push((prefix.clone(), uri.clone()));
                }
            }
            if default.is_none() && root.qname == root.local {
                default = root.namespace.clone();
            }
        }
        Self { map, default }
    }

    fn uri(&self, prefix: &str) -> Option<&str> {
        lookup(&self.map, prefix)
    }
}

/// Parse a fragment, drop its inter-element whitespace, and declare on each
/// top-level element the namespaces it uses that `scope` lacks.
fn prepare(fragment: &str, scope: &[(String, String)], prefixes: &Prefixes) -> Result<Vec<u8>> {
    let bad = || {
        Error::new(Code::InvalidArgument, "the xml fragment is not well-formed")
            .with_suggestion("give one or more complete elements; prefixes resolve as in the XPath")
    };
    let mut reader = quick_xml::Reader::from_str(fragment);
    reader.config_mut().check_end_names = true;
    // First pass: well-formedness, used prefixes, own top-level
    // declarations.
    let mut used: Vec<String> = Vec::new();
    let mut unprefixed = false;
    let mut depth = 0_usize;
    let mut elements = 0_usize;
    loop {
        let (event, is_start) = match reader.read_event().map_err(|_| bad())? {
            Event::Empty(start) => (Event::Start(start), false),
            other => (other, true),
        };
        match event {
            Event::Start(start) => {
                match start.name().prefix() {
                    Some(prefix) => {
                        push_unique(&mut used, &String::from_utf8_lossy(prefix.as_ref()))
                    },
                    None => unprefixed = true,
                }
                let mut attributes = start.attributes();
                attributes.with_checks(false);
                for attribute in attributes {
                    let attribute = attribute.map_err(|_| bad())?;
                    if let Some(prefix) = attribute.key.prefix() {
                        let prefix = String::from_utf8_lossy(prefix.as_ref()).into_owned();
                        if prefix != "xmlns" && prefix != "xml" {
                            push_unique(&mut used, &prefix);
                        }
                    }
                }
                if depth == 0 {
                    elements += 1;
                }
                if is_start {
                    depth += 1;
                }
            },
            Event::End(_) => depth = depth.checked_sub(1).ok_or_else(bad)?,
            Event::Text(text) if depth == 0 => {
                if !text.iter().all(u8::is_ascii_whitespace) {
                    return Err(bad());
                }
            },
            Event::GeneralRef(_) | Event::CData(_) if depth == 0 => return Err(bad()),
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || elements == 0 {
        return Err(bad());
    }
    // The declarations each top-level element needs.
    let mut needed: Vec<(String, String)> = Vec::new();
    for prefix in &used {
        let uri = prefixes.uri(prefix);
        match (lookup(scope, prefix), uri) {
            (Some(existing), Some(uri)) if existing == uri => {},
            (Some(_), None) | (None, None) => {
                // Bound in scope only, or declared by the fragment itself:
                // leave it to the fragment and the audit.
                if lookup(scope, prefix).is_none()
                    && !fragment.contains(&format!("xmlns:{prefix}="))
                {
                    return Err(Error::new(
                        Code::InvalidArgument,
                        "the xml fragment uses an unknown namespace prefix",
                    ));
                }
            },
            (_, Some(uri)) => needed.push((prefix.clone(), uri.to_owned())),
        }
    }
    let default_needed = match (unprefixed, &prefixes.default) {
        (true, Some(default)) => lookup(scope, "") != Some(default.as_str()),
        _ => false,
    };
    // Second pass: re-serialize.
    let mut reader = quick_xml::Reader::from_str(fragment);
    let mut writer = quick_xml::Writer::new(Vec::with_capacity(fragment.len() + 128));
    let mut stack: Vec<Vec<u8>> = Vec::new();
    loop {
        let (event, is_start) = match reader.read_event().map_err(|_| bad())? {
            Event::Empty(start) => (Event::Start(start), false),
            other => (other, true),
        };
        let write = |writer: &mut quick_xml::Writer<Vec<u8>>, event: Event<'_>| {
            writer
                .write_event(event)
                .map_err(|_| Error::new(Code::InternalError, "fragment serialization failed"))
        };
        match event {
            Event::Start(start) => {
                let mut start = start.into_owned();
                if stack.is_empty() {
                    let declared = |prefix: &str| {
                        let mut attributes = start.attributes();
                        attributes.with_checks(false);
                        attributes.flatten().any(|attribute| {
                            let key = attribute.key.as_ref();
                            if prefix.is_empty() {
                                key == b"xmlns"
                            } else {
                                key.strip_prefix(b"xmlns:") == Some(prefix.as_bytes())
                            }
                        })
                    };
                    let mut extra: Vec<(String, String)> = Vec::new();
                    if default_needed && !declared("") {
                        if let Some(default) = &prefixes.default {
                            extra.push(("xmlns".to_owned(), default.clone()));
                        }
                    }
                    for (prefix, uri) in &needed {
                        if !declared(prefix) {
                            extra.push((format!("xmlns:{prefix}"), uri.clone()));
                        }
                    }
                    for (key, value) in &extra {
                        start.push_attribute((key.as_str(), value.as_str()));
                    }
                }
                if is_start {
                    stack.push(start.local_name().as_ref().to_vec());
                    write(&mut writer, Event::Start(start))?;
                } else {
                    write(&mut writer, Event::Empty(start))?;
                }
            },
            Event::End(end) => {
                stack.pop();
                write(&mut writer, Event::End(end))?;
            },
            Event::Text(text) => {
                let leaf = stack
                    .last()
                    .is_some_and(|local| local.as_slice() == b"t" || local.ends_with(b"Text"));
                if leaf || !text.iter().all(u8::is_ascii_whitespace) {
                    write(&mut writer, Event::Text(text))?;
                }
            },
            Event::Eof => break,
            Event::Decl(_) | Event::DocType(_) => return Err(bad()),
            other => write(&mut writer, other)?,
        }
    }
    Ok(writer.into_inner())
}

fn push_unique(values: &mut Vec<String>, value: &str) {
    if !values.iter().any(|existing| existing == value) {
        values.push(value.to_owned());
    }
}
