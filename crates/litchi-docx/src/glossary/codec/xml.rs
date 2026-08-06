//! XML syntax, bounded DOM, and canonical opaque-content emission for glossaries.

use super::super::*;
use super::super::{
    MAX, MAX_DEPTH, MAX_DOM_ATTRIBUTES, MAX_DOM_BYTES, MAX_DOM_CONTENT, MAX_DOM_TOKENS, MAX_NODES,
    MAX_VALUES, R, RS, W, WS,
};
use super::validation::{invalid, split, xml_char, xml_error};
use std::mem::size_of;

#[derive(Clone, Debug, Default)]
pub(in crate::glossary) struct NamespaceFrame {
    pub(in crate::glossary) parent: Option<Arc<NamespaceFrame>>,
    pub(in crate::glossary) local: Vec<(String, Arc<str>)>,
    pub(in crate::glossary) declaration_count: usize,
}

#[derive(Clone, Debug)]
pub(in crate::glossary) struct Attr {
    pub(in crate::glossary) q: String,
    pub(in crate::glossary) ns: Arc<str>,
    pub(in crate::glossary) l: String,
    pub(in crate::glossary) v: String,
}
#[derive(Clone, Debug)]
pub(in crate::glossary) enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}
#[derive(Clone, Debug)]
pub(in crate::glossary) struct Node {
    pub(in crate::glossary) q: String,
    pub(in crate::glossary) ns: Arc<str>,
    pub(in crate::glossary) l: String,
    pub(in crate::glossary) attrs: Vec<Attr>,
    pub(in crate::glossary) bindings: Arc<NamespaceFrame>,
    pub(in crate::glossary) content: Vec<Content>,
}

#[derive(Default)]
pub(in crate::glossary) struct DomBudget {
    pub(in crate::glossary) bytes: usize,
    pub(in crate::glossary) nodes: usize,
    pub(in crate::glossary) attributes: usize,
    pub(in crate::glossary) contents: usize,
    pub(in crate::glossary) tokens: usize,
}

impl DomBudget {
    fn charge(&mut self, bytes: usize) -> Result<()> {
        self.bytes = self
            .bytes
            .checked_add(bytes)
            .ok_or_else(|| invalid("glossary DOM allocation budget overflow"))?;
        if self.bytes > MAX_DOM_BYTES {
            return Err(invalid(
                "glossary DOM exceeds the 64 MiB owned-allocation budget",
            ));
        }
        Ok(())
    }

    fn token(&mut self) -> Result<()> {
        self.tokens = self
            .tokens
            .checked_add(1)
            .ok_or_else(|| invalid("glossary XML token count overflow"))?;
        if self.tokens > MAX_DOM_TOKENS {
            return Err(invalid("glossary XML token limit exceeded"));
        }
        Ok(())
    }

    fn node(&mut self) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| invalid("glossary XML node count overflow"))?;
        if self.nodes > MAX_NODES {
            return Err(invalid("glossary XML node limit exceeded"));
        }
        self.charge(size_of::<Node>() + size_of::<NamespaceFrame>())
    }

    fn node_names(&mut self, qualified: &str, local: &str) -> Result<()> {
        self.charge(
            qualified
                .len()
                .checked_add(local.len())
                .ok_or_else(|| invalid("glossary node-name budget overflow"))?,
        )
    }

    fn binding(&mut self, prefix: &str, namespace: &str) -> Result<()> {
        self.charge(
            size_of::<(String, Arc<str>)>()
                .checked_add(prefix.len())
                .and_then(|bytes| bytes.checked_add(namespace.len()))
                .ok_or_else(|| invalid("glossary namespace budget overflow"))?,
        )
    }

    fn attribute(&mut self, qualified: &str, local: &str, value: &str) -> Result<()> {
        self.attributes = self
            .attributes
            .checked_add(1)
            .ok_or_else(|| invalid("glossary XML attribute count overflow"))?;
        if self.attributes > MAX_DOM_ATTRIBUTES {
            return Err(invalid("glossary XML aggregate attribute limit exceeded"));
        }
        self.charge(
            size_of::<Attr>()
                .checked_add(qualified.len())
                .and_then(|bytes| bytes.checked_add(local.len()))
                .and_then(|bytes| bytes.checked_add(value.len()))
                .ok_or_else(|| invalid("glossary attribute budget overflow"))?,
        )
    }

    fn content(&mut self, content: &Content) -> Result<()> {
        self.contents = self
            .contents
            .checked_add(1)
            .ok_or_else(|| invalid("glossary XML content count overflow"))?;
        if self.contents > MAX_DOM_CONTENT {
            return Err(invalid("glossary XML content token limit exceeded"));
        }
        let owned = match content {
            Content::Node(_) => 0,
            Content::Text(value) | Content::CData(value) | Content::Comment(value) => value.len(),
        };
        self.charge(
            size_of::<Content>()
                .checked_add(owned)
                .ok_or_else(|| invalid("glossary content budget overflow"))?,
        )
    }
}

#[derive(Default)]
pub(in crate::glossary) struct NamespaceResolver {
    pub(in crate::glossary) active: HashMap<String, Vec<Arc<str>>>,
}

impl NamespaceResolver {
    fn push(&mut self, bindings: &[(String, Arc<str>)]) -> Result<()> {
        self.active
            .try_reserve(bindings.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary active namespace resolver",
                source,
            })?;
        for (prefix, namespace) in bindings {
            if let Some(values) = self.active.get_mut(prefix) {
                values.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "glossary active namespace stack",
                    source,
                })?;
                values.push(Arc::clone(namespace));
            } else {
                let mut values = Vec::new();
                values
                    .try_reserve_exact(1)
                    .map_err(|source| Error::Allocation {
                        resource: "glossary active namespace stack",
                        source,
                    })?;
                values.push(Arc::clone(namespace));
                self.active.insert(prefix.clone(), values);
            }
        }
        Ok(())
    }

    fn pop(&mut self, bindings: &[(String, Arc<str>)]) -> Result<()> {
        for (prefix, _) in bindings.iter().rev() {
            let remove = {
                let values = self
                    .active
                    .get_mut(prefix)
                    .ok_or_else(|| invalid("glossary namespace stack is inconsistent"))?;
                values
                    .pop()
                    .ok_or_else(|| invalid("glossary namespace stack is empty"))?;
                values.is_empty()
            };
            if remove {
                self.active.remove(prefix);
            }
        }
        Ok(())
    }

    fn resolve(&self, prefix: &str) -> Result<Arc<str>> {
        if prefix == "xml" {
            return Ok(Arc::from("http://www.w3.org/XML/1998/namespace"));
        }
        self.active
            .get(prefix)
            .and_then(|values| values.last())
            .map(Arc::clone)
            .ok_or_else(|| invalid(format!("unbound prefix '{prefix}'")))
    }
}

pub(in crate::glossary) fn validate_xml_value(value: &str, context: &str) -> Result<()> {
    if value.chars().all(xml_char) {
        Ok(())
    } else {
        Err(invalid(format!(
            "glossary XML {context} contains a character forbidden by XML 1.0"
        )))
    }
}

pub(in crate::glossary) fn parse_dom(xml: &[u8]) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut rd = Reader::from_reader(xml);
    let mut budget = DomBudget::default();
    let mut resolver = NamespaceResolver::default();
    let mut stack: Vec<Node> = Vec::new();
    stack
        .try_reserve_exact(MAX_DEPTH)
        .map_err(|source| Error::Allocation {
            resource: "glossary XML stack",
            source,
        })?;
    let mut root = None;
    loop {
        let d = rd.decoder();
        let event = rd.read_event().map_err(xml_error)?;
        if !matches!(&event, Event::Eof) {
            budget.token()?;
        }
        match event {
            Event::Start(e) => {
                budget.node()?;
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("glossary XML resource limit exceeded"));
                }
                stack.push(make(&e, d, &stack, &mut resolver, &mut budget)?);
            },
            Event::Empty(e) => {
                budget.node()?;
                let n = make(&e, d, &stack, &mut resolver, &mut budget)?;
                resolver.pop(&n.bindings.local)?;
                attach(&mut stack, &mut root, n, &mut budget)?;
            },
            Event::End(_) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected glossary closing element"))?;
                resolver.pop(&n.bindings.local)?;
                attach(&mut stack, &mut root, n, &mut budget)?;
            },
            Event::Text(t) => {
                let v = t.decode().map_err(xml_error)?.into_owned();
                validate_xml_value(&v, "text")?;
                if let Some(n) = stack.last_mut() {
                    push_content(n, Content::Text(v), &mut budget)?
                } else if !v.trim().is_empty() {
                    return Err(invalid("text outside glossary root"));
                }
            },
            Event::CData(t) => {
                let value = t.decode().map_err(xml_error)?.into_owned();
                validate_xml_value(&value, "CDATA")?;
                if let Some(n) = stack.last_mut() {
                    push_content(n, Content::CData(value), &mut budget)?
                } else {
                    return Err(invalid("CDATA outside glossary root"));
                }
            },
            Event::Comment(t) => {
                let value = t.decode().map_err(xml_error)?.into_owned();
                validate_xml_value(&value, "comment")?;
                if value.contains("--") || value.ends_with('-') {
                    return Err(invalid("glossary XML comment has an invalid lexical form"));
                }
                if let Some(n) = stack.last_mut() {
                    push_content(n, Content::Comment(value), &mut budget)?
                }
            },
            Event::GeneralRef(t) => {
                let value =
                    litchi_ooxml_common::xml::decode_xml_reference(&t).map_err(xml_error)?;
                validate_xml_value(&value, "character reference")?;
                if let Some(n) = stack.last_mut() {
                    push_content(n, Content::Text(value), &mut budget)?
                } else {
                    return Err(invalid("entity outside glossary root"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Decl(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated glossary XML"));
    }
    root.ok_or_else(|| invalid("missing glossary root"))
}
pub(in crate::glossary) fn make(
    e: &BytesStart<'_>,
    d: Decoder,
    stack: &[Node],
    resolver: &mut NamespaceResolver,
    budget: &mut DomBudget,
) -> Result<Node> {
    let q = std::str::from_utf8(e.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    validate_xml_value(&q, "qualified name")?;
    let parent = stack.last().map(|node| Arc::clone(&node.bindings));
    let mut raw = Vec::new();
    for a in e.attributes().with_checks(true) {
        if raw.len() >= MAX_VALUES {
            return Err(invalid("glossary XML attribute limit exceeded"));
        }
        raw.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "glossary XML attributes",
            source,
        })?;
        let a = a.map_err(xml_error)?;
        let key = std::str::from_utf8(a.key.as_ref())
            .map_err(xml_error)?
            .to_string();
        let value = a
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
            .map_err(xml_error)?
            .into_owned();
        validate_xml_value(&key, "attribute name")?;
        validate_xml_value(&value, "attribute value")?;
        raw.push((key, value));
    }
    let mut local_bindings: Vec<(String, Arc<str>)> = Vec::new();
    for (k, v) in &raw {
        if k == "xmlns" || k.starts_with("xmlns:") {
            let key = k.strip_prefix("xmlns:").unwrap_or("").to_string();
            local_bindings
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "glossary namespace bindings",
                    source,
                })?;
            local_bindings.push((key, Arc::from(v.as_str())));
        }
    }
    let declaration_count = parent
        .as_ref()
        .map_or(0, |frame| frame.declaration_count)
        .checked_add(local_bindings.len())
        .ok_or_else(|| invalid("glossary namespace binding count overflow"))?;
    if declaration_count > MAX_VALUES {
        return Err(invalid("glossary namespace binding limit exceeded"));
    }
    let bindings = Arc::new(NamespaceFrame {
        parent,
        local: local_bindings,
        declaration_count,
    });
    for (prefix, namespace) in &bindings.local {
        budget.binding(prefix, namespace)?;
    }
    resolver.push(&bindings.local)?;
    let (pr, lo) = split(&q)?;
    let local = lo.to_string();
    budget.node_names(&q, &local)?;
    let ns = resolver.resolve(pr)?;
    let mut attrs = Vec::new();
    for (q, v) in raw {
        if q == "xmlns" || q.starts_with("xmlns:") {
            continue;
        }
        let (pr, lo) = split(&q)?;
        let ans = if pr.is_empty() {
            Arc::from("")
        } else {
            resolver.resolve(pr)?
        };
        let local = lo.to_string();
        budget.attribute(&q, &local, &v)?;
        attrs.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "glossary typed attributes",
            source,
        })?;
        attrs.push(Attr {
            q,
            ns: ans,
            l: local,
            v,
        });
    }
    Ok(Node {
        q,
        ns,
        l: local,
        attrs,
        bindings,
        content: Vec::new(),
    })
}
pub(in crate::glossary) fn attach(
    stack: &mut [Node],
    root: &mut Option<Node>,
    n: Node,
    budget: &mut DomBudget,
) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        push_content(p, Content::Node(n), budget)?
    } else if root.replace(n).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

pub(in crate::glossary) fn push_content(
    node: &mut Node,
    content: Content,
    budget: &mut DomBudget,
) -> Result<()> {
    budget.content(&content)?;
    node.content
        .try_reserve(1)
        .map_err(|source| Error::Allocation {
            resource: "glossary XML content",
            source,
        })?;
    node.content.push(content);
    Ok(())
}

#[derive(Default)]
pub(in crate::glossary) struct XmlSize {
    pub(in crate::glossary) bytes: usize,
}

pub(in crate::glossary) trait XmlSink {
    fn push_str(&mut self, value: &str) -> Result<()>;
    fn push_char(&mut self, value: char) -> Result<()>;
}

impl XmlSink for XmlSize {
    fn push_str(&mut self, value: &str) -> Result<()> {
        self.bytes = self
            .bytes
            .checked_add(value.len())
            .ok_or_else(|| invalid("glossary XML output size overflow"))?;
        if self.bytes > MAX {
            return Err(invalid("serialized glossary document exceeds 32 MiB"));
        }
        Ok(())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        self.bytes = self
            .bytes
            .checked_add(value.len_utf8())
            .ok_or_else(|| invalid("glossary XML output size overflow"))?;
        if self.bytes > MAX {
            return Err(invalid("serialized glossary document exceeds 32 MiB"));
        }
        Ok(())
    }
}

impl XmlSink for String {
    fn push_str(&mut self, value: &str) -> Result<()> {
        String::push_str(self, value);
        Ok(())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        String::push(self, value);
        Ok(())
    }
}

pub(in crate::glossary) fn node_xml(n: &Node, s: bool) -> Result<Vec<u8>> {
    let mut size = XmlSize::default();
    node_write(&mut size, n, s)?;
    let mut x = String::new();
    x.try_reserve_exact(size.bytes)
        .map_err(|source| Error::Allocation {
            resource: "glossary opaque XML",
            source,
        })?;
    node_write(&mut x, n, s)?;
    if x.len() != size.bytes {
        return Err(invalid("glossary opaque XML plan did not match output"));
    }
    Ok(x.into_bytes())
}
pub(in crate::glossary) fn node_write<X: XmlSink>(x: &mut X, n: &Node, s: bool) -> Result<()> {
    node_write_inner(x, n, s, true)
}

pub(in crate::glossary) fn node_write_inner<X: XmlSink>(
    x: &mut X,
    n: &Node,
    s: bool,
    root: bool,
) -> Result<()> {
    x.push_char('<')?;
    x.push_str(&n.q)?;
    if root {
        for (prefix, namespace) in namespace_scope(&n.bindings)? {
            write_namespace_binding(x, prefix, namespace, s)?;
        }
    } else {
        for (prefix, namespace) in &n.bindings.local {
            write_namespace_binding(x, prefix, namespace, s)?;
        }
    }
    for a in &n.attrs {
        x.push_char(' ')?;
        x.push_str(&a.q)?;
        x.push_str("=\"")?;
        esc(x, &a.v)?;
        x.push_char('"')?;
    }
    if n.content.is_empty() {
        x.push_str("/>")?;
        return Ok(());
    }
    x.push_char('>')?;
    for c in &n.content {
        match c {
            Content::Node(n) => node_write_inner(x, n, s, false)?,
            Content::Text(v) => text(x, v)?,
            Content::CData(v) => {
                if v.contains('\r') || v.contains("]]>") {
                    text(x, v)?;
                } else {
                    x.push_str("<![CDATA[")?;
                    x.push_str(v)?;
                    x.push_str("]]>")?;
                }
            },
            Content::Comment(v) => {
                x.push_str("<!--")?;
                x.push_str(v)?;
                x.push_str("-->")?;
            },
        }
    }
    x.push_str("</")?;
    x.push_str(&n.q)?;
    x.push_char('>')?;
    Ok(())
}

pub(in crate::glossary) fn namespace_scope(bindings: &NamespaceFrame) -> Result<Vec<(&str, &str)>> {
    let mut frames = Vec::new();
    frames
        .try_reserve_exact(MAX_DEPTH)
        .map_err(|source| Error::Allocation {
            resource: "glossary namespace frames",
            source,
        })?;
    let mut current = Some(bindings);
    while let Some(frame) = current {
        frames.push(frame);
        current = frame.parent.as_deref();
    }
    frames.reverse();

    let mut result = Vec::new();
    result
        .try_reserve(bindings.declaration_count)
        .map_err(|source| Error::Allocation {
            resource: "glossary namespace scope",
            source,
        })?;
    let mut positions = HashMap::new();
    positions
        .try_reserve(bindings.declaration_count)
        .map_err(|source| Error::Allocation {
            resource: "glossary namespace scope index",
            source,
        })?;
    for frame in frames {
        for (prefix, namespace) in &frame.local {
            if let Some(index) = positions.get(prefix.as_str()).copied() {
                result[index] = (prefix.as_str(), namespace.as_ref());
            } else {
                positions.insert(prefix.as_str(), result.len());
                result.push((prefix.as_str(), namespace.as_ref()));
            }
        }
    }
    Ok(result)
}

pub(in crate::glossary) fn write_namespace_binding<X: XmlSink>(
    x: &mut X,
    prefix: &str,
    namespace: &str,
    strict: bool,
) -> Result<()> {
    if prefix.is_empty() {
        x.push_str(" xmlns=\"")?
    } else {
        x.push_str(" xmlns:")?;
        x.push_str(prefix)?;
        x.push_str("=\"")?
    }
    esc(x, mapns(namespace, strict))?;
    x.push_char('"')
}
pub(in crate::glossary) fn mapns(v: &str, s: bool) -> &str {
    if s {
        match v {
            W => WS,
            R => RS,
            _ => v,
        }
    } else {
        match v {
            WS => W,
            RS => R,
            _ => v,
        }
    }
}
pub(in crate::glossary) fn esc<X: XmlSink>(x: &mut X, v: &str) -> Result<()> {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;")?,
            '<' => x.push_str("&lt;")?,
            '"' => x.push_str("&quot;")?,
            '\r' => x.push_str("&#xD;")?,
            '\n' => x.push_str("&#xA;")?,
            '\t' => x.push_str("&#x9;")?,
            _ => x.push_char(c)?,
        }
    }
    Ok(())
}
pub(in crate::glossary) fn text<X: XmlSink>(x: &mut X, v: &str) -> Result<()> {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;")?,
            '<' => x.push_str("&lt;")?,
            '>' => x.push_str("&gt;")?,
            '\r' => x.push_str("&#xD;")?,
            _ => x.push_char(c)?,
        }
    }
    Ok(())
}
