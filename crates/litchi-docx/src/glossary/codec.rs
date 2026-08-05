//! Bounded Strict/Transitional glossary XML model and codec.

use super::graph::*;
use super::model::*;
use super::*;
use std::mem::size_of;
pub fn read(xml: &[u8]) -> Result<(Catalog, Conformance)> {
    if xml.len() > MAX {
        return Err(invalid("glossary document exceeds 32 MiB"));
    }
    let original = parse_dom(xml)?;
    let original_conformance = Conformance::from_word(original.ns.as_ref())?;
    let producer_entries = extract_producer_entries(original, original_conformance)?;
    let limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: MAX,
        max_output_bytes: MAX,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: MAX_VALUES,
        max_directive_tokens: MAX_VALUES,
        max_choices_per_alternate: MAX_VALUES,
    };
    let xml = litchi_ooxml_common::mce::process_markup_compatibility(
        xml,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &limits,
    )?
    .xml;
    if xml.len() > MAX {
        return Err(invalid("processed glossary document exceeds 32 MiB"));
    }
    let root = parse_dom(xml.as_ref())?;
    let conformance = Conformance::from_word(root.ns.as_ref())?;
    if conformance != original_conformance {
        return Err(invalid(
            "MCE preprocessing changed the glossary conformance family",
        ));
    }
    validate_word_dialect(&root, conformance)?;
    let mut catalog = project(&root)?;
    drop(root);
    attach_producer_entries(&mut catalog, producer_entries);
    catalog.rebuild_state()?;
    Ok((catalog, conformance))
}

/// Serialize a catalog canonically in the selected dialect.
pub fn write(value: &Catalog, conformance: Conformance) -> Result<Vec<u8>> {
    let plan = plan_write(value, conformance)?;
    let mut xml = String::new();
    xml.try_reserve_exact(plan.bytes)
        .map_err(|source| Error::Allocation {
            resource: "glossary XML",
            source,
        })?;
    emit_catalog(&mut xml, value, conformance, &plan)?;
    if xml.len() != plan.bytes {
        return Err(invalid("glossary XML write plan did not match output"));
    }
    Ok(xml.into_bytes())
}

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

pub(in crate::glossary) fn project(n: &Node) -> Result<Catalog> {
    expect(n, "glossaryDocument")?;
    noattrs(n)?;
    let c = kids(n)?;
    let mut out = Catalog::default();
    let mut semantic_bytes = 0usize;
    let mut index = 0;
    if let Some(background) = c.first().filter(|node| node.l == "background") {
        expect(background, "background")?;
        let xml = node_xml(background, false)?;
        add_semantic_xml_bytes(&mut semantic_bytes, xml.len())?;
        out.background = Some(xml);
        index = 1;
    }
    if index < c.len() {
        let parts = c
            .get(index)
            .copied()
            .ok_or_else(|| invalid("glossary child index is out of bounds"))?;
        expect(parts, "docParts")?;
        noattrs(parts)?;
        for p in kids(parts)? {
            if out.entries.len() >= MAX_PARTS {
                return Err(invalid("glossary entry limit exceeded"));
            }
            out.entries.push(parse_entry(p, &mut semantic_bytes)?);
        }
        index += 1;
    }
    if index != c.len() {
        return Err(invalid("unexpected glossary root child"));
    }
    validate_catalog_fields(&out)?;
    Ok(out)
}

pub(in crate::glossary) fn add_semantic_xml_bytes(total: &mut usize, bytes: usize) -> Result<()> {
    *total = total
        .checked_add(bytes)
        .ok_or_else(|| invalid("glossary semantic XML size overflow"))?;
    if *total > MAX {
        return Err(invalid(
            "glossary semantic XML exceeds the 32 MiB aggregate limit",
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn extract_producer_entries(
    root: Node,
    conformance: Conformance,
) -> Result<Option<Vec<ProducerEntry>>> {
    let mut inherited_mce = Vec::new();
    merge_mce_attributes(&mut inherited_mce, &root)?;
    let mut doc_parts = None;
    for content in root.content {
        match content {
            Content::Node(node)
                if node.ns.as_ref() == conformance.word() && node.l == "docParts" =>
            {
                if doc_parts.replace(node).is_some() {
                    return Ok(None);
                }
            },
            Content::Text(text) if text.trim().is_empty() => {},
            Content::Comment(_) | Content::Node(_) => {},
            Content::Text(_) | Content::CData(_) => return Ok(None),
        }
    }
    let Some(doc_parts) = doc_parts else {
        return Ok(None);
    };
    merge_mce_attributes(&mut inherited_mce, &doc_parts)?;
    let mut producer_nodes = Vec::new();
    for content in doc_parts.content {
        match content {
            Content::Node(node)
                if node.ns.as_ref() == conformance.word() && node.l == "docPart" =>
            {
                if producer_nodes.len() >= MAX_PARTS {
                    return Err(invalid("glossary producer entry limit exceeded"));
                }
                producer_nodes
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "glossary producer nodes",
                        source,
                    })?;
                producer_nodes.push(node);
            },
            Content::Text(text) if text.trim().is_empty() => {},
            Content::Comment(_) => {},
            Content::Node(_) | Content::Text(_) | Content::CData(_) => return Ok(None),
        }
    }
    let mut snapshots = Vec::new();
    snapshots
        .try_reserve_exact(producer_nodes.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary producer entries",
            source,
        })?;
    let mut total_bytes = 0usize;
    for mut node in producer_nodes {
        let mut existing = HashSet::new();
        existing
            .try_reserve(node.attrs.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary producer MCE attribute index",
                source,
            })?;
        existing.extend(
            node.attrs
                .iter()
                .filter(|attribute| attribute.ns.as_ref() == MC)
                .map(|attribute| attribute.l.as_str()),
        );
        let mut missing = Vec::new();
        missing
            .try_reserve(inherited_mce.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary missing MCE attributes",
                source,
            })?;
        missing.extend(
            inherited_mce
                .iter()
                .filter(|attribute| !existing.contains(attribute.l.as_str())),
        );
        drop(existing);
        node.attrs
            .try_reserve(missing.len())
            .map_err(|source| Error::Allocation {
                resource: "glossary inherited MCE attributes",
                source,
            })?;
        for attribute in missing {
            node.attrs.push(attribute.clone());
        }
        let xml = String::from_utf8(node_xml(&node, conformance == Conformance::Strict)?)
            .map_err(xml_error)?;
        total_bytes = total_bytes
            .checked_add(xml.len())
            .ok_or_else(|| invalid("glossary producer snapshot size overflow"))?;
        if total_bytes > MAX {
            return Err(invalid(
                "glossary producer snapshots exceed the 32 MiB aggregate limit",
            ));
        }
        let refs = relationship_references(&node)?;
        snapshots.push(ProducerEntry {
            conformance,
            xml: Arc::from(xml),
            refs,
        });
    }
    Ok(Some(snapshots))
}

pub(in crate::glossary) fn attach_producer_entries(
    catalog: &mut Catalog,
    producer_entries: Option<Vec<ProducerEntry>>,
) {
    let Some(producer_entries) = producer_entries else {
        return;
    };
    if producer_entries.len() != catalog.entries.len() {
        return;
    }
    for (entry, producer) in catalog.entries.iter_mut().zip(producer_entries) {
        entry.producer = Some(Box::new(producer));
    }
}

pub(in crate::glossary) fn merge_mce_attributes(output: &mut Vec<Attr>, node: &Node) -> Result<()> {
    let incoming = node
        .attrs
        .iter()
        .filter(|attribute| attribute.ns.as_ref() == MC)
        .count();
    let capacity = output
        .len()
        .checked_add(incoming)
        .ok_or_else(|| invalid("glossary MCE attribute count overflow"))?;
    let mut positions = HashMap::new();
    positions
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "glossary MCE attribute index",
            source,
        })?;
    for (index, attribute) in output.iter().enumerate() {
        positions.insert(attribute.l.clone(), index);
    }
    output
        .try_reserve(incoming)
        .map_err(|source| Error::Allocation {
            resource: "glossary MCE attribute scope",
            source,
        })?;
    for attribute in node
        .attrs
        .iter()
        .filter(|attribute| attribute.ns.as_ref() == MC)
    {
        if let Some(index) = positions.get(attribute.l.as_str()).copied() {
            output[index] = attribute.clone();
        } else {
            positions.insert(attribute.l.clone(), output.len());
            output.push(attribute.clone());
        }
    }
    Ok(())
}

pub(in crate::glossary) fn parse_entry(n: &Node, semantic_bytes: &mut usize) -> Result<Entry> {
    expect(n, "docPart")?;
    noattrs(n)?;
    let c = kids(n)?;
    if c.len() > 2 {
        return Err(invalid("docPart has too many children"));
    }
    let mut props = None;
    let mut body = None;
    let mut i = 0;
    if let Some(properties) = c.first().filter(|node| node.l == "docPartPr") {
        props = Some(parse_props(properties)?);
        i = 1;
    }
    if i < c.len() {
        let body_node = c
            .get(i)
            .copied()
            .ok_or_else(|| invalid("document-part child index is out of bounds"))?;
        expect(body_node, "docPartBody")?;
        let xml = node_xml(body_node, false)?;
        add_semantic_xml_bytes(semantic_bytes, xml.len())?;
        body = Some(xml);
        i += 1;
    }
    if i != c.len() {
        return Err(invalid("invalid docPart child order"));
    }
    Ok(Entry {
        props,
        body,
        producer: None,
        sizes: [0; 2],
        refs: Arc::from([]),
        lineage: None,
    })
}
pub(in crate::glossary) fn parse_props(n: &Node) -> Result<Props> {
    expect(n, "docPartPr")?;
    noattrs(n)?;
    let mut name_node = None;
    let mut style_node = None;
    let mut category_node = None;
    let mut types_node = None;
    let mut behaviors_node = None;
    let mut description_node = None;
    let mut id_node = None;
    for c in kids(n)? {
        expect_w(c)?;
        match c.l.as_str() {
            "name" => name_node = Some(c),
            "style" => style_node = Some(c),
            "category" => category_node = Some(c),
            "types" => types_node = Some(c),
            "behaviors" => behaviors_node = Some(c),
            "description" => description_node = Some(c),
            "guid" => id_node = Some(c),
            _ => return Err(invalid("unexpected docPartPr child")),
        }
    }
    let name = name_node
        .map(|name_node| {
            let value = wattr_get(name_node, "val")?
                .ok_or_else(|| invalid("building-block name is missing w:val"))?;
            bounded(&value)?;
            let name = Name {
                key: Arc::from(name_key(&value)?),
                value,
                decorated: onoff(name_node, "decorated")?,
            };
            only_w(name_node, &["val", "decorated"])?;
            leaf(name_node)?;
            Ok::<Name, Error>(name)
        })
        .transpose()?;

    let style = style_node.map(wval).transpose()?;
    let category = category_node
        .map(|node| {
            noattrs(node)?;
            let children = kids(node)?;
            let [category_name, gallery] = children.as_slice() else {
                return Err(invalid("category requires name then gallery"));
            };
            expect(category_name, "name")?;
            expect(gallery, "gallery")?;
            Ok::<Category, Error>(Category {
                name: wval(category_name)?,
                gallery: Gallery::new(wval(gallery)?)?,
            })
        })
        .transpose()?;
    let (all_kinds, kinds) = if let Some(node) = types_node {
        let all = onoff(node, "all")?;
        only_w(node, &["all"])?;
        let mut values = Kind::empty();
        let children = kids(node)?;
        for (count, child) in children.into_iter().enumerate() {
            if count >= MAX_VALUES {
                return Err(invalid("type limit exceeded"));
            }
            expect(child, "type")?;
            let value = parse_type(&wval(child)?)?;
            values.insert(value);
        }
        (all, values)
    } else {
        (None, Kind::empty())
    };
    let inserts = if let Some(node) = behaviors_node {
        noattrs(node)?;
        let mut values = Insert::empty();
        let children = kids(node)?;
        if children.is_empty() {
            return Err(invalid(
                "document-part behaviors requires at least one behavior",
            ));
        }
        for (count, child) in children.into_iter().enumerate() {
            if count >= MAX_VALUES {
                return Err(invalid("behavior limit exceeded"));
            }
            expect(child, "behavior")?;
            let value = parse_behavior(&wval(child)?)?;
            values.insert(value);
        }
        values
    } else {
        Insert::empty()
    };
    let description = description_node.map(wval).transpose()?;
    let id = id_node
        .map(|node| {
            let value = wattr_get(node, "val")?;
            only_w(node, &["val"])?;
            leaf(node)?;
            value.map(Id::new).transpose()
        })
        .transpose()?
        .flatten();
    Ok(Props {
        name,
        style,
        category,
        all_kinds,
        kinds,
        inserts,
        description,
        id,
    })
}

pub(in crate::glossary) struct WritePlan {
    pub(in crate::glossary) background: Option<Node>,
    pub(in crate::glossary) bodies: Vec<Option<Node>>,
    pub(in crate::glossary) producer_entries: Vec<Option<Arc<str>>>,
    pub(in crate::glossary) bytes: usize,
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

pub(in crate::glossary) fn entry_analysis(entry: &Entry) -> Result<EntryAnalysis> {
    validate_entry_fields(entry)?;
    let body = entry
        .body
        .as_deref()
        .map(|xml| prepare_opaque(xml, "docPartBody"))
        .transpose()?;
    let refs = body
        .as_ref()
        .map(relationship_references)
        .transpose()?
        .unwrap_or_else(|| Arc::from([]));
    let mut sizes = [0usize; 2];
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let mut size = XmlSize::default();
        if let Some(producer) = entry
            .producer
            .as_deref()
            .filter(|producer| producer.conformance == conformance)
        {
            size.push_str(&producer.xml)?;
        } else {
            write_entry(&mut size, entry, body.as_ref(), conformance)?;
        }
        sizes[conformance.index()] = size.bytes;
    }
    Ok(EntryAnalysis { sizes, refs })
}

pub(in crate::glossary) fn background_analysis(background: Option<&[u8]>) -> Result<EntryAnalysis> {
    let mut sizes = [0usize; 2];
    let Some(background) = background else {
        return Ok(EntryAnalysis {
            sizes,
            refs: Arc::from([]),
        });
    };
    let node = prepare_opaque(background, "background")?;
    let refs = relationship_references(&node)?;
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let mut size = XmlSize::default();
        node_write(&mut size, &node, conformance == Conformance::Strict)?;
        sizes[conformance.index()] = size.bytes;
    }
    Ok(EntryAnalysis { sizes, refs })
}

pub(in crate::glossary) fn add_sizes(left: [usize; 2], right: [usize; 2]) -> Result<[usize; 2]> {
    Ok([
        left[0]
            .checked_add(right[0])
            .ok_or_else(|| invalid("Transitional glossary size overflow"))?,
        left[1]
            .checked_add(right[1])
            .ok_or_else(|| invalid("Strict glossary size overflow"))?,
    ])
}

pub(in crate::glossary) fn replace_sizes(
    total: [usize; 2],
    old: [usize; 2],
    new: [usize; 2],
) -> Result<[usize; 2]> {
    Ok([
        total[0]
            .checked_sub(old[0])
            .and_then(|value| value.checked_add(new[0]))
            .ok_or_else(|| invalid("Transitional glossary replacement size overflow"))?,
        total[1]
            .checked_sub(old[1])
            .and_then(|value| value.checked_add(new[1]))
            .ok_or_else(|| invalid("Strict glossary replacement size overflow"))?,
    ])
}

pub(in crate::glossary) fn validate_catalog_sizes(
    entry_bytes: [usize; 2],
    background_bytes: [usize; 2],
    entry_count: usize,
) -> Result<()> {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let mut size = XmlSize::default();
        write_catalog_open(&mut size, conformance)?;
        if entry_count != 0 {
            size.push_str("<w:docParts></w:docParts>")?;
        }
        size.push_str("</w:glossaryDocument>")?;
        let index = conformance.index();
        let bytes = size
            .bytes
            .checked_add(background_bytes[index])
            .and_then(|bytes| bytes.checked_add(entry_bytes[index]))
            .ok_or_else(|| invalid("glossary catalog size overflow"))?;
        if bytes > MAX {
            return Err(invalid("serialized glossary document exceeds 32 MiB"));
        }
    }
    Ok(())
}

pub(in crate::glossary) fn write_catalog_open<X: XmlSink>(
    x: &mut X,
    conformance: Conformance,
) -> Result<()> {
    x.push_str(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument xmlns:w=""#,
    )?;
    x.push_str(conformance.word())?;
    x.push_str(r#"" xmlns:r=""#)?;
    x.push_str(conformance.relationships())?;
    x.push_str(r#"">"#)
}

pub(in crate::glossary) fn plan_write(
    value: &Catalog,
    conformance: Conformance,
) -> Result<WritePlan> {
    validate_catalog_fields(value)?;
    let background = value
        .background
        .as_deref()
        .map(|xml| prepare_opaque_for(xml, "background", conformance))
        .transpose()?;
    let mut bodies = Vec::new();
    bodies
        .try_reserve_exact(value.entries.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary XML write plan",
            source,
        })?;
    let mut producer_entries = Vec::new();
    producer_entries
        .try_reserve_exact(value.entries.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary producer write plan",
            source,
        })?;
    for entry in &value.entries {
        let producer = entry
            .producer
            .as_deref()
            .filter(|producer| producer.conformance == conformance);
        producer_entries.push(producer.map(|producer| Arc::clone(&producer.xml)));
        bodies.push(if producer.is_some() {
            None
        } else {
            entry
                .body
                .as_deref()
                .map(|xml| prepare_opaque_for(xml, "docPartBody", conformance))
                .transpose()?
        });
    }
    let mut plan = WritePlan {
        background,
        bodies,
        producer_entries,
        bytes: 0,
    };
    let mut size = XmlSize::default();
    emit_catalog(&mut size, value, conformance, &plan)?;
    plan.bytes = size.bytes;
    Ok(plan)
}

pub(in crate::glossary) fn emit_catalog<X: XmlSink>(
    x: &mut X,
    value: &Catalog,
    conformance: Conformance,
    plan: &WritePlan,
) -> Result<()> {
    write_catalog_open(x, conformance)?;
    match (&value.background, &plan.background) {
        (Some(_), Some(background)) => {
            node_write(x, background, conformance == Conformance::Strict)?;
        },
        (None, None) => {},
        _ => return Err(invalid("glossary background write plan is inconsistent")),
    }
    let mut bodies = plan.bodies.iter();
    let mut producer_entries = plan.producer_entries.iter();
    if !value.entries.is_empty() {
        x.push_str("<w:docParts>")?;
        for entry in &value.entries {
            let body = bodies
                .next()
                .ok_or_else(|| invalid("glossary entry write plan is incomplete"))?;
            let producer = producer_entries
                .next()
                .ok_or_else(|| invalid("glossary producer write plan is incomplete"))?;
            if let Some(producer) = producer {
                if body.is_some() {
                    return Err(invalid("glossary producer write plan has a duplicate body"));
                }
                x.push_str(producer)?;
            } else {
                write_entry(x, entry, body.as_ref(), conformance)?;
            }
        }
        x.push_str("</w:docParts>")?;
    }
    if bodies.next().is_some() {
        return Err(invalid("glossary entry write plan has unused bodies"));
    }
    if producer_entries.next().is_some() {
        return Err(invalid("glossary producer write plan has unused entries"));
    }
    x.push_str("</w:glossaryDocument>")?;
    Ok(())
}

pub(in crate::glossary) fn write_entry<X: XmlSink>(
    x: &mut X,
    e: &Entry,
    body: Option<&Node>,
    conformance: Conformance,
) -> Result<()> {
    x.push_str("<w:docPart>")?;
    if let Some(p) = &e.props {
        write_props(x, p)?;
    }
    match (&e.body, body) {
        (Some(_), Some(body)) => node_write(x, body, conformance == Conformance::Strict)?,
        (None, None) => {},
        _ => return Err(invalid("building-block body write plan is inconsistent")),
    }
    x.push_str("</w:docPart>")?;
    Ok(())
}
pub(in crate::glossary) fn write_props<X: XmlSink>(x: &mut X, p: &Props) -> Result<()> {
    x.push_str("<w:docPartPr>")?;
    if let Some(name) = &p.name {
        x.push_str("<w:name")?;
        wattr(x, "val", name.as_str())?;
        wonoff(x, "decorated", name.decorated())?;
        x.push_str("/>")?;
    }
    if let Some(v) = &p.style {
        leafval(x, "style", v)?;
    }
    if let Some(c) = &p.category {
        x.push_str("<w:category>")?;
        leafval(x, "name", &c.name)?;
        leafval(x, "gallery", c.gallery.as_str())?;
        x.push_str("</w:category>")?;
    }
    if p.all_kinds.is_some() || !p.kinds.is_empty() {
        x.push_str("<w:types")?;
        wonoff(x, "all", p.all_kinds)?;
        x.push_char('>')?;
        for (kind, value) in KIND_VALUES {
            if p.kinds.contains(kind) {
                leafval(x, "type", value)?;
            }
        }
        x.push_str("</w:types>")?;
    }
    if !p.inserts.is_empty() {
        x.push_str("<w:behaviors>")?;
        for (insert, value) in INSERT_VALUES {
            if p.inserts.contains(insert) {
                leafval(x, "behavior", value)?;
            }
        }
        x.push_str("</w:behaviors>")?;
    }
    if let Some(v) = &p.description {
        leafval(x, "description", v)?;
    }
    if let Some(v) = &p.id {
        leafval(x, "guid", v.as_str())?;
    }
    x.push_str("</w:docPartPr>")?;
    Ok(())
}
pub(in crate::glossary) fn leafval<X: XmlSink>(x: &mut X, n: &str, v: &str) -> Result<()> {
    x.push_str("<w:")?;
    x.push_str(n)?;
    wattr(x, "val", v)?;
    x.push_str("/>")
}
pub(in crate::glossary) fn wattr<X: XmlSink>(x: &mut X, n: &str, v: &str) -> Result<()> {
    x.push_str(" w:")?;
    x.push_str(n)?;
    x.push_str("=\"")?;
    esc(x, v)?;
    x.push_char('"')
}
pub(in crate::glossary) fn wonoff<X: XmlSink>(x: &mut X, n: &str, v: Option<bool>) -> Result<()> {
    if let Some(v) = v {
        wattr(x, n, if v { "1" } else { "0" })?;
    }
    Ok(())
}

pub(in crate::glossary) fn validate_catalog_fields(v: &Catalog) -> Result<()> {
    if v.entries.len() > MAX_PARTS {
        return Err(invalid("glossary entry limit exceeded"));
    }
    for e in &v.entries {
        validate_entry_fields(e)?;
    }
    Ok(())
}

pub(in crate::glossary) fn validate_entry_fields(entry: &Entry) -> Result<()> {
    let Some(props) = &entry.props else {
        return Ok(());
    };
    if let Some(name) = &props.name {
        bounded(name.as_str())?;
    }
    for value in [
        props.style.as_deref(),
        props.category.as_ref().map(Category::name),
        props.description.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        bounded(value)?;
    }
    if !Kind::all().contains(props.kinds) || !Insert::all().contains(props.inserts) {
        return Err(invalid("unknown glossary option flag"));
    }
    if props.all_kinds.is_some() && props.kinds.is_empty() {
        return Err(invalid(
            "document-part types requires at least one kind when present",
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn validate_authored_name(entry: &Entry) -> Result<()> {
    let Some(name) = entry.props.as_ref().and_then(|props| props.name.as_ref()) else {
        return Err(invalid(
            "authored building block requires properties and a name",
        ));
    };
    validate_name(name.as_str())
}

pub(in crate::glossary) fn authored_analysis(entry: &Entry) -> Result<EntryAnalysis> {
    validate_authored_name(entry)?;
    entry_analysis(entry)
}

pub(in crate::glossary) fn prepare_opaque(xml: &[u8], local: &str) -> Result<Node> {
    if xml.len() > MAX {
        return Err(invalid(format!("glossary {local} payload exceeds 32 MiB")));
    }
    let root = parse_dom(xml)?;
    expect(&root, local)?;
    Ok(root)
}

pub(in crate::glossary) fn prepare_opaque_for(
    xml: &[u8],
    local: &str,
    conformance: Conformance,
) -> Result<Node> {
    let root = prepare_opaque(xml, local)?;
    if conformance == Conformance::Strict {
        reject_vml(&root)?;
    }
    Ok(root)
}

pub(in crate::glossary) fn reject_vml(node: &Node) -> Result<()> {
    if node.ns.as_ref() == VML
        || node
            .attrs
            .iter()
            .any(|attribute| attribute.ns.as_ref() == VML)
    {
        return Err(invalid("Strict glossary content cannot contain VML"));
    }
    for content in &node.content {
        if let Content::Node(child) = content {
            reject_vml(child)?;
        }
    }
    Ok(())
}

pub(in crate::glossary) fn validate_name(value: &str) -> Result<()> {
    bounded(value)?;
    if value.trim().is_empty() {
        Err(invalid("building-block name cannot be empty"))
    } else {
        Ok(())
    }
}

pub(in crate::glossary) fn validate_raw_part(
    name: &str,
    content_type: &str,
    len: usize,
) -> Result<()> {
    let uri = validate_physical_part(name, content_type, len)?;
    if uri.as_str() == "/word/glossary/document.xml" {
        return Err(invalid(
            "glossary auxiliary part conflicts with the default root",
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn validate_physical_part(
    name: &str,
    content_type: &str,
    len: usize,
) -> Result<PackURI> {
    let uri = PackURI::new(name).map_err(Error::Uri)?;
    if is_signature_part(&uri) || is_reserved_physical_part(&uri) {
        return Err(invalid(format!(
            "'{}' is reserved OPC package infrastructure",
            uri.as_str()
        )));
    }
    ContentType::new(content_type.to_owned())?;
    let media_type = content_type.split(';').next().unwrap_or_default().trim();
    if [
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
        ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE,
        ct::OPC_RELATIONSHIPS,
    ]
    .iter()
    .any(|reserved| media_type.eq_ignore_ascii_case(reserved))
    {
        return Err(invalid(
            "glossary parts cannot use reserved OPC infrastructure content types",
        ));
    }
    if len > MAX_GRAPH_BYTES {
        return Err(invalid("glossary auxiliary part exceeds 256 MiB"));
    }
    Ok(uri)
}

pub(in crate::glossary) fn name_key(value: &str) -> Result<String> {
    bounded(value)?;
    let mut key = String::new();
    key.try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary semantic-name key",
            source,
        })?;
    for character in value.chars().nfd().default_case_fold().nfd() {
        let len = key
            .len()
            .checked_add(character.len_utf8())
            .ok_or_else(|| invalid("glossary semantic-name key size overflow"))?;
        if len > MAX_NAME_KEY {
            return Err(invalid(
                "glossary semantic-name key exceeds the normalized size limit",
            ));
        }
        if key.capacity() < len {
            key.try_reserve(character.len_utf8())
                .map_err(|source| Error::Allocation {
                    resource: "glossary semantic-name key",
                    source,
                })?;
        }
        key.push(character);
    }
    Ok(key)
}

pub(in crate::glossary) fn canonical_id(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 38
        && bytes.first() == Some(&b'{')
        && bytes.last() == Some(&b'}')
        && bytes.iter().enumerate().all(|(index, byte)| match index {
            0 => *byte == b'{',
            9 | 14 | 19 | 24 => *byte == b'-',
            37 => *byte == b'}',
            _ => byte.is_ascii_digit() || (b'A'..=b'F').contains(byte),
        })
}

pub(in crate::glossary) fn valid_ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_char)
}

pub(in crate::glossary) fn is_ncname_start(character: char) -> bool {
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

pub(in crate::glossary) fn is_ncname_char(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00b7}' | '\u{0300}'..='\u{036f}' | '\u{203f}'..='\u{2040}'
        )
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

pub(in crate::glossary) fn kids(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed glossary metadata")),
        }
    }
    Ok(v)
}
pub(in crate::glossary) fn leaf(n: &Node) -> Result<()> {
    if kids(n)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("glossary metadata leaf has children"))
    }
}
pub(in crate::glossary) fn validate_word_dialect(n: &Node, conformance: Conformance) -> Result<()> {
    if matches!(n.ns.as_ref(), W | WS) && n.ns.as_ref() != conformance.word() {
        return Err(invalid(
            "glossary mixes Strict and Transitional WordprocessingML",
        ));
    }
    if matches!(n.ns.as_ref(), R | RS) && n.ns.as_ref() != conformance.relationships() {
        return Err(invalid(
            "glossary mixes Strict and Transitional relationship namespaces",
        ));
    }
    if conformance == Conformance::Strict && n.ns.as_ref() == VML {
        return Err(invalid("Strict glossary content cannot contain VML"));
    }
    for attribute in &n.attrs {
        if matches!(attribute.ns.as_ref(), W | WS) && attribute.ns.as_ref() != conformance.word() {
            return Err(invalid(
                "glossary mixes Strict and Transitional WordprocessingML attributes",
            ));
        }
        if matches!(attribute.ns.as_ref(), R | RS)
            && attribute.ns.as_ref() != conformance.relationships()
        {
            return Err(invalid(
                "glossary mixes Strict and Transitional relationship attributes",
            ));
        }
        if conformance == Conformance::Strict && attribute.ns.as_ref() == VML {
            return Err(invalid(
                "Strict glossary content cannot contain VML attributes",
            ));
        }
    }
    for content in &n.content {
        if let Content::Node(child) = content {
            validate_word_dialect(child, conformance)?;
        }
    }
    Ok(())
}
pub(in crate::glossary) fn expect(n: &Node, l: &str) -> Result<()> {
    if matches!(n.ns.as_ref(), W | WS) && n.l == l {
        Ok(())
    } else {
        Err(invalid(format!("expected WordprocessingML {l}")))
    }
}
pub(in crate::glossary) fn expect_w(n: &Node) -> Result<()> {
    if matches!(n.ns.as_ref(), W | WS) {
        Ok(())
    } else {
        Err(invalid("expected WordprocessingML metadata"))
    }
}
pub(in crate::glossary) fn wval(n: &Node) -> Result<String> {
    let v = wattr_get(n, "val")?.ok_or_else(|| invalid("missing w:val"))?;
    bounded(&v)?;
    only_w(n, &["val"])?;
    leaf(n)?;
    Ok(v)
}
pub(in crate::glossary) fn wattr_get(n: &Node, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if matches!(a.ns.as_ref(), W | WS) && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate WordprocessingML attribute"));
            }
            v = Some(a.v.clone());
        }
    }
    Ok(v)
}
pub(in crate::glossary) fn onoff(n: &Node, l: &str) -> Result<Option<bool>> {
    let value = wattr_get(n, l)?;
    match value.as_deref() {
        None => Ok(None),
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
        Some("on") if n.ns.as_ref() == W => Ok(Some(true)),
        Some("off") if n.ns.as_ref() == W => Ok(Some(false)),
        _ => Err(invalid(format!("invalid on/off attribute '{l}'"))),
    }
}
pub(in crate::glossary) fn only_w(n: &Node, allowed: &[&str]) -> Result<()> {
    for a in &n.attrs {
        if !(matches!(a.ns.as_ref(), W | WS) && allowed.contains(&a.l.as_str())) {
            return Err(invalid(format!("unexpected glossary attribute '{}'", a.q)));
        }
    }
    Ok(())
}
pub(in crate::glossary) fn noattrs(n: &Node) -> Result<()> {
    if n.attrs.is_empty() {
        Ok(())
    } else {
        Err(invalid("unexpected glossary attributes"))
    }
}
pub(in crate::glossary) fn parse_type(v: &str) -> Result<Kind> {
    match v {
        "none" => Ok(Kind::NONE),
        "normal" => Ok(Kind::NORMAL),
        "autoExp" => Ok(Kind::AUTO_EXPAND),
        "toolbar" => Ok(Kind::TOOLBAR),
        "speller" => Ok(Kind::SPELLER),
        "formFld" => Ok(Kind::FORM_FIELD),
        "bbPlcHdr" => Ok(Kind::SDT_PLACEHOLDER),
        _ => Err(invalid(format!("invalid document-part type '{v}'"))),
    }
}
pub(in crate::glossary) const KIND_VALUES: [(Kind, &str); 7] = [
    (Kind::NONE, "none"),
    (Kind::NORMAL, "normal"),
    (Kind::AUTO_EXPAND, "autoExp"),
    (Kind::TOOLBAR, "toolbar"),
    (Kind::SPELLER, "speller"),
    (Kind::FORM_FIELD, "formFld"),
    (Kind::SDT_PLACEHOLDER, "bbPlcHdr"),
];
pub(in crate::glossary) fn parse_behavior(v: &str) -> Result<Insert> {
    match v {
        "content" => Ok(Insert::CONTENT),
        "p" => Ok(Insert::PARAGRAPH),
        "pg" => Ok(Insert::PAGE),
        _ => Err(invalid(format!("invalid insertion behavior '{v}'"))),
    }
}
pub(in crate::glossary) const INSERT_VALUES: [(Insert, &str); 3] = [
    (Insert::CONTENT, "content"),
    (Insert::PARAGRAPH, "p"),
    (Insert::PAGE, "pg"),
];
pub(in crate::glossary) const GALLERIES: &[&str] = &[
    "placeholder",
    "any",
    "default",
    "docParts",
    "coverPg",
    "eq",
    "ftrs",
    "hdrs",
    "pgNum",
    "tbls",
    "watermarks",
    "autoTxt",
    "txtBox",
    "pgNumT",
    "pgNumB",
    "pgNumMargins",
    "tblOfContents",
    "bib",
    "custQuickParts",
    "custCoverPg",
    "custEq",
    "custFtrs",
    "custHdrs",
    "custPgNum",
    "custTbls",
    "custWatermarks",
    "custAutoTxt",
    "custTxtBox",
    "custPgNumT",
    "custPgNumB",
    "custPgNumMargins",
    "custTblOfContents",
    "custBib",
    "custom1",
    "custom2",
    "custom3",
    "custom4",
    "custom5",
];
pub(in crate::glossary) fn split(q: &str) -> Result<(&str, &str)> {
    if let Some((p, l)) = q.split_once(':') {
        if l.is_empty() || l.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((p, l))
    } else {
        Ok(("", q))
    }
}
pub(in crate::glossary) fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING {
        return Err(invalid("glossary metadata string exceeds 1 MiB"));
    }
    if !v.chars().all(xml_char) {
        return Err(invalid(
            "glossary metadata contains a character forbidden by XML 1.0",
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn xml_char(character: char) -> bool {
    matches!(
        character,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
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
pub(in crate::glossary) fn invalid(value: impl Into<String>) -> Error {
    Error::Invalid(value.into())
}
pub(in crate::glossary) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
