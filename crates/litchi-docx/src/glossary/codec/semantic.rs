//! Semantic glossary projection and typed document-part encoding.

use super::super::graph::*;
use super::super::model::*;
use super::super::*;
use super::super::{MAX, MAX_PARTS, MAX_VALUES, MC, VML};
use super::validation::*;
use super::xml::*;

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
