//! Strict, inert SpreadsheetML volatile-dependency records.

use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::Writer;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const NS: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies";
const STRICT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/volatileDependencies";
const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml";
const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
const MAX_TYPES: usize = 64;
const MAX_MAINS: usize = 16_384;
const MAX_TOPICS: usize = 65_536;
const MAX_SUBTOPICS: usize = 262_144;
const MAX_REFERENCES: usize = 1_048_576;
const MAX_TEXT_BYTES: usize = 1_048_576;

/// Namespace family used for the volatile-dependencies XML and workbook relationship.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum VolatileDependenciesConformance {
    #[default]
    Transitional,
    Strict,
}

impl VolatileDependenciesConformance {
    const fn relationship_type(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }

    /// Whether this conformance uses ISO/IEC 29500 Strict namespace URIs.
    pub const fn is_strict(self) -> bool {
        matches!(self, Self::Strict)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum VolatileDependencyType {
    RealTimeData,
    OlapFunctions,
}

#[derive(Clone, Debug, PartialEq)]
pub enum VolatileValue {
    Unspecified(String),
    Boolean(bool),
    Number(f64),
    Error(String),
    String(String),
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VolatileReference {
    pub cell_reference: String,
    pub sheet_id: u32,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileTopic {
    pub value: VolatileValue,
    pub subtopics: Vec<String>,
    pub references: Vec<VolatileReference>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileMain {
    pub first: String,
    pub topics: Vec<VolatileTopic>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileType {
    pub dependency_type: VolatileDependencyType,
    pub mains: Vec<VolatileMain>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VolatileDependencies {
    pub types: Vec<VolatileType>,
    /// Raw `extLst` markup. Its payload is preserved, never interpreted or executed.
    pub extension_list_xml: Option<Vec<u8>>,
}

impl VolatileDependencies {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_PART_BYTES {
            return Err(invalid("volatile-dependencies part exceeds 8 MiB"));
        }
        let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
        if processed.len() > MAX_PART_BYTES {
            return Err(invalid(
                "processed volatile-dependencies part exceeds 8 MiB",
            ));
        }
        parse_processed(processed.as_ref())
    }

    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate_document(self)?;
        let ns = if strict {
            std::str::from_utf8(STRICT_NS).unwrap()
        } else {
            std::str::from_utf8(NS).unwrap()
        };
        let mut out = String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><volTypes xmlns=\"",
        );
        escape_attr(&mut out, ns);
        out.push_str("\">");
        for ty in &self.types {
            out.push_str("<volType type=\"");
            out.push_str(match ty.dependency_type {
                VolatileDependencyType::RealTimeData => "realTimeData",
                VolatileDependencyType::OlapFunctions => "olapFunctions",
            });
            out.push_str("\">");
            for main in &ty.mains {
                out.push_str("<main first=\"");
                escape_attr(&mut out, &main.first);
                out.push_str("\">");
                for topic in &main.topics {
                    out.push_str("<tp");
                    let number = match &topic.value {
                        VolatileValue::Number(v) => Some(v.to_string()),
                        _ => None,
                    };
                    let (kind, value) = match &topic.value {
                        VolatileValue::Unspecified(v) => (None, v.as_str()),
                        VolatileValue::Boolean(v) => (Some("b"), if *v { "1" } else { "0" }),
                        VolatileValue::Number(_) => (Some("n"), number.as_deref().unwrap()),
                        VolatileValue::Error(v) => (Some("e"), v.as_str()),
                        VolatileValue::String(v) => (Some("s"), v.as_str()),
                    };
                    if let Some(kind) = kind {
                        out.push_str(" t=\"");
                        out.push_str(kind);
                        out.push('"');
                    }
                    out.push_str("><v>");
                    escape_text(&mut out, value);
                    out.push_str("</v>");
                    for value in &topic.subtopics {
                        out.push_str("<stp>");
                        escape_text(&mut out, value);
                        out.push_str("</stp>");
                    }
                    for reference in &topic.references {
                        out.push_str("<tr r=\"");
                        escape_attr(&mut out, &reference.cell_reference);
                        out.push_str("\" s=\"");
                        out.push_str(&reference.sheet_id.to_string());
                        out.push_str("\"/>");
                    }
                    out.push_str("</tp>");
                }
                out.push_str("</main>");
            }
            out.push_str("</volType>");
        }
        let mut bytes = out.into_bytes();
        if let Some(ext) = &self.extension_list_xml {
            bytes.extend_from_slice(ext);
        }
        bytes.extend_from_slice(b"</volTypes>");
        if bytes.len() > MAX_PART_BYTES {
            return Err(invalid(
                "serialized volatile-dependencies part exceeds 8 MiB",
            ));
        }
        Ok(bytes)
    }
}

/// Loads the single volatile-dependencies part related to the package workbook.
pub fn load_from_package(package: &OpcPackage) -> Result<Option<VolatileDependencies>> {
    Ok(load_from_package_with_conformance(package)?.map(|(value, _)| value))
}

/// Load volatile-dependencies metadata with the XML/relationship namespace family.
///
/// Dependency records are metadata only: this never contacts RTD servers, opens
/// OLAP connections, or evaluates/recalculates workbook formulas.
pub fn load_from_package_with_conformance(
    package: &OpcPackage,
) -> Result<Option<(VolatileDependencies, VolatileDependenciesConformance)>> {
    let workbook_uri = main_workbook_uri(package)?;
    load_for_workbook(package, &workbook_uri)
}

/// Store caller-authored inert volatile-dependencies metadata in a SpreadsheetML package.
///
/// Existing invalid package graphs are rejected before mutation. The writer only
/// persists the supplied dependency records and never performs RTD, cube, or
/// formula evaluation work.
pub fn store_in_package(
    package: &mut OpcPackage,
    value: &VolatileDependencies,
    conformance: VolatileDependenciesConformance,
) -> Result<()> {
    let xml = value.to_xml(conformance.is_strict())?;
    let workbook_uri = main_workbook_uri(package)?;
    let existing = volatile_dependencies_relationship(package, &workbook_uri)?;

    if let Some(existing) = existing {
        validate_volatile_dependencies_graph(package, &workbook_uri, Some(&existing))?;
        validate_volatile_dependencies_part(package, &existing.part_name)?;
        package.get_part_mut(&existing.part_name)?.set_blob(xml);
        if existing.conformance != conformance {
            let workbook = package.get_part_mut(&workbook_uri)?;
            workbook.rels_mut().remove(&existing.relationship_id);
            workbook.rels_mut().add_relationship(
                conformance.relationship_type().into(),
                existing.target_reference,
                existing.relationship_id,
                false,
            );
        }
    } else {
        validate_volatile_dependencies_graph(package, &workbook_uri, None)?;
        let part_name = next_volatile_dependencies_part_name(package)?;
        let relationship_id = next_volatile_dependencies_relationship_id(package, &workbook_uri)?;
        let target = part_name.relative_ref(workbook_uri.base_uri());
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name,
            CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(&workbook_uri)?
            .rels_mut()
            .add_relationship(
                conformance.relationship_type().into(),
                target,
                relationship_id,
                false,
            );
    }

    let _ = package.clear_digital_signatures();
    Ok(())
}

/// Remove the workbook volatile-dependencies relationship and its unreferenced part.
///
/// No RTD, cube, or formula work is performed. A target retained by another
/// relationship is left in the package.
pub fn remove_from_package(package: &mut OpcPackage) -> Result<bool> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(existing) = volatile_dependencies_relationship(package, &workbook_uri)? else {
        validate_volatile_dependencies_graph(package, &workbook_uri, None)?;
        return Ok(false);
    };
    validate_volatile_dependencies_graph(package, &workbook_uri, Some(&existing))?;
    validate_volatile_dependencies_part(package, &existing.part_name)?;

    package
        .get_part_mut(&workbook_uri)?
        .rels_mut()
        .remove(&existing.relationship_id);
    if !package_part_is_referenced(package, &existing.part_name) {
        package.remove_part(&existing.part_name);
    }
    let _ = package.clear_digital_signatures();
    Ok(true)
}

fn load_for_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<(VolatileDependencies, VolatileDependenciesConformance)>> {
    let Some(relationship) = volatile_dependencies_relationship(package, workbook_uri)? else {
        validate_volatile_dependencies_graph(package, workbook_uri, None)?;
        return Ok(None);
    };
    validate_volatile_dependencies_graph(package, workbook_uri, Some(&relationship))?;
    validate_volatile_dependencies_part(package, &relationship.part_name)?;
    let part = package.get_part(&relationship.part_name)?;
    Ok(Some((
        VolatileDependencies::parse(part.blob())?,
        relationship.conformance,
    )))
}

#[derive(Clone, Debug)]
struct VolatileDependenciesRelationship {
    relationship_id: String,
    part_name: PackURI,
    target_reference: String,
    conformance: VolatileDependenciesConformance,
}

fn volatile_dependencies_relationship(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<VolatileDependenciesRelationship>> {
    let workbook = package.get_part(workbook_uri)?;
    let mut relationships = workbook
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL));
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(invalid(
            "workbook has multiple volatile-dependencies relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid(
            "volatile-dependencies relationship cannot be external",
        ));
    }
    let conformance = if relationship.reltype() == REL {
        VolatileDependenciesConformance::Transitional
    } else {
        VolatileDependenciesConformance::Strict
    };
    Ok(Some(VolatileDependenciesRelationship {
        relationship_id: relationship.r_id().to_string(),
        part_name: relationship.target_partname()?,
        target_reference: relationship.target_ref().to_string(),
        conformance,
    }))
}

fn validate_volatile_dependencies_part(package: &OpcPackage, part_name: &PackURI) -> Result<()> {
    let part = package.get_part(part_name)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "volatile-dependencies part '{part_name}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid(
            "volatile-dependencies part must not have relationships",
        ));
    }
    if part.blob().len() > MAX_PART_BYTES {
        return Err(invalid("volatile-dependencies part exceeds 8 MiB"));
    }
    Ok(())
}

fn validate_volatile_dependencies_graph(
    package: &OpcPackage,
    workbook_uri: &PackURI,
    expected: Option<&VolatileDependenciesRelationship>,
) -> Result<()> {
    validate_volatile_dependencies_part_set(package, expected.map(|value| &value.part_name))?;

    let mut found = 0usize;
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
        {
            if part.partname() != workbook_uri {
                return Err(invalid(
                    "volatile-dependencies relationships may only originate from the workbook",
                ));
            }
            if relationship.is_external() {
                return Err(invalid(
                    "volatile-dependencies relationship cannot be external",
                ));
            }
            let target = relationship.target_partname()?;
            let Some(expected) = expected else {
                return Err(invalid(
                    "workbook has an unexpected volatile-dependencies relationship",
                ));
            };
            if relationship.r_id() != expected.relationship_id || target != expected.part_name {
                return Err(invalid(
                    "volatile-dependencies relationship graph is inconsistent",
                ));
            }
            found += 1;
        }
    }
    if package
        .rels()
        .iter()
        .any(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
    {
        return Err(invalid(
            "volatile-dependencies relationships may not originate from the package root",
        ));
    }
    match (expected, found) {
        (None, 0) | (Some(_), 1) => Ok(()),
        (None, _) => Err(invalid(
            "workbook has an unexpected volatile-dependencies relationship",
        )),
        (Some(_), _) => Err(invalid(
            "workbook volatile-dependencies relationship graph is incomplete",
        )),
    }
}

fn validate_volatile_dependencies_part_set(
    package: &OpcPackage,
    relationship_target: Option<&PackURI>,
) -> Result<()> {
    let part_names = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE)
        .map(|part| part.partname().clone())
        .collect::<Vec<_>>();
    if part_names.len() > 1 {
        return Err(invalid(
            "package contains more than one volatile-dependencies part",
        ));
    }
    match (relationship_target, part_names.as_slice()) {
        (None, []) => Ok(()),
        (None, _) => Err(invalid(
            "package contains a volatile-dependencies part without a workbook relationship",
        )),
        (Some(_), []) => Err(invalid(
            "workbook volatile-dependencies relationship targets a missing part",
        )),
        (Some(target), [part_name]) if part_name == target => Ok(()),
        (Some(_), _) => Err(invalid(
            "workbook volatile-dependencies relationship does not target the volatile-dependencies part",
        )),
    }
}

fn main_workbook_uri(package: &OpcPackage) -> Result<PackURI> {
    use litchi_opc::constants::content_type as ct;

    let workbook = package.main_document_part()?;
    if !matches!(
        workbook.content_type(),
        ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_SHEET_MACRO_MAIN
            | ct::SML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(invalid(format!(
            "main document part '{}' is not an XML workbook",
            workbook.partname()
        )));
    }
    Ok(workbook.partname().clone())
}

fn next_volatile_dependencies_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/volatileDependencies.xml".to_string()
        } else {
            format!("/xl/volatileDependencies{suffix}.xml")
        };
        let candidate = PackURI::new(&name)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free volatile-dependencies part name"))
}

fn next_volatile_dependencies_relationship_id(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<String> {
    let relationships = package.get_part(workbook_uri)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdVolatileDependencies{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free volatile-dependencies relationship ID"))
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part_name| part_name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|part_name| part_name == *target)
    })
}

#[derive(Clone, Copy)]
enum Context {
    Root,
    Type(usize),
    Main(usize, usize),
    Topic(usize, usize, usize),
    Value(usize, usize, usize),
    Subtopic(usize, usize, usize),
    Reference,
}
#[derive(Default)]
struct TopicBuilder {
    kind: Option<u8>,
    value: Option<String>,
    subtopics: Vec<String>,
    references: Vec<VolatileReference>,
}
struct MainBuilder {
    first: String,
    topics: Vec<TopicBuilder>,
}
struct TypeBuilder {
    dependency_type: VolatileDependencyType,
    mains: Vec<MainBuilder>,
}

fn parse_processed(xml: &[u8]) -> Result<VolatileDependencies> {
    let mut reader = NsReader::from_reader(xml);
    let mut stack = Vec::new();
    let mut types: Vec<TypeBuilder> = Vec::new();
    let mut extension = None;
    let mut root_prefixes = Vec::new();
    let mut root_closed = false;
    let mut capture: Option<(usize, Writer<Vec<u8>>)> = None;
    loop {
        let decoder = reader.decoder();
        let event = reader.read_event()?.into_owned();
        if let Some((depth, writer)) = capture.as_mut() {
            match &event {
                Event::Start(_) => {
                    if *depth >= 256 {
                        return Err(invalid("extension-list depth limit exceeded"));
                    }
                    *depth += 1;
                },
                Event::End(_) => *depth -= 1,
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("DTD and processing instructions are rejected"));
                },
                _ => {},
            }
            if *depth == 0 {
                writer
                    .write_event(Event::End(BytesEnd::new("extLst")))
                    .map_err(xml_error)?;
            } else {
                writer.write_event(event.clone()).map_err(xml_error)?;
            }
            if *depth == 0 {
                extension = Some(capture.take().unwrap().1.into_inner());
            }
            continue;
        }
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(e) if stack.is_empty() => {
                if root_closed || !name(&namespace, &e, b"volTypes") {
                    return Err(invalid("expected one SpreadsheetML volTypes root"));
                }
                root_prefixes = namespace_attributes(&e, decoder)?;
                no_attributes(&e)?;
                stack.push(Context::Root);
            },
            Event::Start(e)
                if matches!(stack.last(), Some(Context::Root))
                    && name(&namespace, &e, b"extLst") =>
            {
                if extension.is_some() || types.is_empty() {
                    return Err(invalid("invalid extLst order or duplicate"));
                }
                no_attributes(&e)?;
                let mut bindings = root_prefixes.clone();
                for (key, value) in namespace_attributes(&e, decoder)? {
                    if let Some(binding) = bindings.iter_mut().find(|binding| binding.0 == key) {
                        binding.1 = value;
                    } else {
                        bindings.push((key, value));
                    }
                }
                bindings.sort_by(|a, b| a.0.cmp(&b.0));
                let mut wrapper = BytesStart::new("extLst");
                for (key, value) in &bindings {
                    wrapper.push_attribute((key.as_str(), value.as_str()));
                }
                let mut writer = Writer::new(Vec::new());
                writer
                    .write_event(Event::Start(wrapper))
                    .map_err(xml_error)?;
                capture = Some((1, writer));
            },
            Event::Start(e) => start(
                &mut stack,
                &mut types,
                extension.is_some(),
                &namespace,
                e,
                decoder,
            )?,
            Event::Empty(e) => empty(&stack, &mut types, &mut extension, &namespace, e, decoder)?,
            Event::Text(t) => push_text(
                &mut types,
                stack.last().copied(),
                &t.decode().map_err(xml_error)?,
            )?,
            Event::CData(t) => push_text(
                &mut types,
                stack.last().copied(),
                &t.decode().map_err(xml_error)?,
            )?,
            Event::GeneralRef(r) => push_text(
                &mut types,
                stack.last().copied(),
                &litchi_ooxml_common::xml::decode_xml_reference(&r)?,
            )?,
            Event::End(_) => {
                let ended = stack
                    .pop()
                    .ok_or_else(|| invalid("closing element outside volTypes"))?;
                if matches!(ended, Context::Root) {
                    root_closed = true;
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) => {},
        }
    }
    if !root_closed || !stack.is_empty() {
        return Err(invalid("unterminated volatile-dependencies XML"));
    }
    let mut result = VolatileDependencies {
        types: Vec::with_capacity(types.len()),
        extension_list_xml: extension,
    };
    for ty in types {
        let mut output = VolatileType {
            dependency_type: ty.dependency_type,
            mains: Vec::with_capacity(ty.mains.len()),
        };
        for main in ty.mains {
            let mut converted = VolatileMain {
                first: main.first,
                topics: Vec::with_capacity(main.topics.len()),
            };
            for topic in main.topics {
                let raw = topic
                    .value
                    .ok_or_else(|| invalid("tp requires one v child"))?;
                let value = parse_value(topic.kind, raw)?;
                converted.topics.push(VolatileTopic {
                    value,
                    subtopics: topic.subtopics,
                    references: topic.references,
                });
            }
            output.mains.push(converted);
        }
        result.types.push(output);
    }
    validate_document(&result)?;
    Ok(result)
}

fn start(
    stack: &mut Vec<Context>,
    types: &mut Vec<TypeBuilder>,
    extension_seen: bool,
    ns: &ResolveResult,
    e: BytesStart<'static>,
    decoder: Decoder,
) -> Result<()> {
    match stack
        .last()
        .copied()
        .ok_or_else(|| invalid("element outside volTypes"))?
    {
        Context::Root if name(ns, &e, b"volType") => {
            if extension_seen || types.len() >= MAX_TYPES {
                return Err(invalid("invalid volType order or limit"));
            }
            let value = required_attr(&e, decoder, b"type")?;
            let dependency_type = match value.as_str() {
                "realTimeData" => VolatileDependencyType::RealTimeData,
                "olapFunctions" => VolatileDependencyType::OlapFunctions,
                _ => return Err(invalid("invalid volatile dependency type")),
            };
            only_attrs(&e, &[b"type"])?;
            types.push(TypeBuilder {
                dependency_type,
                mains: Vec::new(),
            });
            stack.push(Context::Type(types.len() - 1));
        },
        Context::Type(t) if name(ns, &e, b"main") => {
            if total_mains(types) >= MAX_MAINS {
                return Err(invalid("main-topic limit exceeded"));
            }
            let first = required_attr(&e, decoder, b"first")?;
            bounded(&first)?;
            only_attrs(&e, &[b"first"])?;
            types[t].mains.push(MainBuilder {
                first,
                topics: Vec::new(),
            });
            stack.push(Context::Main(t, types[t].mains.len() - 1));
        },
        Context::Main(t, m) if name(ns, &e, b"tp") => {
            if total_topics(types) >= MAX_TOPICS {
                return Err(invalid("topic limit exceeded"));
            }
            let kind = optional_attr(&e, decoder, b"t")?
                .map(|v| match v.as_str() {
                    "b" => Ok(b'b'),
                    "n" => Ok(b'n'),
                    "e" => Ok(b'e'),
                    "s" => Ok(b's'),
                    _ => Err(invalid("invalid volatile value type")),
                })
                .transpose()?;
            only_attrs(&e, &[b"t"])?;
            types[t].mains[m].topics.push(TopicBuilder {
                kind,
                ..Default::default()
            });
            stack.push(Context::Topic(t, m, types[t].mains[m].topics.len() - 1));
        },
        Context::Topic(t, m, p) if name(ns, &e, b"v") => {
            let topic = &mut types[t].mains[m].topics[p];
            if topic.value.is_some() || !topic.subtopics.is_empty() || !topic.references.is_empty()
            {
                return Err(invalid("v must be the first and only value child"));
            }
            no_attributes(&e)?;
            topic.value = Some(String::new());
            stack.push(Context::Value(t, m, p));
        },
        Context::Topic(t, m, p) if name(ns, &e, b"stp") => {
            let limit_reached = total_subtopics(types) >= MAX_SUBTOPICS;
            let topic = &mut types[t].mains[m].topics[p];
            if topic.value.is_none() || !topic.references.is_empty() || limit_reached {
                return Err(invalid("invalid subtopic order or limit"));
            }
            no_attributes(&e)?;
            topic.subtopics.push(String::new());
            stack.push(Context::Subtopic(t, m, p));
        },
        Context::Topic(t, m, p) if name(ns, &e, b"tr") => {
            add_reference(types, t, m, p, &e, decoder)?;
            stack.push(Context::Reference);
        },
        _ => return Err(invalid("unexpected volatile-dependencies element")),
    }
    Ok(())
}

fn empty(
    stack: &[Context],
    types: &mut [TypeBuilder],
    extension: &mut Option<Vec<u8>>,
    ns: &ResolveResult,
    e: BytesStart<'static>,
    decoder: Decoder,
) -> Result<()> {
    match stack.last().copied() {
        Some(Context::Topic(t, m, p)) if name(ns, &e, b"v") => {
            let topic = &mut types[t].mains[m].topics[p];
            if topic.value.is_some() || !topic.subtopics.is_empty() || !topic.references.is_empty()
            {
                return Err(invalid("invalid v order"));
            }
            no_attributes(&e)?;
            topic.value = Some(String::new());
        },
        Some(Context::Topic(t, m, p)) if name(ns, &e, b"stp") => {
            if types[t].mains[m].topics[p].value.is_none()
                || !types[t].mains[m].topics[p].references.is_empty()
                || total_subtopics(types) >= MAX_SUBTOPICS
            {
                return Err(invalid("invalid subtopic order or limit"));
            }
            no_attributes(&e)?;
            types[t].mains[m].topics[p].subtopics.push(String::new());
        },
        Some(Context::Topic(t, m, p)) if name(ns, &e, b"tr") => {
            add_reference(types, t, m, p, &e, decoder)?
        },
        Some(Context::Root) if name(ns, &e, b"extLst") => {
            if extension.is_some() || types.is_empty() {
                return Err(invalid("invalid extLst order or duplicate"));
            }
            no_attributes(&e)?;
            *extension = Some(b"<extLst/>".to_vec());
        },
        _ => return Err(invalid("unexpected empty volatile-dependencies element")),
    }
    Ok(())
}

fn add_reference(
    types: &mut [TypeBuilder],
    t: usize,
    m: usize,
    p: usize,
    e: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<()> {
    if total_references(types) >= MAX_REFERENCES {
        return Err(invalid("reference limit exceeded"));
    }
    let r = required_attr(e, decoder, b"r")?;
    bounded(&r)?;
    if !valid_cell_reference(&r) {
        return Err(invalid("invalid volatile cell reference"));
    }
    let s = required_attr(e, decoder, b"s")?
        .parse::<u32>()
        .map_err(|_| invalid("invalid volatile sheet id"))?;
    only_attrs(e, &[b"r", b"s"])?;
    let topic = &mut types[t].mains[m].topics[p];
    if topic.value.is_none() {
        return Err(invalid("tr must follow v"));
    }
    topic.references.push(VolatileReference {
        cell_reference: r,
        sheet_id: s,
    });
    Ok(())
}
fn push_text(types: &mut [TypeBuilder], ctx: Option<Context>, text: &str) -> Result<()> {
    match ctx {
        Some(Context::Value(t, m, p)) => {
            append(types[t].mains[m].topics[p].value.as_mut().unwrap(), text)
        },
        Some(Context::Subtopic(t, m, p)) => append(
            types[t].mains[m].topics[p].subtopics.last_mut().unwrap(),
            text,
        ),
        _ if text.trim().is_empty() => Ok(()),
        _ => Err(invalid("text outside v or stp")),
    }
}
fn append(out: &mut String, text: &str) -> Result<()> {
    if out.len().saturating_add(text.len()) > MAX_TEXT_BYTES {
        return Err(invalid("volatile text limit exceeded"));
    }
    out.push_str(text);
    Ok(())
}

fn parse_value(kind: Option<u8>, raw: String) -> Result<VolatileValue> {
    Ok(match kind {
        None => VolatileValue::Unspecified(raw),
        Some(b'b') => VolatileValue::Boolean(match raw.as_str() {
            "1" | "true" => true,
            "0" | "false" => false,
            _ => return Err(invalid("invalid volatile boolean")),
        }),
        Some(b'n') => {
            let v = raw
                .parse::<f64>()
                .map_err(|_| invalid("invalid volatile number"))?;
            if !v.is_finite() {
                return Err(invalid("non-finite volatile number"));
            }
            VolatileValue::Number(v)
        },
        Some(b'e') => VolatileValue::Error(raw),
        Some(b's') => VolatileValue::String(raw),
        _ => unreachable!(),
    })
}
fn validate_document(d: &VolatileDependencies) -> Result<()> {
    if d.types.is_empty() || d.types.len() > MAX_TYPES {
        return Err(invalid("volTypes requires 1..64 volType children"));
    }
    let mut mains = 0;
    let mut topics = 0;
    let mut subs = 0;
    let mut refs = 0;
    for ty in &d.types {
        if ty.mains.is_empty() {
            return Err(invalid("volType requires a main child"));
        }
        for main in &ty.mains {
            bounded(&main.first)?;
            if main.topics.is_empty() {
                return Err(invalid("main requires a tp child"));
            }
            for topic in &main.topics {
                match &topic.value {
                    VolatileValue::Unspecified(v)
                    | VolatileValue::Error(v)
                    | VolatileValue::String(v) => bounded(v)?,
                    VolatileValue::Number(v) if !v.is_finite() => {
                        return Err(invalid("non-finite volatile number"));
                    },
                    _ => {},
                }
                if topic.references.is_empty() {
                    return Err(invalid("tp requires a tr child"));
                }
                for v in &topic.subtopics {
                    bounded(v)?;
                }
                for r in &topic.references {
                    bounded(&r.cell_reference)?;
                    if !valid_cell_reference(&r.cell_reference) {
                        return Err(invalid("invalid volatile cell reference"));
                    }
                }
                subs += topic.subtopics.len();
                refs += topic.references.len();
            }
            topics += main.topics.len();
        }
        mains += ty.mains.len();
    }
    if mains > MAX_MAINS || topics > MAX_TOPICS || subs > MAX_SUBTOPICS || refs > MAX_REFERENCES {
        return Err(invalid("volatile-dependencies resource limit exceeded"));
    }
    if let Some(ext) = &d.extension_list_xml
        && ext.len() > MAX_PART_BYTES
    {
        return Err(invalid("extension list exceeds resource limit"));
    }
    Ok(())
}

fn name(ns: &ResolveResult, e: &BytesStart<'_>, local: &[u8]) -> bool {
    let namespace_matches = match ns {
        ResolveResult::Bound(Namespace(v)) => {
            let bytes: &[u8] = v;
            bytes == NS || bytes == STRICT_NS
        },
        _ => false,
    };
    namespace_matches && e.local_name().as_ref() == local
}
fn required_attr(e: &BytesStart<'_>, d: Decoder, n: &[u8]) -> Result<String> {
    optional_attr(e, d, n)?.ok_or_else(|| {
        invalid(format!(
            "missing required attribute '{}'",
            String::from_utf8_lossy(n)
        ))
    })
}
fn optional_attr(e: &BytesStart<'_>, d: Decoder, n: &[u8]) -> Result<Option<String>> {
    let mut value = None;
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        if a.key.as_ref() == n {
            if value.is_some() {
                return Err(invalid("duplicate attribute"));
            }
            value = Some(
                a.decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
                    .map_err(xml_error)?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}
fn only_attrs(e: &BytesStart<'_>, allowed: &[&[u8]]) -> Result<()> {
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        let k = a.key.as_ref();
        if k == b"xmlns" || k.starts_with(b"xmlns:") {
            continue;
        }
        if k.contains(&b':') || !allowed.contains(&k) {
            return Err(invalid(format!(
                "unexpected attribute '{}'",
                String::from_utf8_lossy(k)
            )));
        }
    }
    Ok(())
}
fn namespace_attributes(e: &BytesStart<'_>, d: Decoder) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        let key = std::str::from_utf8(a.key.as_ref()).map_err(xml_error)?;
        if key.starts_with("xmlns:") && key != "xmlns:xml" {
            values.push((
                key.to_string(),
                a.decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
                    .map_err(xml_error)?
                    .into_owned(),
            ));
        }
    }
    Ok(values)
}
fn no_attributes(e: &BytesStart<'_>) -> Result<()> {
    only_attrs(e, &[])
}
fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_TEXT_BYTES {
        Err(invalid("volatile string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}
fn total_mains(v: &[TypeBuilder]) -> usize {
    v.iter().map(|x| x.mains.len()).sum()
}
fn total_topics(v: &[TypeBuilder]) -> usize {
    v.iter()
        .flat_map(|x| &x.mains)
        .map(|x| x.topics.len())
        .sum()
}
fn total_subtopics(v: &[TypeBuilder]) -> usize {
    v.iter()
        .flat_map(|x| &x.mains)
        .flat_map(|x| &x.topics)
        .map(|x| x.subtopics.len())
        .sum()
}
fn total_references(v: &[TypeBuilder]) -> usize {
    v.iter()
        .flat_map(|x| &x.mains)
        .flat_map(|x| &x.topics)
        .map(|x| x.references.len())
        .sum()
}
fn valid_cell_reference(v: &str) -> bool {
    let b = v.as_bytes();
    let mut i = 0;
    if b.get(i) == Some(&b'$') {
        i += 1;
    }
    let start = i;
    while b.get(i).is_some_and(u8::is_ascii_alphabetic) {
        i += 1;
    }
    if i == start || i - start > 3 {
        return false;
    }
    if b.get(i) == Some(&b'$') {
        i += 1;
    }
    let row = i;
    while b.get(i).is_some_and(u8::is_ascii_digit) {
        i += 1;
    }
    i == b.len() && i > row && b[row] != b'0'
}
fn escape_attr(o: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => o.push_str("&amp;"),
            '<' => o.push_str("&lt;"),
            '"' => o.push_str("&quot;"),
            '\r' => o.push_str("&#xD;"),
            '\n' => o.push_str("&#xA;"),
            '\t' => o.push_str("&#x9;"),
            _ => o.push(c),
        }
    }
}
fn escape_text(o: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => o.push_str("&amp;"),
            '<' => o.push_str("&lt;"),
            '>' => o.push_str("&gt;"),
            _ => o.push(c),
        }
    }
}
fn invalid(message: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, message.into()).into()
}
fn xml_error(e: impl std::fmt::Display) -> Box<dyn std::error::Error + Send + Sync> {
    invalid(e.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, Part};

    fn sample(ns: &str) -> Vec<u8> {
        format!(r#"<volTypes xmlns="{ns}"><volType type="realTimeData"><main first="server.id"><tp t="s"><v>ready</v><stp>ticker</stp><tr r="$A$1" s="0"/></tp></main></volType><volType type="olapFunctions"><main first="cube"><tp t="n"><v>42.5</v><tr r="B2" s="1"/></tp></main></volType></volTypes>"#).into_bytes()
    }

    fn value() -> VolatileDependencies {
        VolatileDependencies::parse(&sample(std::str::from_utf8(NS).unwrap())).unwrap()
    }

    fn workbook_package() -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let workbook = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.into(),
            format!(
                r#"<workbook xmlns="{}"><sheets/></workbook>"#,
                std::str::from_utf8(NS).unwrap()
            )
            .into_bytes(),
        );
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook));
        package
    }

    fn synthetic_package(
        relationship_type: &str,
        external: bool,
        content_type: &str,
        outbound: bool,
    ) -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let mut workbook = BlobPart::new(
            workbook_uri.clone(),
            ct::SML_SHEET_MAIN.into(),
            format!(
                r#"<workbook xmlns="{}"><sheets/></workbook>"#,
                std::str::from_utf8(NS).unwrap()
            )
            .into_bytes(),
        );
        if external {
            workbook.relate_to_ext(
                "https://example.invalid/volatileDependencies.xml",
                relationship_type,
            );
        } else {
            workbook.relate_to("volatileDependencies.xml", relationship_type);
        }
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook));
        if !external {
            let mut dependencies = BlobPart::new(
                PackURI::new("/xl/volatileDependencies.xml").unwrap(),
                content_type.into(),
                sample(std::str::from_utf8(NS).unwrap()),
            );
            if outbound {
                dependencies.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
            }
            package.add_part(Box::new(dependencies));
        }
        package
    }
    #[test]
    fn parses_and_writes_transitional_and_strict() {
        for ns in [
            std::str::from_utf8(NS).unwrap(),
            std::str::from_utf8(STRICT_NS).unwrap(),
        ] {
            let parsed = VolatileDependencies::parse(&sample(ns)).unwrap();
            assert_eq!(parsed.types.len(), 2);
            let strict = parsed.to_xml(true).unwrap();
            assert_eq!(VolatileDependencies::parse(&strict).unwrap(), parsed);
        }
    }
    #[test]
    fn applies_mce_fallback() {
        let xml = format!(
            r#"<volTypes xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><u:item/></mc:Choice><mc:Fallback><volType type="realTimeData"><main first="srv"><tp t="b"><v>true</v><tr r="A1" s="0"/></tp></main></volType></mc:Fallback></mc:AlternateContent></volTypes>"#,
            std::str::from_utf8(NS).unwrap()
        );
        assert_eq!(
            VolatileDependencies::parse(xml.as_bytes())
                .unwrap()
                .types
                .len(),
            1
        );
    }
    #[test]
    fn preserves_non_empty_extension_list_and_inherited_namespaces() {
        let xml = format!(
            r#"<volTypes xmlns="{}" xmlns:x14="urn:shadowed" xmlns:x15="urn:inherited"><volType type="realTimeData"><main first="srv"><tp><v>ok</v><tr r="A1" s="0"/></tp></main></volType><extLst xmlns:x14="urn:local"><ext uri="urn:test"><x14:payload value="kept"/><x15:payload/></ext></extLst></volTypes>"#,
            std::str::from_utf8(NS).unwrap()
        );
        let parsed = VolatileDependencies::parse(xml.as_bytes()).unwrap();
        let written = parsed.to_xml(true).unwrap();
        let text = std::str::from_utf8(&written).unwrap();
        assert!(text.contains("xmlns:x14=\"urn:local\""));
        assert!(text.contains("xmlns:x15=\"urn:inherited\""));
        assert!(!text.contains("urn:shadowed"));
        assert!(text.contains("x14:payload"));
        assert_eq!(VolatileDependencies::parse(&written).unwrap(), parsed);
    }
    #[test]
    fn rejects_malformed_and_unsafe_input() {
        for xml in [
            format!(
                r#"<volTypes xmlns="{}"/>"#,
                std::str::from_utf8(NS).unwrap()
            ),
            format!(
                r#"<volTypes xmlns="{}"><volType type="bad"><main first="x"><tp><v/><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,
                std::str::from_utf8(NS).unwrap()
            ),
            format!(
                r#"<volTypes xmlns="{}"><volType type="realTimeData"><main first="x"><tp t="b"><v>maybe</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,
                std::str::from_utf8(NS).unwrap()
            ),
            format!(
                r#"<!DOCTYPE x [<!ENTITY e "boom">]><volTypes xmlns="{}"><volType type="realTimeData"><main first="x"><tp><v>&e;</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#,
                std::str::from_utf8(NS).unwrap()
            ),
        ] {
            assert!(
                VolatileDependencies::parse(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
    }
    #[test]
    fn resolves_package_relationship_and_rejects_outbound_relationships() {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
        let mut workbook = BlobPart::new(
            workbook_uri,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            Vec::new(),
        );
        workbook.relate_to("deps.xml", REL);
        package.relate_to(
            "custom/book.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        );
        package.add_part(Box::new(workbook));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/deps.xml").unwrap(),
            CONTENT_TYPE.into(),
            sample(std::str::from_utf8(NS).unwrap()),
        )));
        assert_eq!(load_from_package(&package).unwrap().unwrap().types.len(), 2);
        package
            .get_part_mut(&PackURI::new("/custom/deps.xml").unwrap())
            .unwrap()
            .relate_to("other.xml", "urn:forbidden");
        assert!(load_from_package(&package).is_err());
    }

    #[test]
    fn stores_rewrites_and_removes_inert_volatile_dependencies_parts() {
        let mut package = workbook_package();
        let value = value();

        store_in_package(
            &mut package,
            &value,
            VolatileDependenciesConformance::Transitional,
        )
        .unwrap();
        assert_eq!(load_from_package(&package).unwrap(), Some(value.clone()));
        assert_eq!(
            load_from_package_with_conformance(&package).unwrap(),
            Some((value.clone(), VolatileDependenciesConformance::Transitional))
        );

        let workbook = package.main_document_part().unwrap();
        let relationship = workbook
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == REL)
            .unwrap();
        let relationship_id = relationship.r_id().to_string();
        let part_name = relationship.target_partname().unwrap();
        assert_eq!(
            part_name,
            PackURI::new("/xl/volatileDependencies.xml").unwrap()
        );
        assert!(
            std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
                .unwrap()
                .contains(std::str::from_utf8(NS).unwrap())
        );

        let mut replacement = value.clone();
        replacement.types[0].mains[0].first = "replacement.server".into();
        store_in_package(
            &mut package,
            &replacement,
            VolatileDependenciesConformance::Strict,
        )
        .unwrap();
        let workbook = package.main_document_part().unwrap();
        let relationship = workbook
            .rels()
            .iter()
            .find(|relationship| relationship.r_id() == relationship_id)
            .unwrap();
        assert_eq!(relationship.reltype(), STRICT_REL);
        assert_eq!(relationship.target_partname().unwrap(), part_name);
        assert!(
            std::str::from_utf8(package.get_part(&part_name).unwrap().blob())
                .unwrap()
                .contains(std::str::from_utf8(STRICT_NS).unwrap())
        );
        assert_eq!(
            load_from_package_with_conformance(&package).unwrap(),
            Some((replacement, VolatileDependenciesConformance::Strict))
        );

        assert!(remove_from_package(&mut package).unwrap());
        assert!(package.get_part(&part_name).is_err());
        assert_eq!(load_from_package(&package).unwrap(), None);
        assert!(!remove_from_package(&mut package).unwrap());
    }

    #[test]
    fn removal_retains_volatile_dependencies_part_referenced_elsewhere() {
        let mut package = workbook_package();
        let value = value();
        store_in_package(
            &mut package,
            &value,
            VolatileDependenciesConformance::Transitional,
        )
        .unwrap();

        let part_name = PackURI::new("/xl/volatileDependencies.xml").unwrap();
        let mut referring_part = BlobPart::new(
            PackURI::new("/xl/retained-reference.xml").unwrap(),
            ct::XML.into(),
            b"<reference/>".to_vec(),
        );
        referring_part.relate_to(
            "volatileDependencies.xml",
            "urn:litchi:test:volatile-dependencies-reference",
        );
        package.add_part(Box::new(referring_part));

        assert!(remove_from_package(&mut package).unwrap());
        assert!(package.get_part(&part_name).is_ok());
        assert!(load_from_package(&package).is_err());
        assert!(
            store_in_package(
                &mut package,
                &value,
                VolatileDependenciesConformance::Transitional,
            )
            .is_err()
        );
    }

    #[test]
    fn workbook_volatile_dependencies_mutators_and_materialization_preserve_metadata() {
        let mut workbook = crate::xlsx::Workbook::create().unwrap();
        let value = value();
        workbook
            .set_volatile_dependencies(&value, VolatileDependenciesConformance::Strict)
            .unwrap();
        assert_eq!(
            workbook.volatile_dependencies().unwrap(),
            Some((value.clone(), VolatileDependenciesConformance::Strict))
        );

        workbook
            .worksheet_mut(0)
            .unwrap()
            .set_cell_value(1, 1, "materialized");
        let directory = tempfile::tempdir().unwrap();
        let path = directory
            .path()
            .join("materialized-volatile-dependencies.xlsx");
        workbook.save(&path).unwrap();
        let reopened = crate::xlsx::Workbook::open(&path).unwrap();
        assert_eq!(
            reopened.volatile_dependencies().unwrap(),
            Some((value.clone(), VolatileDependenciesConformance::Strict))
        );

        let mut reopened = reopened;
        assert!(reopened.remove_volatile_dependencies().unwrap());
        assert_eq!(reopened.volatile_dependencies().unwrap(), None);
        assert!(!reopened.remove_volatile_dependencies().unwrap());
    }

    #[test]
    fn package_volatile_dependencies_mutators_reject_invalid_existing_graphs() {
        let value = value();
        let mut wrong_content_type = synthetic_package(REL, false, ct::SML_STYLES, false);
        let part_name = PackURI::new("/xl/volatileDependencies.xml").unwrap();
        let original = wrong_content_type
            .get_part(&part_name)
            .unwrap()
            .blob()
            .to_vec();
        assert!(
            store_in_package(
                &mut wrong_content_type,
                &value,
                VolatileDependenciesConformance::Transitional,
            )
            .is_err()
        );
        assert_eq!(
            wrong_content_type.get_part(&part_name).unwrap().blob(),
            original
        );
        assert!(remove_from_package(&mut wrong_content_type).is_err());

        let mut duplicate = synthetic_package(REL, false, CONTENT_TYPE, false);
        duplicate
            .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                REL.into(),
                "volatileDependencies.xml".into(),
                "rIdDuplicateVolatileDependencies".into(),
                false,
            );
        assert!(
            store_in_package(
                &mut duplicate,
                &value,
                VolatileDependenciesConformance::Transitional,
            )
            .is_err()
        );
        assert!(remove_from_package(&mut duplicate).is_err());

        let mut duplicate_part = synthetic_package(REL, false, CONTENT_TYPE, false);
        duplicate_part.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/volatileDependenciesExtra.xml").unwrap(),
            CONTENT_TYPE.into(),
            sample(std::str::from_utf8(NS).unwrap()),
        )));
        assert!(load_from_package(&duplicate_part).is_err());
        assert!(
            store_in_package(
                &mut duplicate_part,
                &value,
                VolatileDependenciesConformance::Transitional,
            )
            .is_err()
        );
        assert!(remove_from_package(&mut duplicate_part).is_err());

        let mut external = synthetic_package(REL, true, CONTENT_TYPE, false);
        assert!(
            store_in_package(
                &mut external,
                &value,
                VolatileDependenciesConformance::Transitional,
            )
            .is_err()
        );
        assert!(remove_from_package(&mut external).is_err());

        let mut outbound = synthetic_package(REL, false, CONTENT_TYPE, true);
        assert!(
            store_in_package(
                &mut outbound,
                &value,
                VolatileDependenciesConformance::Transitional,
            )
            .is_err()
        );
        assert!(remove_from_package(&mut outbound).is_err());

        let mut root_relationship = workbook_package();
        root_relationship.relate_to("xl/volatileDependencies.xml", REL);
        assert!(
            store_in_package(
                &mut root_relationship,
                &value,
                VolatileDependenciesConformance::Transitional,
            )
            .is_err()
        );
    }
}
