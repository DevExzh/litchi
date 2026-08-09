use super::CONTENT_TYPE;
use super::codec::{
    MAX_ANNOTATIONS, MAX_INK_BYTES, MAX_TOTAL_INK_BYTES, Summary, inspect, root_namespaces,
    scan_slide,
};
use super::model::{Annotation, StoredAnnotation};
use crate::presentation::embedded::{
    MAX_XML_DEPTH, increment_nodes, invalid, is_presentationml_name, limit,
};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

/// Shared bounds for `InkML` discovery across slides.
#[derive(Debug, Default)]
pub struct Limits {
    annotation_count: usize,
    total_bytes: usize,
}

impl Limits {
    fn add(&mut self, bytes: usize) -> Result<()> {
        self.annotation_count = self
            .annotation_count
            .checked_add(1)
            .ok_or_else(|| limit("InkML annotation count", MAX_ANNOTATIONS))?;
        if self.annotation_count > MAX_ANNOTATIONS {
            return Err(limit("InkML annotation count", MAX_ANNOTATIONS));
        }
        self.total_bytes = self
            .total_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("total InkML bytes", MAX_TOTAL_INK_BYTES))?;
        if self.total_bytes > MAX_TOTAL_INK_BYTES {
            return Err(limit("total InkML bytes", MAX_TOTAL_INK_BYTES));
        }
        Ok(())
    }
}

/// Load one slide's `InkML` content-part graph without copying payload bytes.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_slide(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut Limits,
) -> Result<Vec<Annotation>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid("InkML discovery requires a PresentationML slide"));
    }
    let ids = scan_slide(slide.blob())?;
    let mut annotations = Vec::new();
    for (index, relationship_id) in ids.into_iter().enumerate() {
        let relationship = slide.rels().get(&relationship_id).ok_or_else(|| {
            Error::Relationship(format!(
                "slide {slide_index} InkML relationship '{relationship_id}' is missing"
            ))
        })?;
        if relationship.is_external() || relationship.reltype() != rt::CUSTOM_XML {
            return Err(Error::Relationship(format!(
                "slide {slide_index} InkML relationship '{relationship_id}' must be internal customXml"
            )));
        }
        let part_name = relationship.target_partname()?;
        let part = package.get_part(&part_name)?;
        if part.content_type() != CONTENT_TYPE {
            continue;
        }
        limits.add(part.blob().len())?;
        let Summary { traces, groups } = inspect(part.blob())?;
        annotations.push(Annotation {
            slide_index,
            index,
            relationship_id,
            part_name,
            trace_count: traces,
            trace_group_count: groups,
        });
    }
    Ok(annotations)
}

/// Store one validated `InkML` payload on a slide as a custom XML content part.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn store_slide(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    inkml: &[u8],
) -> Result<StoredAnnotation> {
    if inkml.len() > MAX_INK_BYTES {
        return Err(limit("InkML part bytes", MAX_INK_BYTES));
    }
    inspect(inkml)?;
    let slide = package.get_part(slide_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid("InkML storage requires a PresentationML slide"));
    }
    let (pml, rel) = root_namespaces(slide.blob())?;
    let part_name = allocate_part(package)?;
    let relationship_id = allocate_relationship(slide)?;
    let fragment =
        format!(r#"<p:contentPart xmlns:p="{pml}" xmlns:r="{rel}" r:id="{relationship_id}"/>"#);
    let updated = insert_into_shape_tree(slide.blob(), fragment.as_bytes())?;
    package.get_part_mut(slide_name)?.set_blob(updated);
    package.add_part(Box::new(BlobPart::new(
        part_name.clone(),
        CONTENT_TYPE.to_string(),
        inkml.to_vec(),
    )));
    package
        .get_part_mut(slide_name)?
        .rels_mut()
        .add_relationship(
            rt::CUSTOM_XML.to_string(),
            part_name.relative_ref(slide_name.base_uri()),
            relationship_id.clone(),
            false,
        );
    Ok(StoredAnnotation {
        relationship_id,
        part_name,
    })
}

fn allocate_part(package: &OpcPackage) -> Result<PackURI> {
    for index in 1..1_000_000u32 {
        let name = PackURI::new(format!("/ppt/ink/ink{index}.xml")).map_err(Error::Uri)?;
        if package.get_part(&name).is_err() {
            return Ok(name);
        }
    }
    Err(limit("InkML part namespace", 1_000_000))
}

fn allocate_relationship(slide: &dyn Part) -> Result<String> {
    for index in 1..1_000_000u32 {
        let id = format!("rId{index}");
        if slide.rels().get(&id).is_none() {
            return Ok(id);
        }
    }
    Err(limit("InkML relationship namespace", 1_000_000))
}

fn insert_into_shape_tree(xml: &[u8], fragment: &[u8]) -> Result<Vec<u8>> {
    if xml
        .len()
        .checked_add(fragment.len())
        .is_none_or(|size| size > super::codec::MAX_SLIDE_BYTES)
    {
        return Err(limit(
            "updated slide XML bytes",
            super::codec::MAX_SLIDE_BYTES,
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut tree_depth = None;
    let mut insertion = None;
    loop {
        let before = usize::try_from(reader.buffer_position())
            .map_err(|_err| invalid("slide XML offset overflow"))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("InkML slide depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("InkML slide depth", MAX_XML_DEPTH));
                }
                if depth == 1 {
                    crate::presentation::embedded::validate_root(
                        &namespace,
                        element.name(),
                        root_seen,
                    )?;
                    root_seen = true;
                }
                if is_presentationml_name(&namespace, element.name(), b"spTree")
                    && tree_depth.replace(depth).is_some()
                {
                    return Err(invalid("slide has multiple shape trees"));
                }
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                if is_presentationml_name(&namespace, element.name(), b"spTree") {
                    return Err(invalid("cannot insert InkML into an empty shape tree"));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid slide XML nesting"));
                }
                if tree_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"spTree")
                {
                    insertion = Some(before);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "slide XML rejects DTDs and processing instructions",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !root_seen {
        return Err(invalid("unterminated slide XML"));
    }
    let position = insertion.ok_or_else(|| invalid("slide has no shape tree"))?;
    let mut output = Vec::with_capacity(xml.len() + fragment.len());
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}
