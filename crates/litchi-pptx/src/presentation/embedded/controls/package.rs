use super::codec::{parse_descriptor, scan};
use super::model::{Binary, Control, Descriptor};
use super::{
    BINARY_CONTENT_TYPE, BINARY_RELATIONSHIP, CONTROL_RELATIONSHIP, DESCRIPTOR_CONTENT_TYPE,
    MAX_BINARY_BYTES, MAX_CONTROLS, MAX_TOTAL_BINARY_BYTES, MAX_TOTAL_SLIDE_XML_BYTES,
    STRICT_CONTROL_RELATIONSHIP,
};
use crate::presentation::embedded::{invalid, limit};
use crate::{Error, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, Part};

/// Shared bounds for control discovery across a presentation.
#[derive(Debug, Default)]
pub struct Limits {
    total_slide_xml_bytes: usize,
    control_count: usize,
    total_binary_bytes: usize,
}

impl Limits {
    fn add_slide(&mut self, bytes: usize) -> Result<()> {
        self.total_slide_xml_bytes = self
            .total_slide_xml_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("control slide XML bytes", MAX_TOTAL_SLIDE_XML_BYTES))?;
        if self.total_slide_xml_bytes > MAX_TOTAL_SLIDE_XML_BYTES {
            return Err(limit("control slide XML bytes", MAX_TOTAL_SLIDE_XML_BYTES));
        }
        Ok(())
    }

    fn add_control(&mut self) -> Result<()> {
        self.control_count = self
            .control_count
            .checked_add(1)
            .ok_or_else(|| limit("control count", MAX_CONTROLS))?;
        if self.control_count > MAX_CONTROLS {
            return Err(limit("control count", MAX_CONTROLS));
        }
        Ok(())
    }

    fn add_binary(&mut self, bytes: usize) -> Result<()> {
        if bytes > MAX_BINARY_BYTES {
            return Err(limit("control binary bytes", MAX_BINARY_BYTES));
        }
        self.total_binary_bytes = self
            .total_binary_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("control binary bytes", MAX_TOTAL_BINARY_BYTES))?;
        if self.total_binary_bytes > MAX_TOTAL_BINARY_BYTES {
            return Err(limit("control binary bytes", MAX_TOTAL_BINARY_BYTES));
        }
        Ok(())
    }
}

/// Load one slide's bounded control inventory and validate its OPC graph.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_slide(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut Limits,
) -> Result<Vec<Control>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid("control discovery requires a PresentationML slide"));
    }
    limits.add_slide(slide.blob().len())?;
    let mut count = 0;
    let parsed = scan(slide.blob(), &mut count)?;
    let mut values = Vec::with_capacity(parsed.len());
    for (index, parsed) in parsed.into_iter().enumerate() {
        limits.add_control()?;
        let descriptor = parsed
            .relationship_id
            .as_deref()
            .map(|id| resolve_descriptor(package, slide_index, slide, id, limits))
            .transpose()?;
        values.push(Control {
            slide_index,
            index,
            shape_id: parsed.shape_id,
            name: parsed.name,
            show_as_icon: parsed.show_as_icon,
            image_width: parsed.image_width,
            image_height: parsed.image_height,
            relationship_id: parsed.relationship_id,
            descriptor,
        });
    }
    Ok(values)
}

fn resolve_descriptor(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    relationship_id: &str,
    limits: &mut Limits,
) -> Result<Descriptor> {
    let relationship = slide.rels().get(relationship_id).ok_or_else(|| {
        Error::Relationship(format!(
            "slide {slide_index} control relationship '{relationship_id}' is missing"
        ))
    })?;
    if relationship.is_external()
        || !matches!(
            relationship.reltype(),
            CONTROL_RELATIONSHIP | STRICT_CONTROL_RELATIONSHIP
        )
    {
        return Err(Error::Relationship(format!(
            "slide {slide_index} control relationship '{relationship_id}' is not an internal control relationship"
        )));
    }
    let part_name = relationship.target_partname()?;
    let part = package.get_part(&part_name)?;
    if part.content_type() != DESCRIPTOR_CONTENT_TYPE {
        return Err(Error::ContentType {
            expected: DESCRIPTOR_CONTENT_TYPE.to_string(),
            actual: part.content_type().to_string(),
        });
    }
    let (class_id, license, persistence, binary_id) = parse_descriptor(part.blob())?;
    let binary = binary_id
        .as_deref()
        .map(|id| resolve_binary(package, slide_index, part, id, limits))
        .transpose()?;
    Ok(Descriptor {
        part_name,
        class_id,
        license,
        persistence,
        binary,
    })
}

fn resolve_binary(
    package: &OpcPackage,
    slide_index: usize,
    descriptor: &dyn Part,
    relationship_id: &str,
    limits: &mut Limits,
) -> Result<Binary> {
    let relationship = descriptor.rels().get(relationship_id).ok_or_else(|| {
        Error::Relationship(format!(
            "slide {slide_index} control descriptor binary relationship '{relationship_id}' is missing"
        ))
    })?;
    if relationship.is_external() || relationship.reltype() != BINARY_RELATIONSHIP {
        return Err(Error::Relationship(format!(
            "control binary relationship '{relationship_id}' has an unsupported type"
        )));
    }
    let part_name = relationship.target_partname()?;
    let part = package.get_part(&part_name)?;
    if part.content_type() != BINARY_CONTENT_TYPE {
        return Err(Error::ContentType {
            expected: BINARY_CONTENT_TYPE.to_string(),
            actual: part.content_type().to_string(),
        });
    }
    limits.add_binary(part.blob().len())?;
    Ok(Binary {
        relationship_id: relationship_id.to_string(),
        part_name,
        byte_length: part.blob().len(),
    })
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::part::BlobPart;

    #[test]
    fn scans_contextual_control_metadata() {
        let xml = br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:controls><p:control name="CheckBox" showAsIcon="1" imgW="12"/></p:controls></p:cSld></p:sld>"#;
        let mut package = OpcPackage::new();
        let uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            ct::PML_SLIDE.to_string(),
            xml.to_vec(),
        )));
        let controls = load_slide(
            &package,
            0,
            package.get_part(&uri).unwrap(),
            &mut Limits::default(),
        )
        .unwrap();
        assert_eq!(controls[0].name(), Some("CheckBox"));
        assert_eq!(controls[0].image_width(), Some(12));
    }
}
