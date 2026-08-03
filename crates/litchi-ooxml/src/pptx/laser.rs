//! Host integration for the canonical inert PPTX laser-trace codec.

use crate::error::{OoxmlError, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::part::Part;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::laser::{self, Conformance};

pub use litchi_pptx::laser::{LASER_TRACE_EXTENSION_URI, PptxLaserTrace, PptxLaserTracePoint};

pub(crate) type LaserLoadLimits = litchi_pptx::laser::Limits;

/// Load bounded, inert laser-pointer traces from one PresentationML slide.
pub(crate) fn load_slide_laser_traces(
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut LaserLoadLimits,
) -> Result<Vec<PptxLaserTrace>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "laser-trace discovery requires a PresentationML slide part",
        ));
    }
    Ok(laser::read_with(slide_index, slide.blob(), limits)?)
}

/// Store one laser-pointer trace onto a slide as a PowerPoint 2010
/// p14:laserTraceLst extension.
pub fn store_slide_laser_trace(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    points: &[PptxLaserTracePoint],
) -> Result<()> {
    laser::validate(points)?;
    let slide = package.get_part(slide_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "laser-trace storage requires a PresentationML slide part",
        ));
    }
    if !load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default())?.is_empty() {
        return Err(invalid(
            "slide already contains a laser-trace extension; replacement is not supported",
        ));
    }

    let conformance =
        Conformance::from_namespace(crate::pptx::slide_patch::slide_dialect(slide.blob())?);
    let fragment = laser::write(points, conformance)?;
    let updated = crate::pptx::slide_patch::insert_extension_fragment(slide.blob(), &fragment)?;

    let probe =
        litchi_opc::BlobPart::new(slide_name.clone(), ct::PML_SLIDE.into(), updated.clone());
    let traces = load_slide_laser_traces(0, &probe, &mut LaserLoadLimits::default())?;
    if traces.len() != 1 || traces[0].points().len() != points.len() {
        return Err(invalid("laser-trace storage failed read-back validation"));
    }
    package.get_part_mut(slide_name)?.set_blob(updated);
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use litchi_drawingml::coord::Coordinate;
    use litchi_pptx::time::Offset;

    use super::*;

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    fn slide_package(tail: &str) -> (litchi_opc::OpcPackage, litchi_opc::PackURI) {
        let mut package = litchi_opc::OpcPackage::new();
        let name = litchi_opc::PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>{tail}</p:sld>"#
        );
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            name.clone(),
            ct::PML_SLIDE.into(),
            xml.into_bytes(),
        )));
        (package, name)
    }

    fn sample_points() -> Vec<PptxLaserTracePoint> {
        vec![
            PptxLaserTracePoint::new(
                Offset::ZERO,
                Coordinate::emu(914_400).unwrap(),
                Coordinate::emu(457_200).unwrap(),
            ),
            PptxLaserTracePoint::new(
                Offset::ms(2500),
                Coordinate::parse("1.25cm").unwrap(),
                Coordinate::from(34),
            ),
        ]
    }

    #[test]
    fn stores_laser_trace_and_discovers_it_round_trip() {
        let (mut package, slide_name) = slide_package("");
        let points = sample_points();
        store_slide_laser_trace(&mut package, &slide_name, &points).unwrap();

        let slide = package.get_part(&slide_name).unwrap();
        let traces = load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default()).unwrap();
        assert_eq!(traces.len(), 1);
        assert_eq!(traces[0].points(), points.as_slice());

        // A second trace on the same slide is rejected (no replacement).
        assert!(store_slide_laser_trace(&mut package, &slide_name, &points).is_err());
    }

    #[test]
    fn stores_laser_trace_into_existing_and_empty_extension_lists() {
        // Existing non-empty extLst.
        let (mut package, slide_name) = slide_package(
            r#"<p:extLst><p:ext uri="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"/></p:extLst>"#,
        );
        store_slide_laser_trace(&mut package, &slide_name, &sample_points()).unwrap();
        let slide = package.get_part(&slide_name).unwrap();
        let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
        assert!(xml.contains("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"));
        assert_eq!(
            load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default())
                .unwrap()
                .len(),
            1
        );

        // Empty extLst element.
        let (mut package, slide_name) = slide_package("<p:extLst/>");
        store_slide_laser_trace(&mut package, &slide_name, &sample_points()).unwrap();
        let slide = package.get_part(&slide_name).unwrap();
        let traces = load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default()).unwrap();
        assert_eq!(traces.len(), 1);
        assert_eq!(traces[0].point_count(), 2);
    }

    #[test]
    fn stores_laser_trace_in_strict_dialect() {
        let mut package = litchi_opc::OpcPackage::new();
        let name = litchi_opc::PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let xml = r#"<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#;
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            name.clone(),
            ct::PML_SLIDE.into(),
            xml.as_bytes().to_vec(),
        )));
        store_slide_laser_trace(&mut package, &name, &sample_points()).unwrap();
        let slide = package.get_part(&name).unwrap();
        let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
        assert!(xml.contains("xmlns:p=\"http://purl.oclc.org/ooxml/presentationml/main\""));
        assert_eq!(
            load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default())
                .unwrap()
                .len(),
            1
        );
    }

    #[test]
    fn rejects_invalid_laser_storage_inputs() {
        let (mut package, slide_name) = slide_package("");
        // No points.
        assert!(store_slide_laser_trace(&mut package, &slide_name, &[]).is_err());
        // Bad time offsets cannot enter the typed point constructor.
        assert!(Offset::parse("").is_err());
        assert!(Offset::parse("a<b").is_err());
        // Non-slide part.
        let wrong = litchi_opc::PackURI::new("/ppt/presentation.xml").unwrap();
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            wrong.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            b"<p:presentation/>".to_vec(),
        )));
        assert!(store_slide_laser_trace(&mut package, &wrong, &sample_points()).is_err());
        // Rejection leaves the slide without an extension list.
        let slide = package.get_part(&slide_name).unwrap();
        assert!(
            load_slide_laser_traces(0, slide, &mut LaserLoadLimits::default())
                .unwrap()
                .is_empty()
        );
    }
}
