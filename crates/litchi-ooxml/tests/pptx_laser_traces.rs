use litchi_ooxml::pptx::{LASER_TRACE_EXTENSION_URI, Package, time::Offset};
use litchi_ooxml::{OoxmlError, PackURI};
use tempfile::NamedTempFile;

const LOCAL_LASER: &[u8] = include_bytes!("../../../test-data/ooxml/pptx/laser/basic_laser.xml");

#[test]
fn package_inventory_reports_local_laser_traces() {
    let package = package_with_local_laser();

    let traces = package.laser_traces().unwrap();
    assert_eq!(traces.len(), 2);

    assert_eq!(traces[0].slide_index(), 0);
    assert_eq!(traces[0].trace_index(), 0);
    assert_eq!(traces[0].point_count(), 3);
    assert_eq!(traces[0].points()[0].time(), &Offset::ms(48796));
    assert_eq!(traces[0].points()[0].x().as_emu(), Some(6_062_662));
    assert_eq!(traces[0].points()[1].y().as_emu(), Some(3_179_762));
    assert_eq!(traces[0].points()[2].x().as_emu(), Some(0));

    assert_eq!(traces[1].slide_index(), 0);
    assert_eq!(traces[1].trace_index(), 1);
    assert_eq!(traces[1].point_count(), 2);
    assert_eq!(traces[1].points()[0].time(), &Offset::ms(52000));
    assert_eq!(traces[1].points()[1].y().as_emu(), Some(0));

    assert_eq!(
        package.presentation().unwrap().laser_traces().unwrap(),
        traces
    );
}

#[test]
fn package_inventory_rejects_malformed_laser_trace_points() {
    let malformed = format!(
        r#"<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:ext uri="{LASER_TRACE_EXTENSION_URI}"><p14:laserTraceLst><p14:tracePtLst><p14:tracePt t="not-a-time" x="0" y="0"/></p14:tracePtLst></p14:laserTraceLst></p:ext></p:extLst>"#
    );
    let package = package_with_laser_extension(&malformed);

    assert!(matches!(
        package.laser_traces(),
        Err(OoxmlError::Pptx(litchi_pptx::Error::Invalid(message)))
            if message.contains("universal time offset")
    ));
}

fn package_with_local_laser() -> Package {
    package_with_laser_extension(std::str::from_utf8(LOCAL_LASER).unwrap())
}

fn package_with_laser_extension(extension: &str) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    install_laser_extension(&mut package, extension);
    package
}

fn install_laser_extension(package: &mut Package, extension: &str) {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package
        .edit_opc(|opc| {
            let slide = opc.get_part_mut(&slide_name)?;
            let xml = std::str::from_utf8(slide.blob()).unwrap();
            let updated = xml.replacen("</p:sld>", &format!("{extension}</p:sld>"), 1);
            assert_ne!(updated, xml);
            slide.set_blob(updated.into_bytes());
            Ok(())
        })
        .unwrap();
}
