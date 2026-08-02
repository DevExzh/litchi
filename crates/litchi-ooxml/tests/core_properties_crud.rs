use litchi_ooxml::{Props, docx, pptx, xlsx};
use litchi_ooxml_common::properties;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use tempfile::NamedTempFile;

const STRICT_REL: &str =
    "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties";
const STRICT_NS: &str = "http://purl.oclc.org/ooxml/package/metadata/core-properties";
const CORE_PATH: &str = "/custom/MetaData.XML";

fn strict_source(mut package: OpcPackage) -> (OpcPackage, Vec<u8>) {
    properties::clear(&mut package).expect("remove template properties");
    let xml = format!(
        "<?xml version=\"1.0\"?>\n<cp:coreProperties xmlns:cp=\"{STRICT_NS}\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\n  <dc:title>Original</dc:title>\n</cp:coreProperties>"
    )
    .into_bytes();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new(CORE_PATH).expect("test URI"),
        ct::OPC_CORE_PROPERTIES.to_owned(),
        xml.clone(),
    )));
    package.relate_to(CORE_PATH.trim_start_matches('/'), STRICT_REL);
    (package, xml)
}

fn assert_exact_strict(package: &OpcPackage, expected: &[u8]) {
    let relationship = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == STRICT_REL)
        .expect("Strict core relationship");
    assert_eq!(relationship.target_ref(), CORE_PATH.trim_start_matches('/'));
    let target = relationship.target_partname().expect("internal target");
    let part = package.get_part(&target).expect("core part");
    assert_eq!(part.partname().as_str(), CORE_PATH);
    assert_eq!(part.blob(), expected);
}

fn assert_changed_strict(package: &OpcPackage) {
    let relationship = package
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == STRICT_REL)
        .expect("Strict core relationship");
    let target = relationship.target_partname().expect("internal target");
    let part = package.get_part(&target).expect("core part");
    assert_eq!(part.partname().as_str(), CORE_PATH);
    let xml = std::str::from_utf8(part.blob()).expect("UTF-8 core properties");
    assert!(xml.contains(STRICT_NS));
    assert!(xml.contains("<dc:title>Changed</dc:title>"));
    assert!(
        !package
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == rt::CORE_PROPERTIES)
    );
}

fn assert_absent(package: &OpcPackage) {
    assert!(
        properties::read(package)
            .expect("read properties")
            .is_none()
    );
    assert!(
        !package
            .iter_parts()
            .any(|part| part.content_type() == ct::OPC_CORE_PROPERTIES)
    );
}

#[test]
fn docx_host_preserves_updates_and_clears_core_properties() {
    let fresh = docx::Package::new().expect("fresh DOCX");
    let (source, exact) = strict_source(fresh.opc_package().clone());
    let mut package = docx::Package::from_opc_package(source).expect("open DOCX");
    assert_eq!(
        package.props().and_then(|props| props.title.as_deref()),
        Some("Original")
    );
    assert!(package.props_mut().is_some());

    let clean = NamedTempFile::with_suffix(".docx").expect("temp DOCX");
    package.save(clean.path()).expect("clean save");
    let clean_opc = OpcPackage::open(clean.path()).expect("saved OPC");
    assert_exact_strict(&clean_opc, &exact);

    let mut changed = docx::Package::open(clean.path()).expect("reopen DOCX");
    changed.props_mut().expect("present properties").title = Some("Changed".to_owned());
    let changed_file = NamedTempFile::with_suffix(".docx").expect("temp DOCX");
    changed.save(changed_file.path()).expect("changed save");
    assert_changed_strict(&OpcPackage::open(changed_file.path()).expect("changed OPC"));

    let mut cleared = docx::Package::open(changed_file.path()).expect("reopen changed DOCX");
    assert!(cleared.clear_props().is_some());
    let cleared_file = NamedTempFile::with_suffix(".docx").expect("temp DOCX");
    cleared.save(cleared_file.path()).expect("clear save");
    assert_absent(&OpcPackage::open(cleared_file.path()).expect("cleared OPC"));

    let mut absent = docx::Package::open(cleared_file.path()).expect("open absent DOCX");
    assert!(absent.props().is_none());
    let absent_file = NamedTempFile::with_suffix(".docx").expect("temp DOCX");
    absent.save(absent_file.path()).expect("absent clean save");
    assert_absent(&OpcPackage::open(absent_file.path()).expect("absent OPC"));
}

#[test]
fn pptx_host_preserves_updates_and_clears_core_properties() {
    let fresh = pptx::Package::new().expect("fresh PPTX");
    let (source, exact) = strict_source(fresh.opc_package().clone());
    let mut package = pptx::Package::from_opc_package(source).expect("open PPTX");
    assert_eq!(
        package.props().and_then(|props| props.title.as_deref()),
        Some("Original")
    );
    assert!(package.props_mut().is_some());

    let clean = NamedTempFile::with_suffix(".pptx").expect("temp PPTX");
    package.save(clean.path()).expect("clean save");
    let clean_opc = OpcPackage::open(clean.path()).expect("saved OPC");
    assert_exact_strict(&clean_opc, &exact);

    let mut changed = pptx::Package::open(clean.path()).expect("reopen PPTX");
    changed.props_mut().expect("present properties").title = Some("Changed".to_owned());
    let changed_file = NamedTempFile::with_suffix(".pptx").expect("temp PPTX");
    changed.save(changed_file.path()).expect("changed save");
    assert_changed_strict(&OpcPackage::open(changed_file.path()).expect("changed OPC"));

    let mut cleared = pptx::Package::open(changed_file.path()).expect("reopen changed PPTX");
    assert!(cleared.clear_props().is_some());
    let cleared_file = NamedTempFile::with_suffix(".pptx").expect("temp PPTX");
    cleared.save(cleared_file.path()).expect("clear save");
    assert_absent(&OpcPackage::open(cleared_file.path()).expect("cleared OPC"));

    let mut absent = pptx::Package::open(cleared_file.path()).expect("open absent PPTX");
    assert!(absent.props().is_none());
    let absent_file = NamedTempFile::with_suffix(".pptx").expect("temp PPTX");
    absent.save(absent_file.path()).expect("absent clean save");
    assert_absent(&OpcPackage::open(absent_file.path()).expect("absent OPC"));
}

#[test]
fn xlsx_host_preserves_updates_and_clears_core_properties() {
    let fresh = xlsx::Workbook::create().expect("fresh XLSX");
    let (source, exact) = strict_source(fresh.opc_package().clone());
    let mut workbook = xlsx::Workbook::new(source).expect("open XLSX");
    assert_eq!(
        workbook.props().and_then(|props| props.title.as_deref()),
        Some("Original")
    );
    assert!(workbook.props_mut().is_some());

    let clean = NamedTempFile::with_suffix(".xlsx").expect("temp XLSX");
    workbook.save(clean.path()).expect("clean save");
    let clean_opc = OpcPackage::open(clean.path()).expect("saved OPC");
    assert_exact_strict(&clean_opc, &exact);

    let mut changed = xlsx::Workbook::open(clean.path()).expect("reopen XLSX");
    changed.props_mut().expect("present properties").title = Some("Changed".to_owned());
    let changed_file = NamedTempFile::with_suffix(".xlsx").expect("temp XLSX");
    changed.save(changed_file.path()).expect("changed save");
    assert_changed_strict(&OpcPackage::open(changed_file.path()).expect("changed OPC"));

    let mut cleared = xlsx::Workbook::open(changed_file.path()).expect("reopen changed XLSX");
    assert!(cleared.clear_props().is_some());
    let cleared_file = NamedTempFile::with_suffix(".xlsx").expect("temp XLSX");
    cleared.save(cleared_file.path()).expect("clear save");
    assert_absent(&OpcPackage::open(cleared_file.path()).expect("cleared OPC"));

    let mut absent = xlsx::Workbook::open(cleared_file.path()).expect("open absent XLSX");
    assert!(absent.props().is_none());
    let absent_file = NamedTempFile::with_suffix(".xlsx").expect("temp XLSX");
    absent.save(absent_file.path()).expect("absent clean save");
    assert_absent(&OpcPackage::open(absent_file.path()).expect("absent OPC"));
}

#[test]
fn facade_put_moves_a_new_value_without_exposing_graph_ids() {
    let mut package = docx::Package::new().expect("fresh DOCX");
    assert_eq!(package.props().cloned(), Some(Props::new()));
    let previous = package.put_props(Props::new().title("Moved"));
    assert!(previous.is_some());
    assert_eq!(
        package.props().and_then(|props| props.title.as_deref()),
        Some("Moved")
    );
}

#[test]
fn fresh_hosts_omit_invented_history() {
    let docx = docx::Package::new().expect("fresh DOCX");
    let pptx = pptx::Package::new().expect("fresh PPTX");
    let xlsx = xlsx::Workbook::create().expect("fresh XLSX");

    assert_eq!(docx.props().cloned(), Some(Props::new()));
    assert_eq!(pptx.props().cloned(), Some(Props::new()));
    assert_eq!(xlsx.props().cloned(), Some(Props::new()));
}
