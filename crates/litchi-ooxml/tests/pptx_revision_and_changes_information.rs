use litchi_ooxml::OoxmlError;
use litchi_ooxml::pptx::{
    ChangesInformation, ChangesInformationPart, Package, RevisionInformation,
    RevisionInformationPart,
};
use tempfile::NamedTempFile;

const LOCAL_REVISION_INFORMATION: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/revision-information/basic_revision.xml");
const LOCAL_CHANGES_INFORMATION: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/changes-information/basic_changes.xml");

#[test]
fn package_and_presentation_manage_local_revision_and_changes_information() {
    let mut package = Package::new().unwrap();
    assert_eq!(package.revision_information().unwrap(), None);
    assert_eq!(package.changes_information().unwrap(), None);

    let revision = RevisionInformationPart {
        relationship_id: "rIdRevisionInformation".to_string(),
        part_name: "/ppt/revisionInfo.xml".to_string(),
        revision_information: RevisionInformation::parse(LOCAL_REVISION_INFORMATION).unwrap(),
    };
    let changes = ChangesInformationPart {
        relationship_id: "rIdChangesInformation".to_string(),
        part_name: "/ppt/changesInfo.xml".to_string(),
        changes_information: ChangesInformation::parse(LOCAL_CHANGES_INFORMATION).unwrap(),
    };

    package.store_revision_information(&revision).unwrap();
    package.store_changes_information(&changes).unwrap();
    assert!(matches!(
        package.store_revision_information(&revision),
        Err(OoxmlError::InvalidFormat(_))
    ));
    assert!(matches!(
        package.store_changes_information(&changes),
        Err(OoxmlError::InvalidFormat(_))
    ));

    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    package.save(output.path()).unwrap();
    let package = Package::open(output.path()).unwrap();

    assert_eq!(
        package.revision_information().unwrap(),
        Some(revision.clone())
    );
    assert_eq!(
        package.changes_information().unwrap(),
        Some(changes.clone())
    );
    let presentation = package.presentation().unwrap();
    assert_eq!(presentation.revision_information().unwrap(), Some(revision));
    assert_eq!(presentation.changes_information().unwrap(), Some(changes));
}
