//! OPC/MCE adapter for the package-neutral threaded-comments codec.

use litchi_core::sheet::Result as SheetResult;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Relationships};

use litchi_xlsx::threaded_comments::{
    Comments, MAX_PART_BYTES, People, parse_comments, parse_persons,
};

/// Read the person list related to the workbook's actual main-document part.
pub(crate) fn read_persons(package: &OpcPackage) -> SheetResult<Option<People>> {
    let workbook_part = package.main_document_part()?;
    let Some(persons_uri) = related_part_uri(workbook_part.rels(), rt::PERSONS, "people")? else {
        return Ok(None);
    };
    let persons_part = package.get_part(&persons_uri)?;
    require_content_type(&persons_uri, persons_part.content_type(), ct::SML_PERSONS)?;
    if persons_part.blob().len() > MAX_PART_BYTES {
        return Err("persons part exceeds the configured resource bound".into());
    }
    let bytes = litchi_ooxml_common::mce::process_part(persons_part)?;
    let xml = std::str::from_utf8(bytes.as_ref())?;
    Ok(Some(parse_persons(xml)?))
}

/// Read the threaded-comments part related to a worksheet.
pub(crate) fn read_threaded_comments(
    package: &OpcPackage,
    worksheet_uri: &PackURI,
) -> SheetResult<Option<Comments>> {
    let worksheet_part = package.get_part(worksheet_uri)?;
    let Some(comments_uri) = related_part_uri(
        worksheet_part.rels(),
        rt::THREADED_COMMENTS,
        "threaded comments",
    )?
    else {
        return Ok(None);
    };
    let comments_part = package.get_part(&comments_uri)?;
    require_content_type(
        &comments_uri,
        comments_part.content_type(),
        ct::SML_THREADED_COMMENTS,
    )?;
    if comments_part.blob().len() > MAX_PART_BYTES {
        return Err("threaded-comments part exceeds the configured resource bound".into());
    }
    let bytes = litchi_ooxml_common::mce::process_part(comments_part)?;
    let xml = std::str::from_utf8(bytes.as_ref())?;
    Ok(Some(parse_comments(xml)?))
}

fn related_part_uri(
    relationships: &Relationships,
    relationship_type: &str,
    description: &str,
) -> SheetResult<Option<PackURI>> {
    let mut matching = relationships
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type);
    let Some(relationship) = matching.next() else {
        return Ok(None);
    };
    if matching.next().is_some() {
        return Err(format!("part has multiple {description} relationships").into());
    }
    if relationship.is_external() {
        return Err(format!("{description} relationship cannot be external").into());
    }
    Ok(Some(relationship.target_partname()?))
}

fn require_content_type(uri: &PackURI, actual: &str, expected: &str) -> SheetResult<()> {
    if actual != expected {
        return Err(
            format!("part '{uri}' has content type '{actual}', expected '{expected}'").into(),
        );
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::part::BlobPart;
    use litchi_opc::{OpcPackage, PackURI, Part};

    use super::*;

    const NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

    fn package_with_threaded_parts() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
        let worksheet_uri = PackURI::new("/custom/sheets/sheet.xml").unwrap();
        let mut workbook_part =
            BlobPart::new(workbook_uri, ct::SML_SHEET_MAIN.to_string(), Vec::new());
        workbook_part.relate_to("people.xml", rt::PERSONS);
        let mut worksheet_part = BlobPart::new(
            worksheet_uri.clone(),
            ct::SML_WORKSHEET.to_string(),
            Vec::new(),
        );
        worksheet_part.relate_to("../threads.xml", rt::THREADED_COMMENTS);
        package.relate_to("custom/book.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(worksheet_part));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/people.xml").unwrap(),
            ct::SML_PERSONS.to_string(),
            format!(r#"<personList xmlns="{NS}"/>"#).into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/threads.xml").unwrap(),
            ct::SML_THREADED_COMMENTS.to_string(),
            format!(r#"<ThreadedComments xmlns="{NS}"/>"#).into_bytes(),
        )));
        (package, worksheet_uri)
    }

    #[test]
    fn resolves_custom_part_locations() {
        let (package, worksheet_uri) = package_with_threaded_parts();

        assert!(read_persons(&package).unwrap().is_some());
        assert!(
            read_threaded_comments(&package, &worksheet_uri)
                .unwrap()
                .is_some()
        );
    }

    #[test]
    fn rejects_external_duplicate_and_wrong_content_type_relationships() {
        let (mut package, worksheet_uri) = package_with_threaded_parts();
        let relationships = package.get_part_mut(&worksheet_uri).unwrap().rels_mut();
        relationships.remove("rId1").unwrap();
        relationships.add_relationship(
            rt::THREADED_COMMENTS.to_string(),
            "https://example.com/thread.xml".to_string(),
            "rId1".to_string(),
            true,
        );
        assert!(read_threaded_comments(&package, &worksheet_uri).is_err());

        let (mut package, worksheet_uri) = package_with_threaded_parts();
        package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::THREADED_COMMENTS.to_string(),
                "https://example.com/thread.xml".to_string(),
                "rId2".to_string(),
                true,
            );
        assert!(read_threaded_comments(&package, &worksheet_uri).is_err());

        let (mut package, worksheet_uri) = package_with_threaded_parts();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/threads.xml").unwrap(),
            ct::SML_PERSONS.to_string(),
            Vec::new(),
        )));
        assert!(read_threaded_comments(&package, &worksheet_uri).is_err());
    }
}
