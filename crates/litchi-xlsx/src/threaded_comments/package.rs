//! Threaded-comment reads of a worksheet and the workbook's people
//! (unified-ops record 0322): a worksheet names its threaded-comments part,
//! the workbook names its persons part. Only those relationships and parts
//! are read.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Relationships};

use super::{Comments, MAX_PART_BYTES, People, parse_comments, parse_persons};
use crate::error::{Result, invalid};

/// The one internal relationship of `kind` among `relationships`.
fn target(relationships: &Relationships, kind: &str, what: &str) -> Result<Option<PackURI>> {
    let mut found = relationships
        .iter()
        .filter(|relationship| relationship.reltype() == kind);
    let Some(first) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid(format!("more than one {what} relationship")));
    }
    if first.is_external() {
        return Err(invalid(format!("a {what} relationship cannot be external")));
    }
    Ok(Some(first.target_partname()?))
}

fn text<'a>(bytes: &'a [u8], content_type: &str, expected: &str, what: &str) -> Result<&'a str> {
    if content_type != expected {
        return Err(invalid(format!(
            "the {what} target has another content type"
        )));
    }
    if bytes.len() > MAX_PART_BYTES {
        return Err(invalid(format!("the {what} part exceeds the size limit")));
    }
    std::str::from_utf8(bytes).map_err(|_| invalid(format!("the {what} part is not UTF-8")))
}

/// The threaded comments of a worksheet of an eager package.
pub(crate) fn comments(package: &OpcPackage, worksheet: &PackURI) -> Result<Option<Comments>> {
    let Some(name) = target(
        package.get_part(worksheet)?.rels(),
        rt::THREADED_COMMENTS,
        "threaded-comments",
    )?
    else {
        return Ok(None);
    };
    let part = package.get_part(&name)?;
    let xml = text(
        part.blob(),
        part.content_type(),
        ct::SML_THREADED_COMMENTS,
        "threaded-comments",
    )?;
    Ok(Some(
        parse_comments(xml).map_err(|error| invalid(error.to_string()))?,
    ))
}

/// The threaded comments of a worksheet of a source-backed package.
pub(crate) fn source_comments(
    package: &litchi_opc::SourceBackedPackage,
    worksheet: &PackURI,
) -> Result<Option<Comments>> {
    package.check_execution()?;
    let Some(name) = target(
        package.part(worksheet)?.rels(),
        rt::THREADED_COMMENTS,
        "threaded-comments",
    )?
    else {
        return Ok(None);
    };
    let part = package.part(&name)?;
    if part.declared_uncompressed_size()? > MAX_PART_BYTES as u64 {
        return Err(invalid("the threaded-comments part exceeds the size limit"));
    }
    let data = part.data()?;
    let xml = text(
        data.as_bytes(),
        part.content_type(),
        ct::SML_THREADED_COMMENTS,
        "threaded-comments",
    )?;
    Ok(Some(
        parse_comments(xml).map_err(|error| invalid(error.to_string()))?,
    ))
}

/// The people of an eager package's workbook.
pub(crate) fn people(package: &OpcPackage) -> Result<Option<People>> {
    let workbook = package.main_document_part()?;
    let Some(name) = target(workbook.rels(), rt::PERSONS, "persons")? else {
        return Ok(None);
    };
    let part = package.get_part(&name)?;
    let xml = text(part.blob(), part.content_type(), ct::SML_PERSONS, "persons")?;
    Ok(Some(
        parse_persons(xml).map_err(|error| invalid(error.to_string()))?,
    ))
}

/// The people of a source-backed package's workbook.
pub(crate) fn source_people(package: &litchi_opc::SourceBackedPackage) -> Result<Option<People>> {
    package.check_execution()?;
    let workbook = package.main_document_part()?;
    let Some(name) = target(workbook.rels(), rt::PERSONS, "persons")? else {
        return Ok(None);
    };
    let part = package.part(&name)?;
    if part.declared_uncompressed_size()? > MAX_PART_BYTES as u64 {
        return Err(invalid("the persons part exceeds the size limit"));
    }
    let data = part.data()?;
    let xml = text(
        data.as_bytes(),
        part.content_type(),
        ct::SML_PERSONS,
        "persons",
    )?;
    Ok(Some(
        parse_persons(xml).map_err(|error| invalid(error.to_string()))?,
    ))
}

#[cfg(test)]
mod tests {
    #[test]
    fn worksheets_and_people_read_on_both_owners() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/right-to-left.xlsx"
        );
        let bytes = std::fs::read(path).unwrap();
        let package = crate::Package::from_bytes(&bytes).unwrap();
        let workbook = package.workbook();
        let sheet = workbook.sheet(0).unwrap().unwrap();
        let comments = sheet.threaded_comments().unwrap().unwrap();
        assert_eq!(comments.comments[0].cell_ref.as_deref(), Some("A3"));
        assert_eq!(comments.comments[0].text.as_deref(), Some("test comment"));
        let people = workbook.people().unwrap().unwrap();
        assert_eq!(people.persons[0].display_name, "PJ Fanning");
        assert_eq!(people.persons[0].id, comments.comments[0].person_id);

        let source = crate::workbook::source::SourceBackedWorkbook::open(path).unwrap();
        let sheet = source.sheet(0).unwrap().unwrap();
        let read = sheet.threaded_comments().unwrap().unwrap();
        assert_eq!(read.comments.len(), comments.comments.len());
        assert_eq!(read.comments[0].id, comments.comments[0].id);
        assert_eq!(
            source.people().unwrap().unwrap().persons[0].display_name,
            "PJ Fanning"
        );
    }
}
