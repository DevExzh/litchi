//! Worksheet OPC graph lifecycle for classic comments parts.

use std::collections::HashSet;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

use crate::error::{Result, invalid};

use super::codec::{parse_comments, validate_comments, write_comments};
use super::model::{Comments, Part as CommentsPart};

/// Content type of a classic SpreadsheetML comments part.
pub const COMMENTS_CONTENT_TYPE: &str = ct::SML_COMMENTS;
/// Transitional worksheet-to-comments relationship type.
pub const COMMENTS_RELATIONSHIP_TYPE: &str = rt::COMMENTS;
/// ISO/IEC 29500 Strict worksheet-to-comments relationship type.
pub const STRICT_COMMENTS_RELATIONSHIP_TYPE: &str = rt::STRICT_COMMENTS;

/// Load the optional classic comments part owned by one worksheet.
pub fn load_from_worksheet(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<Option<CommentsPart>> {
    let worksheet = package.get_part(worksheet_part)?;
    require_worksheet(worksheet.content_type())?;
    let relationship = find_relationship(worksheet)?;
    let Some(relationship) = relationship else {
        return Ok(None);
    };
    let comments_name = relationship.target_partname()?;
    let comments = package.get_part(&comments_name)?;
    validate_comments_part(comments)?;
    let xml = std::str::from_utf8(comments.blob())
        .map_err(|error| invalid(format!("classic comments XML is not UTF-8: {error}")))?;
    Ok(Some(CommentsPart {
        worksheet_part_name: worksheet_part.to_string(),
        relationship_id: relationship.r_id().to_owned(),
        part_name: comments_name.to_string(),
        comments: parse_comments(xml)?,
    }))
}

/// Store or replace the worksheet's sole classic comments part.
pub fn store_on_worksheet(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    value: &Comments,
) -> Result<CommentsPart> {
    validate_comments(value)?;
    let relationship = find_relationship(package.get_part(worksheet_part)?)?
        .map(|relationship| {
            Ok::<_, crate::Error>((
                relationship.r_id().to_owned(),
                relationship.target_partname()?,
                relationship.reltype().to_owned(),
            ))
        })
        .transpose()?;
    let (relationship_id, part_name, relationship_type) =
        if let Some((relationship_id, part_name, relationship_type)) = relationship {
            validate_comments_part(package.get_part(&part_name)?)?;
            (relationship_id, part_name, relationship_type)
        } else {
            let part_name = next_part_name(package)?;
            let relationship_id = next_relationship_id(package, worksheet_part)?;
            (
                relationship_id,
                part_name,
                COMMENTS_RELATIONSHIP_TYPE.to_owned(),
            )
        };

    let xml = write_comments(value)?;
    if package.get_part(&part_name).is_ok() {
        package.get_part_mut(&part_name)?.set_blob(xml);
    } else {
        package.try_add_part(Box::new(BlobPart::new(
            part_name.clone(),
            COMMENTS_CONTENT_TYPE.into(),
            xml,
        )))?;
    }
    if package
        .get_part(worksheet_part)?
        .rels()
        .get(&relationship_id)
        .is_none()
    {
        package
            .get_part_mut(worksheet_part)?
            .rels_mut()
            .add_relationship(
                relationship_type,
                part_name.relative_ref(worksheet_part.base_uri()),
                relationship_id.clone(),
                false,
            );
    }
    package.unsign();
    Ok(CommentsPart {
        worksheet_part_name: worksheet_part.to_string(),
        relationship_id,
        part_name: part_name.to_string(),
        comments: value.clone(),
    })
}

/// Replace the worksheet's classic comments graph atomically at the semantic level.
pub fn replace_on_worksheet(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    value: &Comments,
) -> Result<CommentsPart> {
    store_on_worksheet(package, worksheet_part, value)
}

/// Remove the worksheet's classic comments relationship and unreferenced part.
pub fn remove_from_worksheet(package: &mut OpcPackage, worksheet_part: &PackURI) -> Result<bool> {
    let Some((relationship_id, part_name)) = find_relationship(package.get_part(worksheet_part)?)?
        .map(|relationship| {
            Ok::<_, crate::Error>((
                relationship.r_id().to_owned(),
                relationship.target_partname()?,
            ))
        })
        .transpose()?
    else {
        return Ok(false);
    };
    package
        .get_part_mut(worksheet_part)?
        .rels_mut()
        .remove(&relationship_id);
    if !part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    package.unsign();
    Ok(true)
}

/// Validate every classic comments relationship and reject orphan resources.
pub fn validate_graph(package: &OpcPackage) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_comments_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "package root cannot source a worksheet comments relationship",
        ));
    }

    let mut targets = HashSet::new();
    for source in package.iter_parts() {
        let mut relationships = source
            .rels()
            .iter()
            .filter(|relationship| is_comments_relationship(relationship.reltype()));
        let Some(relationship) = relationships.next() else {
            continue;
        };
        require_worksheet(source.content_type())?;
        if relationships.next().is_some() {
            return Err(invalid(format!(
                "worksheet '{}' has multiple classic comments relationships",
                source.partname()
            )));
        }
        if relationship.is_external() {
            return Err(invalid(
                "worksheet comments relationship cannot be external",
            ));
        }
        let target = relationship.target_partname()?;
        if !targets.insert(target.clone()) {
            return Err(invalid(format!(
                "classic comments part '{target}' is targeted more than once"
            )));
        }
        validate_comments_part(package.get_part(&target)?)?;
    }

    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == COMMENTS_CONTENT_TYPE)
    {
        if !targets.contains(part.partname()) {
            return Err(invalid(format!(
                "classic comments part '{}' has no worksheet relationship",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn find_relationship<'a>(worksheet: &'a dyn Part) -> Result<Option<&'a litchi_opc::Relationship>> {
    let mut relationships = worksheet
        .rels()
        .iter()
        .filter(|relationship| is_comments_relationship(relationship.reltype()));
    let first = relationships.next();
    if relationships.next().is_some() {
        return Err(invalid(
            "worksheet has multiple classic comments relationships",
        ));
    }
    Ok(first)
}

fn is_comments_relationship(value: &str) -> bool {
    matches!(
        value,
        COMMENTS_RELATIONSHIP_TYPE | STRICT_COMMENTS_RELATIONSHIP_TYPE
    )
}

fn require_worksheet(content_type: &str) -> Result<()> {
    if content_type == ct::SML_WORKSHEET {
        Ok(())
    } else {
        Err(invalid(format!(
            "classic comments operations require a worksheet part, got '{content_type}'"
        )))
    }
}

fn validate_comments_part(part: &dyn Part) -> Result<()> {
    if part.content_type() != COMMENTS_CONTENT_TYPE {
        return Err(invalid(format!(
            "comments part '{}' has content type '{}', expected '{}'",
            part.partname(),
            part.content_type(),
            COMMENTS_CONTENT_TYPE
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid(
            "classic comments parts must not have relationships",
        ));
    }
    Ok(())
}

fn next_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 1..=65_536u32 {
        let candidate = PackURI::new(format!("/xl/comments{suffix}.xml"))
            .map_err(|error| invalid(error.to_string()))?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free classic comments part name"))
}

fn next_relationship_id(package: &OpcPackage, worksheet_part: &PackURI) -> Result<String> {
    let relationships = package.get_part(worksheet_part)?.rels();
    for suffix in 1..=65_536u32 {
        let candidate = format!("rIdComments{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free classic comments relationship ID"))
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name == *target)
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeMap;

    fn fixture() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        let part = BlobPart::new(
            worksheet.clone(),
            ct::SML_WORKSHEET.into(),
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
                .to_vec(),
        );
        package.add_part(Box::new(part));
        (package, worksheet)
    }

    fn comments() -> Comments {
        let mut comments = Comments {
            authors: vec!["Alice".into()],
            comments: BTreeMap::new(),
        };
        comments.comments.insert(
            "A1".into(),
            super::super::model::Comment {
                cell_ref: "A1".into(),
                author: "Alice".into(),
                author_id: 0,
                text: "note".into(),
                guid: None,
                shape_id: None,
            },
        );
        comments
    }

    #[test]
    fn stores_loads_and_removes_worksheet_graph() {
        let (mut package, worksheet) = fixture();
        let stored = store_on_worksheet(&mut package, &worksheet, &comments()).unwrap();
        assert_eq!(stored.relationship_id, "rIdComments1");
        validate_graph(&package).unwrap();
        let loaded = load_from_worksheet(&package, &worksheet).unwrap().unwrap();
        assert_eq!(loaded.comments, comments());
        assert!(remove_from_worksheet(&mut package, &worksheet).unwrap());
        assert!(load_from_worksheet(&package, &worksheet).unwrap().is_none());
        assert!(
            !package
                .iter_parts()
                .any(|part| part.content_type() == COMMENTS_CONTENT_TYPE)
        );
    }
}
