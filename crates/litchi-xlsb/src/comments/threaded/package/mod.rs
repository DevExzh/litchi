//! OPC part and relationship lifecycle for XLSB threaded comments.
//!
//! The binary worksheet and legacy comments*.bin parts are deliberately
//! outside this owner. These operations only create, replace, inspect, or
//! remove the XML parts described by MS-XLSB 2.1.17--2.1.18. Mutations stage a
//! cloned OPC graph and publish it only after relationship and semantic
//! validation accepts the candidate.

use std::collections::HashSet;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};

use crate::package::error::{Error, Result};

use super::codec::{parse_comments, parse_persons, write_comments, write_persons};
use super::semantic::{Comments, CommentsPart, Graph, People, PeoplePart};
use super::validation::{validate_graph as validate_model_graph, validate_people};

/// Content type of an XLSB threaded-comments XML part.
pub const COMMENTS_CONTENT_TYPE: &str = ct::SML_THREADED_COMMENTS;
/// Content type of an XLSB persons XML part.
pub const PERSONS_CONTENT_TYPE: &str = ct::SML_PERSONS;
/// Relationship from one XLSB worksheet to its threaded-comments part.
pub const COMMENTS_RELATIONSHIP_TYPE: &str = rt::THREADED_COMMENTS;
/// Relationship from the XLSB workbook to its persons part.
pub const PERSONS_RELATIONSHIP_TYPE: &str = rt::PERSONS;
/// Content type used by BIFF12 worksheet parts in the XLSB package graph.
pub const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";

const PERSONS_PART_TEMPLATE: &str = "/xl/persons/person%d.xml";
const COMMENTS_PART_TEMPLATE: &str = "/xl/threadedComments/threadedComment%d.xml";
const MAX_NAME_ATTEMPTS: u32 = 65_536;

/// Read every typed threaded-comments/person part in an OPC package.
pub fn load_graph(package: &OpcPackage) -> Result<Graph> {
    validate_package_graph(package)?;
    let workbook = workbook_uri(package)?;
    let persons = load_people(package, &workbook)?;
    let mut worksheets = Vec::new();
    let mut worksheet_uris: Vec<PackURI> = package
        .iter_parts()
        .filter(|part| part.content_type() == WORKSHEET_CONTENT_TYPE)
        .map(|part| part.partname().clone())
        .collect();
    worksheet_uris.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    for worksheet in worksheet_uris {
        if let Some(comments) = load_from_worksheet(package, &worksheet)? {
            worksheets.push(comments);
        }
    }
    Ok(Graph {
        persons,
        worksheets,
    })
}

/// Replace the complete typed threaded-comments graph atomically.
pub fn store_graph(package: &mut OpcPackage, graph: &Graph) -> Result<()> {
    validate_package_graph(package)?;
    validate_model_graph(graph).map_err(map_model_error)?;
    let mut candidate = package.clone();
    remove_graph_inner(&mut candidate)?;
    let workbook = workbook_uri(&candidate)?;
    if let Some(persons) = graph.persons.as_ref() {
        store_people_part(&mut candidate, &workbook, persons)?;
    }
    for worksheet in &graph.worksheets {
        let worksheet_uri = PackURI::new(&worksheet.worksheet_part_name)
            .map_err(|error| Error::InvalidUri(error.to_string()))?;
        store_comments_part(&mut candidate, &worksheet_uri, worksheet)?;
    }
    validate_package_graph(&candidate)?;
    *package = candidate;
    Ok(())
}

/// Remove every threaded-comments/person relationship and unreferenced part.
pub fn remove_graph(package: &mut OpcPackage) -> Result<bool> {
    validate_package_graph(package)?;
    let mut candidate = package.clone();
    let changed = remove_graph_inner(&mut candidate)?;
    validate_package_graph(&candidate)?;
    *package = candidate;
    Ok(changed)
}

/// Validate all XLSB threaded-comments/person relationships and their XML.
pub fn validate_graph(package: &OpcPackage) -> Result<()> {
    validate_package_graph(package)
}

/// Load the optional workbook persons part.
pub fn load_people(package: &OpcPackage, workbook_part: &PackURI) -> Result<Option<PeoplePart>> {
    require_workbook(package.get_part(workbook_part)?)?;
    let Some((relationship_id, part_name)) = find_relationship(
        package.get_part(workbook_part)?,
        PERSONS_RELATIONSHIP_TYPE,
        "persons",
    )?
    else {
        return Ok(None);
    };
    let part = package.get_part(&part_name)?;
    validate_xml_part(part, PERSONS_CONTENT_TYPE, "persons")?;
    let persons = parse_persons(part.blob()).map_err(map_model_error)?;
    Ok(Some(PeoplePart {
        relationship_id,
        part_name: part_name.to_string(),
        persons,
    }))
}

/// Store or replace the workbook persons part using a bounded canonical name.
pub fn store_people(
    package: &mut OpcPackage,
    workbook_part: &PackURI,
    people: &People,
) -> Result<PeoplePart> {
    validate_package_graph(package)?;
    let mut candidate = package.clone();
    let result = store_people_part(
        &mut candidate,
        workbook_part,
        &PeoplePart {
            relationship_id: String::new(),
            part_name: String::new(),
            persons: people.clone(),
        },
    )?;
    validate_package_graph(&candidate)?;
    *package = candidate;
    Ok(result)
}

/// Remove the workbook persons relationship and unreferenced part.
pub fn remove_people(package: &mut OpcPackage, workbook_part: &PackURI) -> Result<bool> {
    validate_package_graph(package)?;
    let Some((relationship_id, part_name)) = find_relationship(
        package.get_part(workbook_part)?,
        PERSONS_RELATIONSHIP_TYPE,
        "persons",
    )?
    else {
        return Ok(false);
    };
    let mut candidate = package.clone();
    candidate
        .get_part_mut(workbook_part)?
        .rels_mut()
        .remove(&relationship_id);
    remove_if_unreferenced(&mut candidate, &part_name);
    validate_package_graph(&candidate)?;
    *package = candidate;
    Ok(true)
}

/// Load the optional threaded-comments part owned by one XLSB worksheet.
pub fn load_from_worksheet(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<Option<CommentsPart>> {
    require_worksheet(package.get_part(worksheet_part)?)?;
    let Some((relationship_id, part_name)) = find_relationship(
        package.get_part(worksheet_part)?,
        COMMENTS_RELATIONSHIP_TYPE,
        "threaded comments",
    )?
    else {
        return Ok(None);
    };
    let part = package.get_part(&part_name)?;
    validate_xml_part(part, COMMENTS_CONTENT_TYPE, "threaded comments")?;
    let comments = parse_comments(part.blob()).map_err(map_model_error)?;
    Ok(Some(CommentsPart {
        worksheet_part_name: worksheet_part.to_string(),
        relationship_id,
        part_name: part_name.to_string(),
        comments,
    }))
}

/// Store or replace one worksheet's threaded-comments part.
pub fn store_on_worksheet(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    comments: &Comments,
) -> Result<CommentsPart> {
    validate_package_graph(package)?;
    let mut candidate = package.clone();
    let result = store_comments_part(
        &mut candidate,
        worksheet_part,
        &CommentsPart {
            worksheet_part_name: worksheet_part.to_string(),
            relationship_id: String::new(),
            part_name: String::new(),
            comments: comments.clone(),
        },
    )?;
    validate_package_graph(&candidate)?;
    *package = candidate;
    Ok(result)
}

/// Remove one worksheet's threaded-comments relationship and unreferenced part.
pub fn remove_from_worksheet(package: &mut OpcPackage, worksheet_part: &PackURI) -> Result<bool> {
    validate_package_graph(package)?;
    require_worksheet(package.get_part(worksheet_part)?)?;
    let Some((relationship_id, part_name)) = find_relationship(
        package.get_part(worksheet_part)?,
        COMMENTS_RELATIONSHIP_TYPE,
        "threaded comments",
    )?
    else {
        return Ok(false);
    };
    package
        .get_part_mut(worksheet_part)?
        .rels_mut()
        .remove(&relationship_id);
    remove_if_unreferenced(package, &part_name);
    package.unsign();
    Ok(true)
}

fn store_people_part(
    package: &mut OpcPackage,
    workbook_part: &PackURI,
    value: &PeoplePart,
) -> Result<PeoplePart> {
    require_workbook(package.get_part(workbook_part)?)?;
    validate_people(&value.persons).map_err(map_model_error)?;
    let payload = write_persons(&value.persons).map_err(map_model_error)?;
    let existing = find_relationship(
        package.get_part(workbook_part)?,
        PERSONS_RELATIONSHIP_TYPE,
        "persons",
    )?;
    let (old_relationship_id, old_part) = existing
        .map(|(id, name)| (id, Some(name)))
        .unwrap_or_else(|| (String::new(), None));
    let part_name = choose_part_name(
        package,
        if value.part_name.is_empty() {
            None
        } else {
            Some(value.part_name.as_str())
        },
        PERSONS_PART_TEMPLATE,
    )?;
    let relationship_id = choose_relationship_id(
        package.get_part(workbook_part)?.rels(),
        if old_relationship_id.is_empty() {
            (!value.relationship_id.is_empty()).then_some(value.relationship_id.as_str())
        } else {
            Some(old_relationship_id.as_str())
        },
        "rIdPersons",
    )?;
    replace_part_and_relationship(
        package,
        workbook_part,
        old_part.as_ref(),
        &part_name,
        &relationship_id,
        PERSONS_CONTENT_TYPE,
        PERSONS_RELATIONSHIP_TYPE,
        payload,
    )?;
    Ok(PeoplePart {
        relationship_id,
        part_name: part_name.to_string(),
        persons: value.persons.clone(),
    })
}

fn store_comments_part(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    value: &CommentsPart,
) -> Result<CommentsPart> {
    require_worksheet(package.get_part(worksheet_part)?)?;
    let payload = write_comments(&value.comments).map_err(map_model_error)?;
    let existing = find_relationship(
        package.get_part(worksheet_part)?,
        COMMENTS_RELATIONSHIP_TYPE,
        "threaded comments",
    )?;
    let (old_relationship_id, old_part) = existing
        .map(|(id, name)| (id, Some(name)))
        .unwrap_or_else(|| (String::new(), None));
    let part_name = choose_part_name(
        package,
        if value.part_name.is_empty() {
            None
        } else {
            Some(value.part_name.as_str())
        },
        COMMENTS_PART_TEMPLATE,
    )?;
    let relationship_id = choose_relationship_id(
        package.get_part(worksheet_part)?.rels(),
        if old_relationship_id.is_empty() {
            (!value.relationship_id.is_empty()).then_some(value.relationship_id.as_str())
        } else {
            Some(old_relationship_id.as_str())
        },
        "rIdThreadedComments",
    )?;
    replace_part_and_relationship(
        package,
        worksheet_part,
        old_part.as_ref(),
        &part_name,
        &relationship_id,
        COMMENTS_CONTENT_TYPE,
        COMMENTS_RELATIONSHIP_TYPE,
        payload,
    )?;
    Ok(CommentsPart {
        worksheet_part_name: worksheet_part.to_string(),
        relationship_id,
        part_name: part_name.to_string(),
        comments: value.comments.clone(),
    })
}

fn validate_package_graph(package: &OpcPackage) -> Result<()> {
    if package.rels().iter().any(|relationship| {
        matches!(
            relationship.reltype(),
            PERSONS_RELATIONSHIP_TYPE | COMMENTS_RELATIONSHIP_TYPE
        )
    }) {
        return Err(invalid(
            "package root cannot source threaded-comments relationships",
        ));
    }

    let workbook = package
        .iter_parts()
        .find(|part| part.content_type() == ct::XLSB_BIN);
    let workbook_name = workbook.map(|part| part.partname().clone());
    let mut people_target = None;
    let mut comment_targets = HashSet::new();
    let mut comment_sources = HashSet::new();
    let mut graph = Graph::default();

    let mut sources: Vec<(PackURI, String, String, String, PackURI)> = Vec::new();
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if !matches!(
                relationship.reltype(),
                PERSONS_RELATIONSHIP_TYPE | COMMENTS_RELATIONSHIP_TYPE
            ) {
                continue;
            }
            if relationship.is_external() {
                return Err(invalid(format!(
                    "{} relationship cannot be external",
                    relationship.reltype()
                )));
            }
            let target = relationship.target_partname()?;
            sources.push((
                source.partname().clone(),
                source.content_type().to_owned(),
                relationship.reltype().to_owned(),
                relationship.r_id().to_owned(),
                target,
            ));
        }
    }

    for (source_name, source_type, relationship_type, relationship_id, target) in sources {
        let target_part = package.get_part(&target)?;
        if !target_part.rels().is_empty() {
            return Err(invalid(format!(
                "threaded-comments part '{}' must not have outgoing relationships",
                target
            )));
        }
        if relationship_type == PERSONS_RELATIONSHIP_TYPE {
            if source_type != ct::XLSB_BIN || workbook_name.as_ref() != Some(&source_name) {
                return Err(invalid(
                    "persons relationship must originate at the XLSB workbook",
                ));
            }
            if target_part.content_type() != PERSONS_CONTENT_TYPE {
                return Err(invalid(
                    "persons relationship targets the wrong content type",
                ));
            }
            if people_target.replace(target.clone()).is_some() {
                return Err(invalid("workbook has multiple persons relationships"));
            }
            let persons = parse_persons(target_part.blob()).map_err(map_model_error)?;
            graph.persons = Some(PeoplePart {
                relationship_id,
                part_name: target.to_string(),
                persons,
            });
        } else {
            if source_type != WORKSHEET_CONTENT_TYPE {
                return Err(invalid(
                    "threaded-comments relationship must originate at an XLSB worksheet",
                ));
            }
            if target_part.content_type() != COMMENTS_CONTENT_TYPE {
                return Err(invalid(
                    "threaded-comments relationship targets the wrong content type",
                ));
            }
            if !comment_sources.insert(source_name.clone()) {
                return Err(invalid(
                    "worksheet has multiple threaded-comments relationships",
                ));
            }
            if !comment_targets.insert(target.clone()) {
                return Err(invalid(format!(
                    "threaded-comments part '{}' is targeted more than once",
                    target
                )));
            }
            let comments = parse_comments(target_part.blob()).map_err(map_model_error)?;
            graph.worksheets.push(CommentsPart {
                worksheet_part_name: source_name.to_string(),
                relationship_id,
                part_name: target.to_string(),
                comments,
            });
        }
    }

    for part in package.iter_parts() {
        if part.content_type() == PERSONS_CONTENT_TYPE
            && people_target.as_ref() != Some(part.partname())
        {
            return Err(invalid(format!(
                "persons part '{}' has no workbook relationship",
                part.partname()
            )));
        }
        if part.content_type() == COMMENTS_CONTENT_TYPE
            && !comment_targets.contains(part.partname())
        {
            return Err(invalid(format!(
                "threaded-comments part '{}' has no worksheet relationship",
                part.partname()
            )));
        }
    }
    validate_model_graph(&graph).map_err(map_model_error)
}

fn remove_graph_inner(package: &mut OpcPackage) -> Result<bool> {
    let mut changed = false;
    let mut removals = Vec::new();
    let workbook = package
        .iter_parts()
        .find(|part| part.content_type() == ct::XLSB_BIN)
        .map(|part| part.partname().clone());
    let sources: Vec<PackURI> = package
        .iter_parts()
        .map(|part| part.partname().clone())
        .collect();
    for source in sources {
        let is_workbook = workbook.as_ref() == Some(&source);
        let part = package.get_part(&source)?;
        let ids: Vec<(String, PackURI)> = part
            .rels()
            .iter()
            .filter(|relationship| {
                (is_workbook && relationship.reltype() == PERSONS_RELATIONSHIP_TYPE)
                    || (!is_workbook && relationship.reltype() == COMMENTS_RELATIONSHIP_TYPE)
            })
            .map(|relationship| {
                Ok::<_, Error>((
                    relationship.r_id().to_owned(),
                    relationship.target_partname()?,
                ))
            })
            .collect::<Result<Vec<_>>>()?;
        if ids.is_empty() {
            continue;
        }
        let part_mut = package.get_part_mut(&source)?;
        for (relationship_id, target) in &ids {
            part_mut.rels_mut().remove(relationship_id);
            removals.push(target.clone());
            changed = true;
        }
    }
    for target in removals {
        remove_if_unreferenced(package, &target);
    }
    if changed {
        package.unsign();
    }
    Ok(changed)
}

fn replace_part_and_relationship(
    package: &mut OpcPackage,
    source: &PackURI,
    old_part: Option<&PackURI>,
    part_name: &PackURI,
    relationship_id: &str,
    content_type: &str,
    relationship_type: &str,
    payload: Vec<u8>,
) -> Result<()> {
    if let Some(old_part) = old_part
        && old_part != part_name
    {
        package
            .get_part_mut(source)?
            .rels_mut()
            .remove(relationship_id);
        remove_if_unreferenced(package, old_part);
    }
    if package.get_part(part_name).is_ok() {
        let part = package.get_part(part_name)?;
        validate_xml_part(part, content_type, "threaded-comments")?;
        package.get_part_mut(part_name)?.set_blob(payload);
    } else {
        package.validate_new_part_name(part_name)?;
        package.try_add_part(Box::new(BlobPart::new(
            part_name.clone(),
            content_type.to_owned(),
            payload,
        )))?;
    }
    let target = part_name.relative_ref(source.base_uri());
    let existing_relationship =
        package
            .get_part(source)?
            .rels()
            .get(relationship_id)
            .map(|relationship| {
                (
                    relationship.reltype().to_owned(),
                    relationship.is_external(),
                    relationship.target_partname(),
                )
            });
    if let Some((existing_type, external, existing_target)) = existing_relationship {
        if external || existing_type != relationship_type || existing_target? != *part_name {
            return Err(invalid(format!(
                "relationship ID '{relationship_id}' is already used by another target",
            )));
        }
    } else {
        package
            .get_part_mut(source)?
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_owned(),
                target,
                relationship_id.to_owned(),
                TargetMode::Internal,
            )?;
    }
    package.unsign();
    Ok(())
}

fn find_relationship(
    source: &dyn Part,
    relationship_type: &str,
    description: &str,
) -> Result<Option<(String, PackURI)>> {
    let mut found = None;
    for relationship in source
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type)
    {
        if found.is_some() {
            return Err(invalid(format!(
                "{} has multiple {} relationships",
                source.partname(),
                description
            )));
        }
        if relationship.is_external() {
            return Err(invalid(format!(
                "{description} relationship cannot be external"
            )));
        }
        found = Some((
            relationship.r_id().to_owned(),
            relationship.target_partname()?,
        ));
    }
    Ok(found)
}

fn validate_xml_part(part: &dyn Part, expected: &str, description: &str) -> Result<()> {
    if part.content_type() != expected {
        return Err(invalid(format!(
            "{description} part '{}' has content type '{}', expected '{}'",
            part.partname(),
            part.content_type(),
            expected
        )));
    }
    if !part.rels().is_empty() {
        return Err(invalid(format!(
            "{description} part '{}' must not have relationships",
            part.partname()
        )));
    }
    Ok(())
}

fn require_workbook(part: &dyn Part) -> Result<()> {
    if part.content_type() == ct::XLSB_BIN {
        Ok(())
    } else {
        Err(invalid(format!(
            "threaded-comment persons operations require an XLSB workbook, got '{}'",
            part.content_type()
        )))
    }
}

fn require_worksheet(part: &dyn Part) -> Result<()> {
    if part.content_type() == WORKSHEET_CONTENT_TYPE {
        Ok(())
    } else {
        Err(invalid(format!(
            "threaded-comment operations require an XLSB worksheet, got '{}'",
            part.content_type()
        )))
    }
}

fn workbook_uri(package: &OpcPackage) -> Result<PackURI> {
    if let Ok(main) = package.main_document_part()
        && main.content_type() == ct::XLSB_BIN
    {
        return Ok(main.partname().clone());
    }
    let candidate = PackURI::new("/xl/workbook.bin")?;
    require_workbook(package.get_part(&candidate)?)?;
    Ok(candidate)
}

fn choose_part_name(
    package: &OpcPackage,
    requested: Option<&str>,
    template: &str,
) -> Result<PackURI> {
    if let Some(requested) = requested {
        let part_name =
            PackURI::new(requested).map_err(|error| Error::InvalidUri(error.to_string()))?;
        if package.get_part(&part_name).is_ok() {
            return Ok(part_name);
        }
        package.validate_new_part_name(&part_name)?;
        return Ok(part_name);
    }
    for suffix in 1..=MAX_NAME_ATTEMPTS {
        let candidate = PackURI::new(template.replace("%d", &suffix.to_string()))?;
        if package.get_part(&candidate).is_err() {
            package.validate_new_part_name(&candidate)?;
            return Ok(candidate);
        }
    }
    Err(invalid("no free threaded-comments part name"))
}

fn choose_relationship_id(
    relationships: &litchi_opc::Relationships,
    requested: Option<&str>,
    prefix: &str,
) -> Result<String> {
    if let Some(requested) = requested {
        if relationships.get(requested).is_none() {
            return Ok(requested.to_owned());
        }
        return Err(invalid(format!(
            "relationship ID '{requested}' is already used"
        )));
    }
    for suffix in 1..=MAX_NAME_ATTEMPTS {
        let candidate = format!("{prefix}{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free threaded-comments relationship ID"))
}

fn remove_if_unreferenced(package: &mut OpcPackage, target: &PackURI) {
    let referenced = package.iter_parts().any(|part| {
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
    });
    if !referenced {
        package.remove_part(target);
    }
}

fn map_model_error(error: impl std::fmt::Display) -> Error {
    Error::Unrecognized {
        typ: "XLSB threaded-comments".to_owned(),
        val: error.to_string(),
    }
}

fn invalid(value: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: "XLSB threaded-comments package graph".to_owned(),
        val: value.into(),
    }
}
