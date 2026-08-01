//! Package-level CRUD for SpreadsheetML threaded comments and persons.

use std::collections::{HashMap, HashSet};

use litchi_core::sheet::Result as SheetResult;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

use super::{
    Person, PersonList, ThreadedComment, ThreadedComments, read_persons, read_threaded_comments,
    validate_threaded_timestamp, write_persons, write_threaded_comments,
};

/// The workbook-level persons part and its relationship identity.
#[derive(Debug, Clone)]
pub struct WorkbookPersonPart {
    pub relationship_id: String,
    pub part_name: String,
    pub persons: PersonList,
}

/// One worksheet's threaded-comments part and relationship identity.
#[derive(Debug, Clone)]
pub struct WorksheetThreadedCommentPart {
    pub worksheet_part_name: String,
    pub relationship_id: String,
    pub part_name: String,
    pub comments: ThreadedComments,
}

/// Complete typed threaded-comment graph for one workbook.
#[derive(Debug, Clone, Default)]
pub struct ThreadedCommentGraph {
    pub persons: Option<WorkbookPersonPart>,
    pub worksheets: Vec<WorksheetThreadedCommentPart>,
}

/// Load and cross-validate all persons and worksheet comment threads.
pub fn load_threaded_comment_graph(package: &OpcPackage) -> SheetResult<ThreadedCommentGraph> {
    let workbook = package.main_document_part()?;
    let persons_relationship = one_internal_relationship(workbook, rt::PERSONS, "persons")?;
    let persons = match (persons_relationship, read_persons(package)?) {
        (Some((relationship_id, part_name)), Some(persons)) => Some(WorkbookPersonPart {
            relationship_id,
            part_name: part_name.to_string(),
            persons,
        }),
        (None, None) => None,
        _ => return Err("inconsistent persons relationship graph".into()),
    };

    let worksheet_names: Vec<PackURI> = package
        .iter_parts()
        .filter(|part| {
            part.rels()
                .iter()
                .any(|relationship| relationship.reltype() == rt::THREADED_COMMENTS)
        })
        .map(|part| part.partname().clone())
        .collect();
    let mut worksheets = Vec::with_capacity(worksheet_names.len());
    for worksheet_name in worksheet_names {
        let worksheet = package.get_part(&worksheet_name)?;
        let Some((relationship_id, part_name)) =
            one_internal_relationship(worksheet, rt::THREADED_COMMENTS, "threaded comments")?
        else {
            continue;
        };
        let comments = read_threaded_comments(package, &worksheet_name)?
            .ok_or("inconsistent threaded-comments relationship graph")?;
        worksheets.push(WorksheetThreadedCommentPart {
            worksheet_part_name: worksheet_name.to_string(),
            relationship_id,
            part_name: part_name.to_string(),
            comments,
        });
    }
    worksheets.sort_by(|left, right| left.worksheet_part_name.cmp(&right.worksheet_part_name));
    let graph = ThreadedCommentGraph {
        persons,
        worksheets,
    };
    validate_threaded_comment_graph(&graph)?;
    Ok(graph)
}

pub fn find_threaded_comment_person(
    package: &OpcPackage,
    person_id: &str,
) -> SheetResult<Option<Person>> {
    Ok(load_threaded_comment_graph(package)?
        .persons
        .and_then(|part| {
            part.persons
                .persons
                .into_iter()
                .find(|person| person.id == person_id)
        }))
}

pub fn add_threaded_comment_person(
    package: &mut OpcPackage,
    person: Person,
) -> SheetResult<WorkbookPersonPart> {
    let mut graph = load_threaded_comment_graph(package)?;
    if graph.persons.as_ref().is_some_and(|part| {
        part.persons
            .persons
            .iter()
            .any(|candidate| candidate.id == person.id)
    }) {
        return Err(format!("duplicate threaded-comment person ID '{}'", person.id).into());
    }
    if let Some(mut part) = graph.persons.take() {
        part.persons.persons.push(person);
        graph.persons = Some(part.clone());
        validate_threaded_comment_graph(&graph)?;
        commit_person_part(package, &part)?;
        return Ok(part);
    }

    let workbook_name = package.main_document_part()?.partname().clone();
    let part_name = next_person_part_name(package)?;
    let relationship_id = next_relationship_id(package.get_part(&workbook_name)?, "rIdPersons")?;
    let part = WorkbookPersonPart {
        relationship_id: relationship_id.clone(),
        part_name: part_name.to_string(),
        persons: PersonList {
            persons: vec![person],
        },
    };
    graph.persons = Some(part.clone());
    validate_threaded_comment_graph(&graph)?;
    let xml = write_persons(&part.persons)?.into_bytes();
    package.try_add_part(Box::new(BlobPart::new(
        part_name.clone(),
        ct::SML_PERSONS.into(),
        xml,
    )))?;
    package
        .get_part_mut(&workbook_name)?
        .rels_mut()
        .add_relationship(
            rt::PERSONS.into(),
            part_name.relative_ref(workbook_name.base_uri()),
            relationship_id,
            false,
        );
    package.unsign();
    Ok(part)
}

pub fn update_threaded_comment_person<F>(
    package: &mut OpcPackage,
    person_id: &str,
    update: F,
) -> SheetResult<bool>
where
    F: FnOnce(&mut Person),
{
    let mut graph = load_threaded_comment_graph(package)?;
    let Some(mut part) = graph.persons.clone() else {
        return Ok(false);
    };
    let Some(person) = part
        .persons
        .persons
        .iter_mut()
        .find(|person| person.id == person_id)
    else {
        return Ok(false);
    };
    update(person);
    if person.id != person_id {
        return Err("threaded-comment person update cannot change its ID".into());
    }
    graph.persons = Some(part.clone());
    validate_threaded_comment_graph(&graph)?;
    commit_person_part(package, &part)?;
    Ok(true)
}

pub fn replace_threaded_comment_person(
    package: &mut OpcPackage,
    person_id: &str,
    replacement: Person,
) -> SheetResult<bool> {
    if replacement.id != person_id {
        return Err("replacement threaded-comment person ID must match".into());
    }
    update_threaded_comment_person(package, person_id, move |person| *person = replacement)
}

pub fn remove_threaded_comment_person(
    package: &mut OpcPackage,
    person_id: &str,
) -> SheetResult<bool> {
    let mut graph = load_threaded_comment_graph(package)?;
    let Some(mut part) = graph.persons.clone() else {
        return Ok(false);
    };
    let Some(index) = part
        .persons
        .persons
        .iter()
        .position(|person| person.id == person_id)
    else {
        return Ok(false);
    };
    if graph.worksheets.iter().any(|sheet| {
        sheet.comments.comments.iter().any(|comment| {
            comment.person_id == person_id
                || comment
                    .mentions
                    .iter()
                    .any(|mention| mention.mention_person_id == person_id)
        })
    }) {
        return Err(format!("threaded-comment person {person_id} is still referenced").into());
    }
    part.persons.persons.remove(index);
    if !part.persons.persons.is_empty() {
        graph.persons = Some(part.clone());
        validate_threaded_comment_graph(&graph)?;
        commit_person_part(package, &part)?;
        return Ok(true);
    }

    graph.persons = None;
    validate_threaded_comment_graph(&graph)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let part_name = PackURI::new(&part.part_name)?;
    package
        .get_part_mut(&workbook_name)?
        .rels_mut()
        .remove(&part.relationship_id);
    if !part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    package.unsign();
    Ok(true)
}

pub fn reorder_threaded_comment_persons(
    package: &mut OpcPackage,
    ordered_person_ids: &[String],
) -> SheetResult<Vec<Person>> {
    let mut graph = load_threaded_comment_graph(package)?;
    let Some(mut part) = graph.persons.clone() else {
        if ordered_person_ids.is_empty() {
            return Ok(Vec::new());
        }
        return Err("threaded-comment persons part is missing".into());
    };
    let ordered = reorder_by_id(
        std::mem::take(&mut part.persons.persons),
        ordered_person_ids,
        |person| &person.id,
        "person",
    )?;
    part.persons.persons = ordered.clone();
    graph.persons = Some(part.clone());
    validate_threaded_comment_graph(&graph)?;
    commit_person_part(package, &part)?;
    Ok(ordered)
}

pub fn find_threaded_comment(
    package: &OpcPackage,
    worksheet_part_name: &PackURI,
    cell_ref: &str,
    comment_id: &str,
) -> SheetResult<Option<ThreadedComment>> {
    let graph = load_threaded_comment_graph(package)?;
    let Some(sheet) = graph
        .worksheets
        .iter()
        .find(|sheet| sheet.worksheet_part_name == worksheet_part_name.to_string())
    else {
        return Ok(None);
    };
    Ok(sheet
        .comments
        .comments
        .iter()
        .find(|comment| {
            comment.id == comment_id
                && resolved_cell_ref(&sheet.comments, comment)
                    .is_some_and(|resolved| resolved == cell_ref)
        })
        .cloned())
}

pub fn add_threaded_comment(
    package: &mut OpcPackage,
    worksheet_part_name: &PackURI,
    comment: ThreadedComment,
) -> SheetResult<WorksheetThreadedCommentPart> {
    let mut graph = load_threaded_comment_graph(package)?;
    if graph.worksheets.iter().any(|sheet| {
        sheet
            .comments
            .comments
            .iter()
            .any(|candidate| candidate.id == comment.id)
    }) {
        return Err(format!("duplicate threaded-comment ID '{}'", comment.id).into());
    }
    if let Some(index) = graph
        .worksheets
        .iter()
        .position(|sheet| sheet.worksheet_part_name == worksheet_part_name.to_string())
    {
        graph.worksheets[index].comments.comments.push(comment);
        validate_threaded_comment_graph(&graph)?;
        let part = graph.worksheets[index].clone();
        commit_comment_part(package, &part)?;
        return Ok(part);
    }

    package.get_part(worksheet_part_name)?;
    let part_name = next_threaded_comment_part_name(package)?;
    let relationship_id = next_relationship_id(
        package.get_part(worksheet_part_name)?,
        "rIdThreadedComments",
    )?;
    let part = WorksheetThreadedCommentPart {
        worksheet_part_name: worksheet_part_name.to_string(),
        relationship_id: relationship_id.clone(),
        part_name: part_name.to_string(),
        comments: ThreadedComments {
            comments: vec![comment],
        },
    };
    graph.worksheets.push(part.clone());
    validate_threaded_comment_graph(&graph)?;
    let xml = write_threaded_comments(&part.comments)?.into_bytes();
    package.try_add_part(Box::new(BlobPart::new(
        part_name.clone(),
        ct::SML_THREADED_COMMENTS.into(),
        xml,
    )))?;
    package
        .get_part_mut(worksheet_part_name)?
        .rels_mut()
        .add_relationship(
            rt::THREADED_COMMENTS.into(),
            part_name.relative_ref(worksheet_part_name.base_uri()),
            relationship_id,
            false,
        );
    package.unsign();
    Ok(part)
}

pub fn add_threaded_comment_reply(
    package: &mut OpcPackage,
    worksheet_part_name: &PackURI,
    cell_ref: &str,
    parent_id: &str,
    mut reply: ThreadedComment,
) -> SheetResult<WorksheetThreadedCommentPart> {
    let parent = find_threaded_comment(package, worksheet_part_name, cell_ref, parent_id)?
        .ok_or("threaded-comment parent was not found at the requested cell")?;
    if parent.parent_id.is_some() {
        return Err("threaded-comment replies must target a thread root".into());
    }
    reply.parent_id = Some(parent_id.into());
    reply.cell_ref = None;
    add_threaded_comment(package, worksheet_part_name, reply)
}

pub fn update_threaded_comment<F>(
    package: &mut OpcPackage,
    worksheet_part_name: &PackURI,
    cell_ref: &str,
    comment_id: &str,
    update: F,
) -> SheetResult<bool>
where
    F: FnOnce(&mut ThreadedComment),
{
    let mut graph = load_threaded_comment_graph(package)?;
    let Some(sheet_index) = graph
        .worksheets
        .iter()
        .position(|sheet| sheet.worksheet_part_name == worksheet_part_name.to_string())
    else {
        return Ok(false);
    };
    let Some(comment_index) = graph.worksheets[sheet_index]
        .comments
        .comments
        .iter()
        .position(|comment| {
            comment.id == comment_id
                && resolved_cell_ref(&graph.worksheets[sheet_index].comments, comment)
                    .is_some_and(|resolved| resolved == cell_ref)
        })
    else {
        return Ok(false);
    };
    update(&mut graph.worksheets[sheet_index].comments.comments[comment_index]);
    if graph.worksheets[sheet_index].comments.comments[comment_index].id != comment_id {
        return Err("threaded-comment update cannot change its ID".into());
    }
    validate_threaded_comment_graph(&graph)?;
    let part = graph.worksheets[sheet_index].clone();
    commit_comment_part(package, &part)?;
    Ok(true)
}

pub fn replace_threaded_comment(
    package: &mut OpcPackage,
    worksheet_part_name: &PackURI,
    cell_ref: &str,
    comment_id: &str,
    replacement: ThreadedComment,
) -> SheetResult<bool> {
    if replacement.id != comment_id {
        return Err("replacement threaded-comment ID must match".into());
    }
    update_threaded_comment(
        package,
        worksheet_part_name,
        cell_ref,
        comment_id,
        move |comment| *comment = replacement,
    )
}

pub fn remove_threaded_comment(
    package: &mut OpcPackage,
    worksheet_part_name: &PackURI,
    cell_ref: &str,
    comment_id: &str,
) -> SheetResult<bool> {
    let mut graph = load_threaded_comment_graph(package)?;
    let Some(sheet_index) = graph
        .worksheets
        .iter()
        .position(|sheet| sheet.worksheet_part_name == worksheet_part_name.to_string())
    else {
        return Ok(false);
    };
    let Some(comment_index) = graph.worksheets[sheet_index]
        .comments
        .comments
        .iter()
        .position(|comment| {
            comment.id == comment_id
                && resolved_cell_ref(&graph.worksheets[sheet_index].comments, comment)
                    .is_some_and(|resolved| resolved == cell_ref)
        })
    else {
        return Ok(false);
    };
    if graph.worksheets[sheet_index]
        .comments
        .comments
        .iter()
        .any(|comment| comment.parent_id.as_deref() == Some(comment_id))
    {
        return Err("cannot remove a threaded-comment root while replies remain".into());
    }
    graph.worksheets[sheet_index]
        .comments
        .comments
        .remove(comment_index);
    if !graph.worksheets[sheet_index].comments.comments.is_empty() {
        validate_threaded_comment_graph(&graph)?;
        let part = graph.worksheets[sheet_index].clone();
        commit_comment_part(package, &part)?;
        return Ok(true);
    }

    let part = graph.worksheets.remove(sheet_index);
    validate_threaded_comment_graph(&graph)?;
    let part_name = PackURI::new(&part.part_name)?;
    package
        .get_part_mut(worksheet_part_name)?
        .rels_mut()
        .remove(&part.relationship_id);
    if !part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    package.unsign();
    Ok(true)
}

pub fn reorder_threaded_comments(
    package: &mut OpcPackage,
    worksheet_part_name: &PackURI,
    ordered_comment_ids: &[String],
) -> SheetResult<Vec<ThreadedComment>> {
    let mut graph = load_threaded_comment_graph(package)?;
    let Some(sheet_index) = graph
        .worksheets
        .iter()
        .position(|sheet| sheet.worksheet_part_name == worksheet_part_name.to_string())
    else {
        if ordered_comment_ids.is_empty() {
            return Ok(Vec::new());
        }
        return Err("worksheet has no threaded-comments part".into());
    };
    let ordered = reorder_by_id(
        std::mem::take(&mut graph.worksheets[sheet_index].comments.comments),
        ordered_comment_ids,
        |comment| &comment.id,
        "threaded comment",
    )?;
    graph.worksheets[sheet_index].comments.comments = ordered.clone();
    validate_threaded_comment_graph(&graph)?;
    let part = graph.worksheets[sheet_index].clone();
    commit_comment_part(package, &part)?;
    Ok(ordered)
}

pub fn validate_threaded_comment_graph(graph: &ThreadedCommentGraph) -> SheetResult<()> {
    let empty_persons = PersonList::default();
    let persons = graph
        .persons
        .as_ref()
        .map(|part| &part.persons)
        .unwrap_or(&empty_persons);
    write_persons(persons)?;
    let person_ids: HashSet<&str> = persons
        .persons
        .iter()
        .map(|person| person.id.as_str())
        .collect();
    let mut comment_ids = HashSet::new();
    let mut mention_ids = HashSet::new();
    for sheet in &graph.worksheets {
        write_threaded_comments(&sheet.comments)?;
        let mut root_cells = HashSet::new();
        for comment in &sheet.comments.comments {
            if !comment_ids.insert(comment.id.as_str()) {
                return Err(
                    format!("duplicate workbook threaded-comment ID '{}'", comment.id).into(),
                );
            }
            if !person_ids.contains(comment.person_id.as_str()) {
                return Err(format!(
                    "threaded comment '{}' references missing person '{}'",
                    comment.id, comment.person_id
                )
                .into());
            }
            validate_threaded_timestamp(comment.date_time.as_deref())?;
            if comment.parent_id.is_none() {
                let cell = comment.cell_ref.as_deref().ok_or_else(|| {
                    format!(
                        "threaded-comment root '{}' is missing its cell reference",
                        comment.id
                    )
                })?;
                if !root_cells.insert(cell) {
                    return Err(
                        format!("worksheet has multiple threaded-comment roots at {cell}").into(),
                    );
                }
            } else if comment.cell_ref.is_some() {
                return Err(format!(
                    "threaded-comment reply '{}' must not carry a cell reference",
                    comment.id
                )
                .into());
            }
            for mention in &comment.mentions {
                if !person_ids.contains(mention.mention_person_id.as_str()) {
                    return Err(format!(
                        "mention '{}' references missing person '{}'",
                        mention.mention_id, mention.mention_person_id
                    )
                    .into());
                }
                if !mention_ids.insert(mention.mention_id.as_str()) {
                    return Err(
                        format!("duplicate workbook mention ID '{}'", mention.mention_id).into(),
                    );
                }
            }
        }
        let roots: HashSet<&str> = sheet
            .comments
            .comments
            .iter()
            .filter(|comment| comment.parent_id.is_none())
            .map(|comment| comment.id.as_str())
            .collect();
        for reply in sheet
            .comments
            .comments
            .iter()
            .filter(|comment| comment.parent_id.is_some())
        {
            if !roots.contains(reply.parent_id.as_deref().expect("filtered")) {
                return Err(format!(
                    "threaded-comment reply '{}' must reference a root in the same worksheet",
                    reply.id
                )
                .into());
            }
        }
    }
    Ok(())
}

fn commit_person_part(package: &mut OpcPackage, part: &WorkbookPersonPart) -> SheetResult<()> {
    let xml = write_persons(&part.persons)?.into_bytes();
    let part_name = PackURI::new(&part.part_name)?;
    package.get_part_mut(&part_name)?.set_blob(xml);
    package.unsign();
    Ok(())
}

fn commit_comment_part(
    package: &mut OpcPackage,
    part: &WorksheetThreadedCommentPart,
) -> SheetResult<()> {
    let xml = write_threaded_comments(&part.comments)?.into_bytes();
    let part_name = PackURI::new(&part.part_name)?;
    package.get_part_mut(&part_name)?.set_blob(xml);
    package.unsign();
    Ok(())
}

fn one_internal_relationship(
    owner: &dyn Part,
    relationship_type: &str,
    description: &str,
) -> SheetResult<Option<(String, PackURI)>> {
    let mut found = owner
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type);
    let Some(relationship) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(format!("part has multiple {description} relationships").into());
    }
    if relationship.is_external() {
        return Err(format!("{description} relationship cannot be external").into());
    }
    Ok(Some((
        relationship.r_id().into(),
        relationship.target_partname()?,
    )))
}

fn resolved_cell_ref<'a>(
    comments: &'a ThreadedComments,
    comment: &'a ThreadedComment,
) -> Option<&'a str> {
    if let Some(cell_ref) = comment.cell_ref.as_deref() {
        return Some(cell_ref);
    }
    let parent_id = comment.parent_id.as_deref()?;
    comments
        .comments
        .iter()
        .find(|candidate| candidate.id == parent_id)
        .and_then(|parent| parent.cell_ref.as_deref())
}

fn reorder_by_id<T, F>(
    values: Vec<T>,
    ordered_ids: &[String],
    id: F,
    description: &str,
) -> SheetResult<Vec<T>>
where
    F: Fn(&T) -> &str,
{
    if values.len() != ordered_ids.len() {
        return Err(format!("{description} reorder must contain every ID").into());
    }
    let mut remaining: HashMap<String, T> = HashMap::with_capacity(values.len());
    for value in values {
        let value_id = id(&value).to_string();
        if remaining.insert(value_id, value).is_some() {
            return Err(format!("duplicate {description} ID").into());
        }
    }
    let mut ordered = Vec::with_capacity(ordered_ids.len());
    for requested in ordered_ids {
        ordered.push(
            remaining
                .remove(requested)
                .ok_or_else(|| format!("unknown or duplicate {description} ID '{requested}'"))?,
        );
    }
    if !remaining.is_empty() {
        return Err(format!("{description} reorder must contain every ID").into());
    }
    Ok(ordered)
}

fn next_person_part_name(package: &OpcPackage) -> SheetResult<PackURI> {
    let canonical = PackURI::new("/xl/persons/person.xml")?;
    if package.get_part(&canonical).is_err() {
        return Ok(canonical);
    }
    for suffix in 1..=65_537u32 {
        let candidate = PackURI::new(format!("/xl/persons/person{suffix}.xml"))?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err("no free persons part name".into())
}

fn next_threaded_comment_part_name(package: &OpcPackage) -> SheetResult<PackURI> {
    for suffix in 1..=65_537u32 {
        let candidate = PackURI::new(format!("/xl/threadedComments/threadedComment{suffix}.xml"))?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err("no free threaded-comments part name".into())
}

fn next_relationship_id(owner: &dyn Part, prefix: &str) -> SheetResult<String> {
    for suffix in 1..=65_537u32 {
        let candidate = format!("{prefix}{suffix}");
        if owner.rels().get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(format!("no free {prefix} relationship ID").into())
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part_name| part_name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|part_name| part_name == *target)
    })
}
