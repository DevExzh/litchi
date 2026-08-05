//! Typed threaded-comments graph with transactional XLSX source workbooks.
//!
//! The standalone workbook transaction does not yet attach threaded-comment
//! parts to a worksheet. This example retains the complete people/comment
//! graph and validates both XML parts through the XLSX codec instead.

use litchi_xlsx::Workbook;
use litchi_xlsx::threaded_comments::{
    Comment, Comments, Graph, Mention, People, PeoplePart, Person, validate_graph, write_comments,
    write_persons,
};
use std::env;
use std::fs;
use std::path::Path;

type ExampleResult<T> = Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn people() -> People {
    People {
        persons: vec![
            Person {
                display_name: "John Smith".to_string(),
                id: "{11111111-2222-3333-4444-555555555555}".to_string(),
                ..Person::default()
            },
            Person {
                display_name: "Jane Doe".to_string(),
                id: "{22222222-3333-4444-5555-666666666666}".to_string(),
                ..Person::default()
            },
            Person {
                display_name: "Alice Chen".to_string(),
                id: "{33333333-4444-5555-6666-777777777777}".to_string(),
                ..Person::default()
            },
        ],
    }
}

fn comment(id: &str, person_id: &str, cell_ref: &str, text: &str) -> Comment {
    Comment {
        cell_ref: Some(cell_ref.to_string()),
        id: id.to_string(),
        person_id: person_id.to_string(),
        text: Some(text.to_string()),
        date_time: Some("2024-01-16T09:00:00Z".to_string()),
        ..Comment::default()
    }
}

fn basic_graph() -> Graph {
    let persons = people();
    let john = persons.persons[0].id.clone();
    let jane = persons.persons[1].id.clone();
    Graph {
        persons: Some(PeoplePart {
            relationship_id: "rIdPeople".to_string(),
            part_name: "/xl/persons/person.xml".to_string(),
            persons,
        }),
        worksheets: vec![litchi_xlsx::threaded_comments::CommentsPart {
            worksheet_part_name: "/xl/worksheets/sheet1.xml".to_string(),
            relationship_id: "rIdComments".to_string(),
            part_name: "/xl/threadedComments/threadedComment1.xml".to_string(),
            comments: Comments {
                comments: vec![
                    comment(
                        "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}",
                        &john,
                        "B2",
                        "Please use OAuth 2.0.",
                    ),
                    comment(
                        "{B2C3D4E5-F6A7-8901-BCDE-F12345678901}",
                        &jane,
                        "B3",
                        "Use Material Design.",
                    ),
                ],
            },
        }],
    }
}

fn main() -> ExampleResult<()> {
    let output_dir = env::args()
        .nth(1)
        .unwrap_or_else(|| "threaded_comments_examples".to_string());
    fs::create_dir_all(&output_dir)?;

    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("default worksheet is missing")?
        .rename("TaskList")?;
    {
        let mut sheet = edit
            .sheet("TaskList")?
            .ok_or("TaskList worksheet is missing")?;
        for (cell, value) in [("A1", "Task ID"), ("B1", "Task Name"), ("C1", "Status")] {
            sheet.set(cell, value)?;
        }
        sheet
            .set("A2", "T-001")?
            .set("B2", "Implement authentication")?
            .set("C2", "In Progress")?;
        sheet
            .set("A3", "T-002")?
            .set("B3", "Design user interface")?
            .set("C3", "Not Started")?;
    }
    let workbook = edit.commit()?.into_workbook();
    let workbook_path = Path::new(&output_dir).join("comments_basic.xlsx");
    workbook.save(&workbook_path)?;

    let graph = basic_graph();
    validate_graph(&graph)?;
    let people_xml = write_persons(&graph.persons.as_ref().expect("people").persons)?;
    let comments_xml = write_comments(&graph.worksheets[0].comments)?;
    fs::write(
        Path::new(&output_dir).join("comments_basic.people.xml"),
        people_xml,
    )?;
    fs::write(
        Path::new(&output_dir).join("comments_basic.threaded-comments.xml"),
        comments_xml,
    )?;
    println!(
        "Generated {} and validated its typed threaded-comments graph",
        workbook_path.display()
    );
    Ok(())
}

#[allow(dead_code)]
fn mention_for(person_id: &str) -> Mention {
    Mention {
        mention_person_id: person_id.to_string(),
        mention_id: "{A1A2A3A4-A5A6-A7A8-A9A0-AABBCCDDEEFF}".to_string(),
        start_index: 0,
        length: 11,
    }
}
