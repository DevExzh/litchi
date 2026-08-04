//! Threaded comments showcase for XLSX writer.
//!
//! This example creates workbooks with threaded comments (modern Excel comments)
//! that support conversation threads, @mentions, and resolution status.
//!
//! ```bash
//! cargo run --example xlsx_threaded_comments_test -- threaded_comments_examples
//! # -> threaded_comments_examples/comments_basic.xlsx
//! # -> threaded_comments_examples/comments_with_threads.xlsx
//! ```

use litchi::ooxml::xlsx::{
    Workbook,
    threaded_comments::{Comment, Mention, People, Person},
};
use std::env;
use std::fs;
use std::path::Path;

type ExampleResult<T> = Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn main() -> ExampleResult<()> {
    let output_dir = env::args()
        .nth(1)
        .unwrap_or_else(|| "threaded_comments_examples".to_string());
    fs::create_dir_all(&output_dir)?;

    let basic_path = Path::new(&output_dir).join("comments_basic.xlsx");
    let threaded_path = Path::new(&output_dir).join("comments_with_threads.xlsx");

    generate_basic_example(&basic_path)?;
    generate_threaded_example(&threaded_path)?;

    println!("Generated threaded comments workbooks:");
    println!("  - {}", basic_path.display());
    println!("  - {}", threaded_path.display());
    println!("Open them in Microsoft Excel to verify the threaded comments.");
    Ok(())
}

/// Generate a basic example with simple threaded comments.
fn generate_basic_example(path: &Path) -> ExampleResult<()> {
    let mut wb = Workbook::create()?;

    // Create person list for comment authors
    let mut person_list = People::default();
    person_list.persons.push(Person {
        display_name: "John Smith".to_string(),
        id: "{11111111-2222-3333-4444-555555555555}".to_string(),
        user_id: None,
        provider_id: None,
    });
    person_list.persons.push(Person {
        display_name: "Jane Doe".to_string(),
        id: "{22222222-3333-4444-5555-666666666666}".to_string(),
        user_id: None,
        provider_id: None,
    });
    wb.set_person_list(person_list);

    let sheet = wb.worksheet_mut(0)?;
    sheet.set_name("TaskList".to_string());

    // Add some data
    sheet.set_cell_value(1, 1, "Task ID");
    sheet.set_cell_value(1, 2, "Task Name");
    sheet.set_cell_value(1, 3, "Status");
    sheet.set_cell_value(1, 4, "Priority");

    sheet.set_cell_value(2, 1, "T-001");
    sheet.set_cell_value(2, 2, "Implement authentication");
    sheet.set_cell_value(2, 3, "In Progress");
    sheet.set_cell_value(2, 4, "High");

    sheet.set_cell_value(3, 1, "T-002");
    sheet.set_cell_value(3, 2, "Design user interface");
    sheet.set_cell_value(3, 3, "Not Started");
    sheet.set_cell_value(3, 4, "Medium");

    sheet.set_cell_value(4, 1, "T-003");
    sheet.set_cell_value(4, 2, "Write documentation");
    sheet.set_cell_value(4, 3, "Completed");
    sheet.set_cell_value(4, 4, "Low");

    // Add threaded comments to specific cells
    let comment1 = Comment {
        cell_ref: Some("B2".to_string()),
        id: "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}".to_string(),
        person_id: "{11111111-2222-3333-4444-555555555555}".to_string(),
        text: Some("Please use OAuth 2.0 for this task.".to_string()),
        date_time: Some("2024-01-15T10:30:00Z".to_string()),
        parent_id: None,
        done: None,
        mentions: Vec::new(),
    };

    let comment2 = Comment {
        cell_ref: Some("B3".to_string()),
        id: "{B2C3D4E5-F6A7-8901-BCDE-F12345678901}".to_string(),
        person_id: "{22222222-3333-4444-5555-666666666666}".to_string(),
        text: Some("We should follow Material Design guidelines.".to_string()),
        date_time: Some("2024-01-15T11:15:00Z".to_string()),
        parent_id: None,
        done: None,
        mentions: Vec::new(),
    };

    let comment3 = Comment {
        cell_ref: Some("B4".to_string()),
        id: "{C3D4E5F6-A7B8-9012-CDEF-123456789012}".to_string(),
        person_id: "{11111111-2222-3333-4444-555555555555}".to_string(),
        text: Some("Great job on the documentation!".to_string()),
        date_time: Some("2024-01-15T14:45:00Z".to_string()),
        parent_id: None,
        done: Some(true),
        mentions: Vec::new(),
    };

    sheet.add_threaded_comment(comment1);
    sheet.add_threaded_comment(comment2);
    sheet.add_threaded_comment(comment3);

    wb.save(path)?;
    Ok(())
}

/// Generate an example with conversation threads and mentions.
fn generate_threaded_example(path: &Path) -> ExampleResult<()> {
    let mut wb = Workbook::create()?;

    // Create person list for comment authors
    let mut person_list = People::default();
    person_list.persons.push(Person {
        display_name: "John Smith".to_string(),
        id: "{11111111-2222-3333-4444-555555555555}".to_string(),
        user_id: None,
        provider_id: None,
    });
    person_list.persons.push(Person {
        display_name: "Jane Doe".to_string(),
        id: "{22222222-3333-4444-5555-666666666666}".to_string(),
        user_id: None,
        provider_id: None,
    });
    person_list.persons.push(Person {
        display_name: "Alice Chen".to_string(),
        id: "{33333333-4444-5555-6666-777777777777}".to_string(),
        user_id: None,
        provider_id: None,
    });
    wb.set_person_list(person_list);

    let sheet = wb.worksheet_mut(0)?;
    sheet.set_name("TeamDiscussion".to_string());

    // Add data
    sheet.set_cell_value(1, 1, "Feature");
    sheet.set_cell_value(1, 2, "Owner");
    sheet.set_cell_value(1, 3, "Due Date");

    sheet.set_cell_value(2, 1, "Search functionality");
    sheet.set_cell_value(2, 2, "Alice Chen");
    sheet.set_cell_value(2, 3, "2024-02-01");

    sheet.set_cell_value(3, 1, "Export to PDF");
    sheet.set_cell_value(3, 2, "Bob Smith");
    sheet.set_cell_value(3, 3, "2024-02-15");

    // Create a conversation thread with replies
    let main_comment = Comment {
        cell_ref: Some("A2".to_string()),
        id: "{D4E5F6A7-A8B9-0123-DEFA-234567890123}".to_string(),
        person_id: "{33333333-4444-5555-6666-777777777777}".to_string(),
        text: Some("Should we use Elasticsearch or build our own search?".to_string()),
        date_time: Some("2024-01-16T09:00:00Z".to_string()),
        parent_id: None,
        done: None,
        mentions: Vec::new(),
    };

    // Reply to the main comment
    let reply1 = Comment {
        cell_ref: Some("A2".to_string()),
        id: "{E5F6A7B8-A9B0-1234-EFAB-345678901234}".to_string(),
        person_id: "{11111111-2222-3333-4444-555555555555}".to_string(),
        text: Some("I recommend Elasticsearch for scalability.".to_string()),
        date_time: Some("2024-01-16T09:15:00Z".to_string()),
        parent_id: Some("{D4E5F6A7-A8B9-0123-DEFA-234567890123}".to_string()),
        done: None,
        mentions: Vec::new(),
    };

    // Another reply with @mention
    let reply2 = Comment {
        cell_ref: Some("A2".to_string()),
        id: "{F6A7B8C9-A0B1-2345-FABC-456789012345}".to_string(),
        person_id: "{22222222-3333-4444-5555-666666666666}".to_string(),
        text: Some("@Alice Chen can you evaluate both options?".to_string()),
        date_time: Some("2024-01-16T10:30:00Z".to_string()),
        parent_id: Some("{D4E5F6A7-A8B9-0123-DEFA-234567890123}".to_string()),
        done: None,
        mentions: vec![Mention {
            mention_person_id: "{33333333-4444-5555-6666-777777777777}".to_string(),
            mention_id: "{A1A2A3A4-A5A6-A7A8-A9A0-AABBCCDDEEFF}".to_string(),
            start_index: 0,
            length: 11,
        }],
    };

    // Resolved comment on another cell
    let resolved_comment = Comment {
        cell_ref: Some("A3".to_string()),
        id: "{A7B8C9D0-A1B2-3456-ABCD-567890123456}".to_string(),
        person_id: "{11111111-2222-3333-4444-555555555555}".to_string(),
        text: Some("We need to use a proper PDF library for this.".to_string()),
        date_time: Some("2024-01-16T11:00:00Z".to_string()),
        parent_id: None,
        done: Some(true),
        mentions: Vec::new(),
    };

    sheet.add_threaded_comment(main_comment);
    sheet.add_threaded_comment(reply1);
    sheet.add_threaded_comment(reply2);
    sheet.add_threaded_comment(resolved_comment);

    wb.save(path)?;
    Ok(())
}
