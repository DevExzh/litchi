//! Publish one bounded source-backed ODP append and independently reopen it.

use std::{fs::OpenOptions, sync::Arc};

use litchi_core::FileSource;
use litchi_odf_common::SourceContentPublicationOptions;
use litchi_odp::{SourceBackedPresentation, SourceBackedTailAppendEdit};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() != 4 {
        return Err("expected source-path output-path title body".into());
    }
    let edit = SourceBackedTailAppendEdit::from_read_at(
        Arc::new(FileSource::open(&args[0])?),
        args[2].clone(),
        args[3].clone(),
    )?;
    let options = SourceContentPublicationOptions::default();
    let plan = edit.plan(&options)?;
    let output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&args[1])?;
    let report = plan.write_to(output, options)?;
    let reopened = SourceBackedPresentation::from_path(&args[1])?;
    let expected_count = plan.proof().slide_count() + 1;
    if reopened.slide_count()? != expected_count {
        return Err("published slide count differs from source plus one".into());
    }
    let last = reopened
        .slide(expected_count - 1)?
        .ok_or("missing appended slide")?;
    if last.title()? != Some(args[2].as_str()) || last.text()? != args[3] {
        return Err("reopened appended title/body differ".into());
    }
    println!(
        "{}",
        serde_json::json!({
            "status": "published", "bytes": report.bytes(),
            "source_slides": plan.proof().slide_count(), "target_slides": expected_count,
            "name": plan.proof().page_name(), "insert_at": plan.decoded_offset(),
            "source_content_bytes": plan.source_content_length(),
            "target_content_bytes": plan.target_content_length(),
            "semantic_reopen": true,
        })
    );
    Ok(())
}
