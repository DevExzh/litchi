//! Publish one bounded source-backed ODP append and independently reopen it.

use std::{fs::OpenOptions, sync::Arc};

use litchi_core::FileSource;
use litchi_odf_common::SourceContentPublicationOptions;
use litchi_odp::SourceBackedPresentation;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() != 4 {
        return Err("expected source-path output-path title body".into());
    }
    let source = SourceBackedPresentation::from_read_at(Arc::new(FileSource::open(&args[0])?))?;
    let options = SourceContentPublicationOptions::default();
    let mut edit = source.edit_tail_append(args[2].clone(), args[3].clone())?;
    edit.set_options(options.clone());
    let plan = edit.prepare()?;
    let proof = plan.proof().clone();
    let (decoded_offset, source_content_length, target_content_length) = (
        plan.decoded_offset(),
        plan.source_content_length(),
        plan.target_content_length(),
    );
    let commit = plan.commit()?;
    let output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&args[1])?;
    let report = commit.package().write_to_with_options(output, options)?;
    let reopened = SourceBackedPresentation::open(&args[1])?;
    let expected_count = proof.slide_count() + 1;
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
            "source_slides": proof.slide_count(), "target_slides": expected_count,
            "name": proof.page_name(), "insert_at": decoded_offset,
            "source_content_bytes": source_content_length,
            "target_content_bytes": target_content_length,
            "semantic_reopen": true,
        })
    );
    Ok(())
}
