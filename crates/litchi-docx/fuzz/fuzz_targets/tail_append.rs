#![no_main]

use std::io::Cursor;

use libfuzzer_sys::fuzz_target;
use litchi_docx::SourceLimits;
use litchi_docx::source_backed::{
    Package,
    tail_append::{Commit, EditOptions, Error, Limits, Plan},
};

fn publish<'package>(plan: Plan<'package>, output: &mut Vec<u8>) -> Result<Commit, Error> {
    let commit = plan.commit()?;
    commit.package().write_to(output).map_err(Error::Document)?;
    Ok(commit)
}

fn read_limits() -> SourceLimits {
    SourceLimits::builder()
        .max_input_bytes(256 * 1024)
        .unwrap()
        .max_members(32)
        .unwrap()
        .max_relationship_parts(32)
        .unwrap()
        .max_parts(32)
        .unwrap()
        .max_entry_bytes(128 * 1024)
        .unwrap()
        .max_aggregate_bytes(256 * 1024)
        .unwrap()
        .max_archive_compressed_bytes(256 * 1024)
        .unwrap()
        .max_archive_metadata_bytes(64 * 1024)
        .unwrap()
        .max_part_bytes(128 * 1024)
        .unwrap()
        .max_total_part_bytes(256 * 1024)
        .unwrap()
        .max_content_types_bytes(16 * 1024)
        .unwrap()
        .max_relationship_xml_bytes(16 * 1024)
        .unwrap()
        .max_total_relationship_xml_bytes(64 * 1024)
        .unwrap()
        .max_events(8192)
        .unwrap()
        .build()
        .unwrap()
}

fuzz_target!(|data: &[u8]| {
    if data.len() > 64 * 1024 {
        return;
    }
    let Ok(package) = Package::from_reader_with_options(
        Cursor::new(data),
        litchi_docx::OpenOptions::default().source_limits(read_limits()),
    ) else {
        return;
    };
    let limits = Limits {
        max_source_xml_bytes: 64 * 1024,
        max_text_bytes: 1024,
        max_fragment_bytes: 8192,
        max_candidate_xml_bytes: 128 * 1024,
        max_events: 8192,
        max_depth: 16,
        max_paragraphs: 1024,
        max_settings_xml_bytes: 8192,
        max_workspace_bytes: 16 * 1024 * 1024,
        max_output_bytes: 256 * 1024,
        max_token_bytes: 8192,
    };
    if let Ok(plan) = package
        .tail_append_noop()
        .with_options(EditOptions::new(limits))
        .prepare()
    {
        let mut output = Vec::new();
        if let Ok(publication) = publish(plan, &mut output) {
            assert!(publication.patch().is_noop());
            assert_eq!(output, data);
        }
    }
    let Ok(plan) = package
        .tail_append_plain_paragraph("fuzz <&> café")
        .with_options(EditOptions::new(limits))
        .prepare()
    else {
        return;
    };
    let before = plan.source_proof();
    let after = plan.candidate_proof();
    assert_eq!(after.paragraph_count, before.paragraph_count + 1);
    assert!(after.generated_once);
    assert_eq!(before.sect_pr_len, after.sect_pr_len);
    assert_eq!(before.sect_pr_sha256, after.sect_pr_sha256);
    let mut output = Vec::new();
    if let Ok(publication) = publish(plan, &mut output) {
        assert!(!publication.patch().is_noop());
        let current = Package::from_reader_with_options(
            Cursor::new(&output),
            litchi_docx::OpenOptions::default().source_limits(read_limits()),
        )
        .expect("a successfully published bounded candidate must reopen");
        let mut restored = Vec::new();
        current
            .apply_tail_append(&publication.patch().inverse())
            .expect("the exact published candidate must authorize its inverse")
            .package()
            .write_to(&mut restored)
            .expect("the exact inverse candidate must publish");
        assert_eq!(restored, data);
    }
});
