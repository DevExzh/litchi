//! Native Word/Excel MS-OVBA data: budgets and inert reference edits.

use litchi_cfb::OleFile;
use litchi_vba::{LimitKind, Limits, dir::Dir, project::Project};
use std::{io::Cursor, path::PathBuf};

fn native(path: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(path),
    )
    .unwrap()
}

#[test]
fn native_vba_aggregate_source_limit_is_exact_across_recorded_modules() {
    for (file, root, total) in [
        ("word/word-recorded-macro.doc", "Macros", 733),
        ("excel/excel-recorded-macro.xls", "_VBA_PROJECT_CUR", 428),
    ] {
        let bytes = native(file);
        let mut ole =
            OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default()).unwrap();
        let limits = Limits {
            max_total_source_bytes: total,
            ..Limits::default()
        };
        let project = Project::open(&mut ole, &[root], &limits).unwrap();
        assert_eq!(
            project
                .modules()
                .iter()
                .map(|m| m.source().raw().len())
                .sum::<usize>(),
            total
        );
        let limits = Limits {
            max_total_source_bytes: total - 1,
            ..limits
        };
        assert!(
            matches!(Project::open(&mut ole, &[root], &limits), Err(litchi_vba::Error::Limit {
            kind: LimitKind::TotalSourceBytes, observed, maximum,
        }) if observed > maximum && maximum == total - 1)
        );
    }
}

#[test]
fn native_vba_reference_name_edits_preserve_module_metadata_and_other_references() {
    for (file, root, count) in [
        ("word/word-recorded-macro.doc", "Macros", 3),
        ("excel/excel-recorded-macro.xls", "_VBA_PROJECT_CUR", 2),
    ] {
        let bytes = native(file);
        let mut ole =
            OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default()).unwrap();
        let compressed = ole.open_stream(&[root, "VBA", "dir"]).unwrap();
        let limits = Limits::default();
        let project = Project::open(&mut ole, &[root], &limits).unwrap();
        let original = Dir::parse(&compressed, &limits).unwrap();
        assert_eq!(project.references().len(), count);
        assert_eq!(
            project
                .encode_dir_with_references(project.references(), &limits)
                .unwrap(),
            compressed
        );
        let mut references = project.references().to_vec();
        let selected = references
            .iter()
            .position(|r| r.name() == Some("Office"))
            .unwrap();
        references[selected] = references[selected]
            .clone()
            .edit_name_source_bound(Some("OfficeInspection"), project.page(), &limits)
            .unwrap();
        let encoded = project
            .encode_dir_with_references(&references, &limits)
            .unwrap();
        let parsed = Dir::parse(&encoded, &limits).unwrap();
        assert_eq!(parsed.project_name(), original.project_name());
        assert_eq!(parsed.page_id(), 10000);
        assert_eq!(parsed.modules(), original.modules());
        assert_eq!(parsed.references().len(), count);
        for (index, reference) in parsed.references().iter().enumerate() {
            assert_eq!(reference.kind(), original.references()[index].kind());
            if index == selected {
                assert_eq!(reference.name(), Some("OfficeInspection"));
            } else {
                assert_eq!(reference, &original.references()[index]);
            }
        }
        // The original immutable project remains an exact source for a no-op.
        assert_eq!(
            project
                .encode_dir_with_references(project.references(), &limits)
                .unwrap(),
            compressed
        );
    }
}
