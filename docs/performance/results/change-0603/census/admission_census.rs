// Scratch census for change 0603, appended temporarily to
// crates/litchi-xlsx/src/cell_values/shared_traversal_tests.rs, run with
// `cargo test -p litchi-xlsx --lib marker_admission_census -- --nocapture`,
// and removed again. It counts, for every worksheet part of every real .xlsx
// fixture, which traversal the part takes before and after this change.
#[test]
fn marker_admission_census() {
    use crate::raw::worksheet::{
        SourceAdmission, SourceParseAttempt, parse_source_with_observer, source_stream_admission,
    };
    use litchi_opc::OpcPackage;

    let roots = [
        concat!(env!("CARGO_MANIFEST_DIR"), "/../../test-data/ooxml/xlsx"),
        concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/office-interop/libreoffice-resaved"
        ),
        concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/office-interop/litchi-changed"
        ),
    ];
    let mut paths: Vec<std::path::PathBuf> = Vec::new();
    for root in roots {
        let Ok(entries) = std::fs::read_dir(root) else {
            continue;
        };
        paths.extend(
            entries
                .filter_map(|entry| entry.ok().map(|entry| entry.path()))
                .filter(|path| path.extension().is_some_and(|value| value == "xlsx")),
        );
    }
    paths.sort();

    let (mut parts, mut files) = (0usize, 0usize);
    let (mut borrowed, mut rewritten, mut refused) = (0usize, 0usize, 0usize);
    let (mut fused_before, mut fused_after) = (0usize, 0usize);
    let mut rewritten_fused: Vec<String> = Vec::new();
    for path in &paths {
        let Ok(package) = OpcPackage::open(path) else {
            continue;
        };
        files += 1;
        for part in package.iter_parts() {
            let name = part.partname().as_str().to_owned();
            if !name.starts_with("/xl/worksheets/sheet") || !name.ends_with(".xml") {
                continue;
            }
            let content = part.blob();
            parts += 1;
            let Some(admission) = source_stream_admission(content) else {
                refused += 1;
                continue;
            };
            let mut validator = Validator::new(XmlOwner::Worksheet);
            let complete = matches!(
                parse_source_with_observer(content, admission, || Ok(None), |namespace, event| {
                    validator.observe(namespace, event)
                }),
                SourceParseAttempt::Complete(_)
            ) && validator.finish().is_ok();
            match admission {
                SourceAdmission::Borrowed => {
                    borrowed += 1;
                    fused_before += usize::from(complete);
                    fused_after += usize::from(complete);
                },
                SourceAdmission::Rewritten => {
                    rewritten += 1;
                    fused_after += usize::from(complete);
                    if complete {
                        rewritten_fused.push(format!("{}{name}", path.display()));
                    }
                },
            }
        }
    }
    println!("files\t{files}");
    println!("worksheet_parts\t{parts}");
    println!("admission_refused\t{refused}");
    println!("admission_borrowed\t{borrowed}");
    println!("admission_rewritten\t{rewritten}");
    println!("fused_completed_before\t{fused_before}");
    println!("fused_completed_after\t{fused_after}");
    for name in &rewritten_fused {
        println!("newly_fused\t{name}");
    }
}
