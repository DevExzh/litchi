//! Deterministic query compilation and bounded traversal smoke.
use litchi_ops::Code;
use litchi_ops::query::__fake::{Document, Family, compile_for, run};
use litchi_ops::query::Query;
fn main() {
    let documents = [
        (Family::Word, Document::word(12)),
        (Family::Presentation, Document::deck(2, 3)),
        (Family::Workbook, Document::workbook(6)),
    ];
    let seeds = include_str!("../../../litchi-core/tests/fixtures/officecli_selectors.txt");
    let mut inputs = 0;
    let mut compiled = 0;
    for seed in seeds
        .lines()
        .filter(|line| !line.is_empty() && !line.starts_with('#'))
    {
        let mut mutations = vec![seed.to_owned()];
        for suffix in ["[", ":has(", "[SENSITIVE!=x]", ",p", ":not(:empty)"] {
            mutations.push(format!("{seed}{suffix}"));
        }
        for input in mutations {
            inputs += 1;
            let Ok(query) = Query::parse(&input) else {
                continue;
            };
            for (family, document) in &documents {
                match compile_for(*family, &query) {
                    Ok(plan) => {
                        compiled += 1;
                        match run(document, &plan, 16) {
                            Ok(result) => assert!(result.matches <= 16),
                            Err(error) => assert_ne!(error.code(), Code::InternalError),
                        }
                    },
                    Err(error) => assert_ne!(error.code(), Code::InternalError),
                }
            }
        }
    }
    println!("{inputs} inputs; {compiled} compiled family plans");
}
