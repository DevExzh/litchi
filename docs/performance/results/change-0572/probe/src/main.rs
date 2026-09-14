//! Throwaway probe for change 0572: OOXML source-backed range attribution.
//!
//! This binary drives three source-backed OOXML read scenarios against a
//! counting `litchi_core::ReadAt`, under three read-ahead policies, two
//! construction routes and three transports, and writes every ordered
//! `(offset, length, returned)` triple to a JSON capture. It changes nothing
//! under `crates/`; it only calls the public API.

mod counting;
mod harness;

use std::{path::PathBuf, sync::Arc, time::Instant};

use counting::CountingSource;
use harness::{Fixture, Outcome, Policy, Repeat, Route, TransportArm, arm_json, run_repeats};
use litchi_core::ReadAt;

const REPEATS: usize = 5;

/// A generous but bounded text policy; the scenario is "full text", not a cap test.
fn text_options() -> litchi_core::TextOutputOptions<'static> {
    litchi_core::TextOutputOptions::new("\n", "\n\n", 64 * 1024 * 1024, 1_000_000)
}

fn package_with_policy(
    source: Arc<dyn ReadAt>,
    policy: Policy,
) -> litchi_opc::SourceBackedPackage {
    litchi_opc::SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
        source,
        litchi_opc::ReadLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        policy.to_source_read_policy(),
    )
    .expect("source-backed package")
}

// --- scenario bodies -------------------------------------------------------

fn docx_text(source: Arc<CountingSource>, policy: Policy, route: Route) -> Outcome {
    let counter = Arc::clone(&source);
    let dynamic: Arc<dyn ReadAt> = source;
    let open_started = Instant::now();
    let package = match route {
        Route::NativeLeaf => litchi_docx::source_backed::Package::from_read_at(dynamic)
            .expect("docx native leaf open"),
        Route::PackageThenAdopt => litchi_docx::source_backed::Package::from_source_backed_package(
            package_with_policy(dynamic, policy),
        )
        .expect("docx adopt"),
    };
    let open_ns = open_started.elapsed().as_nanos();
    let open_requests = counter.calls();
    let mut out: Vec<u8> = Vec::new();
    let observation = match package.write_text_to(&mut out, text_options()) {
        Ok(report) => format!(
            "text_bytes={} objects={}",
            out.len(),
            report.objects_written()
        ),
        Err(error) => format!("REFUSED: {error:?}"),
    };
    Outcome { open_requests, open_ns, observation }
}

fn pptx_middle_slide_text(source: Arc<CountingSource>, policy: Policy, route: Route) -> Outcome {
    let counter = Arc::clone(&source);
    let dynamic: Arc<dyn ReadAt> = source;
    let open_started = Instant::now();
    let presentation = match route {
        Route::NativeLeaf => litchi_pptx::SourceBackedPresentation::from_read_at(dynamic)
            .expect("pptx native leaf open"),
        Route::PackageThenAdopt => {
            litchi_pptx::SourceBackedPresentation::from_source_backed_package(
                package_with_policy(dynamic, policy),
            )
            .expect("pptx adopt")
        }
    };
    let open_ns = open_started.elapsed().as_nanos();
    let open_requests = counter.calls();
    let count = presentation.slide_count();
    let middle = count / 2;
    let slide = presentation.slide(middle).expect("middle slide");
    let observation = match slide.text() {
        Ok(text) => format!("slides={count} middle={middle} text_bytes={}", text.len()),
        Err(error) => format!("slides={count} middle={middle} REFUSED: {error:?}"),
    };
    Outcome { open_requests, open_ns, observation }
}

fn xlsx_cell_a1(source: Arc<CountingSource>, policy: Policy, route: Route) -> Outcome {
    let counter = Arc::clone(&source);
    let dynamic: Arc<dyn ReadAt> = source;
    let open_started = Instant::now();
    let workbook = match route {
        Route::NativeLeaf => litchi_xlsx::SourceBackedWorkbook::from_read_at(dynamic)
            .expect("xlsx native leaf open"),
        Route::PackageThenAdopt => litchi_xlsx::SourceBackedWorkbook::from_source_backed_package(
            package_with_policy(dynamic, policy),
        )
        .expect("xlsx adopt"),
    };
    let open_ns = open_started.elapsed().as_nanos();
    let open_requests = counter.calls();
    let sheet = workbook.sheets().next().expect("sheet 0");
    let name = sheet.name().to_string();
    let observation = match sheet.cell("A1") {
        Ok(view) => format!(
            "sheet0={name:?} a1={}",
            format!("{view:?}").chars().take(72).collect::<String>()
        ),
        Err(error) => format!("sheet0={name:?} REFUSED: {error:?}"),
    };
    Outcome { open_requests, open_ns, observation }
}

// --- driver ----------------------------------------------------------------

struct Arm {
    scenario: &'static str,
    policy: Policy,
    route: Route,
}

fn arms() -> Vec<Arm> {
    let mut out = Vec::new();
    for scenario in ["docx_open_full_text", "pptx_open_middle_slide_text", "xlsx_open_cell_a1"] {
        out.push(Arm { scenario, policy: Policy::Exact, route: Route::NativeLeaf });
        out.push(Arm { scenario, policy: Policy::Exact, route: Route::PackageThenAdopt });
        out.push(Arm {
            scenario,
            policy: Policy::ForwardStart(4096),
            route: Route::PackageThenAdopt,
        });
        out.push(Arm {
            scenario,
            policy: Policy::ForwardStart(65536),
            route: Route::PackageThenAdopt,
        });
    }
    out
}

fn main() {
    let mut args = std::env::args().skip(1);
    let repo = PathBuf::from(args.next().expect("usage: probe <repo-root> <out.json>"));
    let out_path = PathBuf::from(args.next().expect("usage: probe <repo-root> <out.json>"));
    let revision = args.next().unwrap_or_else(|| "unspecified".to_string());

    let corpus = vec![
        Fixture::load(&repo, "docx-comment", "test-data/ooxml/docx/comment.docx", "docx", "small, every member carries a data descriptor"),
        Fixture::load(&repo, "docx-endnotes", "test-data/ooxml/docx/endnotes.docx", "docx", "16 members; the largest-member-count DOCX whose full text this scenario can actually extract"),
        Fixture::load(&repo, "docx-test-comment", "test-data/ooxml/docx/testComment.docx", "docx", "17 members, 65,298 B; the largest DOCX by bytes that this scenario completes"),
        Fixture::load(&repo, "pptx-shape-glow-effect", "test-data/ooxml/pptx/shape-glow-effect.pptx", "pptx", "smallest PPTX member count in the corpus"),
        Fixture::load(&repo, "pptx-shapes", "test-data/ooxml/pptx/shapes.pptx", "pptx", "the plan's six-slide fixture"),
        Fixture::load(&repo, "pptx-shape-soft-edges", "test-data/ooxml/pptx/shape-soft-edges.pptx", "pptx", "only descriptor-bearing PPTX in the corpus"),
        Fixture::load(&repo, "xlsx-sheet-names", "test-data/ooxml/xlsx/sheet-names.xlsx", "xlsx", "the plan's small fixture"),
        Fixture::load(&repo, "xlsx-universal-content", "test-data/ooxml/xlsx/universal-content.xlsx", "xlsx", "the plan's descriptor-bearing fixture"),
        Fixture::load(&repo, "xlsx-conditional-formatting", "test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx", "xlsx", "the plan's 132-member fixture"),
        Fixture::load(&repo, "xlsx-simple-normal", "test-data/ooxml/xlsx/SimpleNormal.xlsx", "xlsx", "added beyond the plan: 12 members, A1 is stored and the streaming cell read completes"),
        Fixture::load(&repo, "xlsx-pivot-sample", "test-data/ooxml/xlsx/ExcelPivotTableSample.xlsx", "xlsx", "added beyond the plan: 27 members, A1 is stored and the streaming cell read completes"),
    ];

    let transports = [
        TransportArm::ZeroDelay,
        TransportArm::ZeroDelayCapped,
        TransportArm::Delayed0493,
    ];

    let loadavg_start = std::fs::read_to_string("/proc/loadavg").unwrap_or_default();
    let mut runs = Vec::new();
    for fixture in &corpus {
        for arm in arms() {
            let applies = match (fixture.format, arm.scenario) {
                ("docx", "docx_open_full_text") => true,
                ("pptx", "pptx_open_middle_slide_text") => true,
                ("xlsx", "xlsx_open_cell_a1") => true,
                _ => false,
            };
            if !applies {
                continue;
            }
            for transport in transports {
                let policy = arm.policy;
                let route = arm.route;
                let scenario = arm.scenario;
                let repeats: Vec<Repeat> = run_repeats(fixture, transport, REPEATS, |source| {
                    match scenario {
                        "docx_open_full_text" => docx_text(source, policy, route),
                        "pptx_open_middle_slide_text" => {
                            pptx_middle_slide_text(source, policy, route)
                        }
                        "xlsx_open_cell_a1" => xlsx_cell_a1(source, policy, route),
                        other => unreachable!("unknown scenario {other}"),
                    }
                });
                eprintln!(
                    "{:<32} {:<28} {:<20} {:<26} requests={} median_ns={}",
                    fixture.id,
                    scenario,
                    policy.label(),
                    transport.label(),
                    repeats[0].requests.len(),
                    {
                        let mut v: Vec<u128> = repeats.iter().map(|r| r.elapsed_ns).collect();
                        v.sort_unstable();
                        v[v.len() / 2]
                    }
                );
                runs.push(arm_json(scenario, fixture, policy, route, transport, &repeats));
            }
        }
    }

    let loadavg = std::fs::read_to_string("/proc/loadavg").unwrap_or_default();
    let document = serde_json::json!({
        "schema_version": 1,
        "record_kind": "litchi-perf-0572-request-capture",
        "library_revision": revision,
        "library_source": "git archive of the revision above, extracted to a scratch tree, so a concurrent working-tree edit to crates/ cannot reach this build",
        "toolchain": "1.95.0",
        "profile": "release (the probe crate's own workspace: opt-level 3, no LTO, unwind panics)",
        "repeats": REPEATS,
        "loadavg_before_capture": loadavg_start.trim(),
        "loadavg_after_capture": loadavg.trim(),
        "corpus": corpus.iter().map(|f| serde_json::json!({
            "id": f.id,
            "path": f.path,
            "format": f.format,
            "role": f.role,
            "size": f.bytes.len(),
            "sha256": f.sha256,
        })).collect::<Vec<_>>(),
        "runs": runs,
    });
    std::fs::write(&out_path, serde_json::to_string(&document).expect("serialise"))
        .expect("write capture");
    eprintln!("wrote {}", out_path.display());
}
