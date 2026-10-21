//! Paired native Graph probe for record 0359c; the driver builds both revisions.
use litchi_ograph::{Limits, chart::{Context, Model, Refs}};
use std::{hint::black_box, io::Cursor, time::Instant};

fn main() {
    let args: Vec<_> = std::env::args().collect();
    let route = &args[1];
    let iterations: usize = args[2].parse().unwrap();
    let bytes = std::fs::read(&args[3]).unwrap();
    let mut cfb = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let workbook = cfb.open_stream(&["Workbook"]).unwrap();
    let refs: Vec<_> = Refs::open_workbook(&workbook, Limits::default()).unwrap()
        .collect::<Result<_, _>>().unwrap();
    assert!(!refs.is_empty());
    // Prove exact untouched output and forward/inverse output before timing.
    for raw in &refs {
        assert_eq!(Model::parse(*raw, Context::excel()).unwrap().encode().unwrap().as_bytes(), raw.as_bytes());
    }
    let base = || Model::parse(black_box(refs[0]), Context::excel()).unwrap();
    let mut rect = base().rect();
    rect.width = rect.width.checked_add(1).unwrap();
    let mut edit = base().edit().unwrap();
    edit.set_rect(rect).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();
    let after = commit.into_package();
    let after_bytes = patch.apply(base()).unwrap().into_package().encode().unwrap().into_bytes();
    assert_ne!(after_bytes, refs[0].as_bytes());
    assert_eq!(patch.apply(base()).unwrap().into_package().encode().unwrap().as_bytes(), after_bytes);
    assert_eq!(patch.inverse().apply(after).unwrap().into_package().encode().unwrap().as_bytes(), refs[0].as_bytes());
    let run = || -> usize {
        match route.as_str() {
            "validate" => black_box(Limits::default()).validate().unwrap().max_unknown_bytes,
            "scan" => Refs::open_workbook(black_box(&workbook), black_box(Limits::default())).unwrap()
                .map(|raw| raw.unwrap().as_bytes().len()).sum(),
            "parse_encode" => refs.iter().map(|raw| {
                Model::parse(black_box(*raw), Context::excel()).unwrap().encode().unwrap().as_bytes().len()
            }).sum(),
            "parse_edit_commit" => {
                let mut edit = base().edit().unwrap();
                edit.set_rect(black_box(rect)).unwrap();
                edit.commit().unwrap().into_package().rect().width as usize
            },
            "parse_apply" => black_box(&patch).apply(base()).unwrap().into_package().rect().width as usize,
            _ => panic!("unknown route"),
        }
    };
    let expected = run();
    for _ in 0..200 { black_box(run()); }
    let start = Instant::now();
    let mut checksum = 0usize;
    for _ in 0..iterations { checksum = checksum.wrapping_add(black_box(run())); }
    let ns = start.elapsed().as_nanos();
    assert_eq!(checksum, expected.wrapping_mul(iterations));
    println!("{{\"route\":\"{}\",\"iterations\":{},\"ns\":{},\"checksum\":{},\"charts\":{},\"error_bytes\":{},\"result_bytes\":{}}}",
        route, iterations, ns, checksum, refs.len(), std::mem::size_of::<litchi_ograph::Error>(), std::mem::size_of::<litchi_ograph::Result<()>>());
}
