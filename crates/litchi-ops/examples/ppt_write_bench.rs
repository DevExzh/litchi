//! Paired native and unified immutable PPT shape-text writes.
//!
//! Run with `cargo run --release -p litchi-ops --no-default-features --features ppt
//! --example ppt_write_bench`. Fixture/open/path preparation and byte-equivalence
//! checks are outside timing. This probe is not frozen-harness qualification.

#[cfg(feature = "ppt")]
mod enabled {
    use litchi_core::Position;
    use litchi_ppt::text_edit::Target;
    use serde::Serialize;
    use std::{hint::black_box, time::Instant};

    #[derive(Serialize)]
    struct Case {
        shapes: usize,
        source: &'static str,
        repeats: usize,
        native_ns: Vec<u128>,
        unified_ns: Vec<u128>,
    }

    fn measure<T>(repeats: usize, mut operation: impl FnMut() -> T) -> u128 {
        let start = Instant::now();
        for _ in 0..repeats {
            drop(black_box(operation()));
        }
        start.elapsed().as_nanos()
    }

    pub(super) fn run() {
        let mut cases = Vec::new();
        for shapes in [1, 100, 1_000] {
            let mut builder = litchi_ppt::Builder::new();
            let slide = builder.add_slide().expect("slide");
            for _ in 0..shapes {
                builder
                    .add_textbox(slide, 10, 10, 240, 40, "before")
                    .expect("shape");
            }
            let bytes = builder.build().expect("package").to_bytes().expect("bytes");
            let temporary = tempfile::NamedTempFile::new().expect("file");
            std::fs::write(temporary.path(), &bytes).expect("fixture");
            for positional in [false, true] {
                let package = if positional {
                    litchi_ppt::Package::open(temporary.path()).expect("native source")
                } else {
                    litchi_ppt::Package::from_bytes(&bytes).expect("native bytes")
                };
                let file = if positional {
                    litchi_ops::open(temporary.path()).expect("unified source")
                } else {
                    litchi_ops::open_bytes(bytes.clone()).expect("unified bytes")
                };
                let path = litchi_ops::path::Path::parse("/slide[1]/shape[1]").expect("path");
                let target = Target::new(Position::new(0), Position::new(0));
                let native = || {
                    let mut edit = package.edit().expect("native edit");
                    edit.set_shape_text(target, "after").expect("native set");
                    edit.commit().expect("native commit").into_package()
                };
                let unified = || {
                    file.set(&path, litchi_ops::props! { "text" => "after" })
                        .expect("unified set")
                };
                assert_eq!(
                    native().to_bytes().expect("native bytes"),
                    unified().to_bytes().expect("unified bytes")
                );
                let mut case = Case {
                    shapes,
                    source: if positional { "positional" } else { "memory" },
                    repeats: 10,
                    native_ns: Vec::new(),
                    unified_ns: Vec::new(),
                };
                for sample in 0..7 {
                    if sample % 2 == 0 {
                        case.native_ns.push(measure(case.repeats, native));
                        case.unified_ns.push(measure(case.repeats, unified));
                    } else {
                        case.unified_ns.push(measure(case.repeats, unified));
                        case.native_ns.push(measure(case.repeats, native));
                    }
                }
                cases.push(case);
            }
        }
        serde_json::to_writer_pretty(std::io::stdout().lock(), &cases).expect("JSON");
        println!();
    }
}

fn main() {
    #[cfg(feature = "ppt")]
    enabled::run();
    #[cfg(not(feature = "ppt"))]
    eprintln!("enable the ppt feature");
}
