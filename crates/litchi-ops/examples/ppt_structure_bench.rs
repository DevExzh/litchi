//! Paired native and unified immutable PPT slide-structure timings.
//!
//! Each case authors the same slide corpus and compares the native
//! `Package::edit` structural operation with the corresponding unified
//! `File` operation.  Fixture generation, opening, path and placement
//! construction, warmup, serialization, and byte-equivalence checks are
//! outside the timed regions.  Seven timing pairs alternate native→unified
//! and unified→native order; each leg repeats the operation ten times.
//!
//! ```text
//! cargo run --release -p litchi-ops --no-default-features --features ppt \
//!   --example ppt_structure_bench
//! ```
//!
//! This exploratory probe does not replace the frozen performance harness.

#[cfg(feature = "ppt")]
mod enabled {
    use litchi_core::Position;
    use litchi_ops::{
        File,
        path::{Path, Position as PathPosition},
    };
    use litchi_ppt::Package;
    use serde::Serialize;
    use std::{hint::black_box, time::Instant};

    const SLIDE_COUNTS: &[usize] = &[3, 10, 100];
    const PAIRS: usize = 7;
    const REPEATS: usize = 10;

    #[derive(Clone, Copy, Debug)]
    enum Kind {
        Remove,
        Move,
        Swap,
    }

    impl Kind {
        const ALL: [Self; 3] = [Self::Remove, Self::Move, Self::Swap];

        const fn name(self) -> &'static str {
            match self {
                Self::Remove => "remove",
                Self::Move => "move",
                Self::Swap => "swap",
            }
        }
    }

    #[derive(Debug, Serialize)]
    struct Case {
        kind: &'static str,
        slides: usize,
        source: &'static str,
        pairs: usize,
        repeats: usize,
        native_ns: Vec<u128>,
        unified_ns: Vec<u128>,
    }

    struct Targets {
        remove: Path<'static>,
        low: Path<'static>,
        high: Path<'static>,
    }

    fn measure<T>(repeats: usize, mut operation: impl FnMut() -> T) -> u128 {
        let start = Instant::now();
        for _ in 0..repeats {
            drop(black_box(operation()));
        }
        start.elapsed().as_nanos()
    }

    fn targets(slides: usize) -> Targets {
        let middle = slides / 2 + 1;
        Targets {
            remove: Path::parse(&format!("/slide[{middle}]"))
                .expect("remove path")
                .into_owned(),
            low: Path::parse("/slide[1]").expect("first path"),
            high: Path::parse(&format!("/slide[{slides}]"))
                .expect("last path")
                .into_owned(),
        }
    }

    fn native_operation(package: &Package, kind: Kind, slides: usize) -> Package {
        let mut edit = package.edit().expect("native edit");
        match kind {
            Kind::Remove => edit
                .remove_slide(Position::new(slides / 2))
                .expect("native remove"),
            Kind::Move => edit
                .move_slide(Position::new(0), Position::new(slides - 1))
                .expect("native move"),
            Kind::Swap if slides == 2 => edit
                .move_slide(Position::new(0), Position::new(1))
                .expect("native adjacent swap"),
            Kind::Swap => {
                // The native editor has move but no swap primitive.  Moving
                // low to high first shifts the old high slide to high - 1;
                // moving that slide back to low exchanges the endpoints.
                edit.move_slide(Position::new(0), Position::new(slides - 1))
                    .expect("native swap first move");
                edit.move_slide(Position::new(slides - 2), Position::new(0))
                    .expect("native swap second move");
            },
        }
        edit.commit().expect("native commit").into_package()
    }

    fn unified_operation(file: &File, kind: Kind, targets: &Targets, slides: usize) -> File {
        match kind {
            Kind::Remove => file.remove(&targets.remove).expect("unified remove"),
            Kind::Move => file
                .move_to(&targets.low, PathPosition::new(slides - 1))
                .expect("unified move"),
            Kind::Swap => file
                .swap(&targets.low, &targets.high)
                .expect("unified swap"),
        }
    }

    fn fixture(slides: usize) -> Vec<u8> {
        let mut builder = litchi_ppt::Builder::new();
        for index in 0..slides {
            let slide = builder.add_slide().expect("slide");
            let text = format!("slide-{index}");
            builder
                .add_textbox(slide, 10, 10, 240, 40, &text)
                .expect("shape");
        }
        builder
            .build()
            .expect("package")
            .to_bytes()
            .expect("fixture bytes")
    }

    fn assert_source_unchanged(
        package: &Package,
        file: &File,
        original: &[u8],
        kind: Kind,
        slides: usize,
    ) {
        assert_eq!(
            package.to_bytes().expect("native source bytes"),
            original,
            "native source changed for {} with {slides} slides",
            kind.name()
        );
        assert_eq!(
            file.to_bytes().expect("unified source bytes"),
            original,
            "unified source changed for {} with {slides} slides",
            kind.name()
        );
    }

    pub(super) fn run() {
        let mut cases = Vec::with_capacity(SLIDE_COUNTS.len() * Kind::ALL.len() * 2);
        for &slides in SLIDE_COUNTS {
            let bytes = fixture(slides);
            let temporary = tempfile::NamedTempFile::new().expect("file");
            std::fs::write(temporary.path(), &bytes).expect("write fixture");
            let targets = targets(slides);

            for positional in [false, true] {
                let package = if positional {
                    Package::open(temporary.path()).expect("native positional source")
                } else {
                    Package::from_bytes(&bytes).expect("native eager source")
                };
                let file = if positional {
                    litchi_ops::open(temporary.path()).expect("unified positional source")
                } else {
                    litchi_ops::open_bytes(bytes.clone()).expect("unified eager source")
                };
                let source = if positional { "positional" } else { "eager" };

                for kind in Kind::ALL {
                    let native_warmup = native_operation(&package, kind, slides);
                    let unified_warmup = unified_operation(&file, kind, &targets, slides);
                    let native_bytes = native_warmup.to_bytes().expect("native warmup bytes");
                    let unified_bytes = unified_warmup.to_bytes().expect("unified warmup bytes");
                    assert_eq!(
                        native_bytes,
                        unified_bytes,
                        "native and unified {} outputs differ for {slides} slides, source={source}",
                        kind.name()
                    );
                    assert_source_unchanged(&package, &file, &bytes, kind, slides);

                    let mut case = Case {
                        kind: kind.name(),
                        slides,
                        source,
                        pairs: PAIRS,
                        repeats: REPEATS,
                        native_ns: Vec::with_capacity(PAIRS),
                        unified_ns: Vec::with_capacity(PAIRS),
                    };
                    for pair in 0..PAIRS {
                        if pair % 2 == 0 {
                            case.native_ns.push(measure(REPEATS, || {
                                native_operation(&package, kind, slides)
                            }));
                            case.unified_ns.push(measure(REPEATS, || {
                                unified_operation(&file, kind, &targets, slides)
                            }));
                        } else {
                            case.unified_ns.push(measure(REPEATS, || {
                                unified_operation(&file, kind, &targets, slides)
                            }));
                            case.native_ns.push(measure(REPEATS, || {
                                native_operation(&package, kind, slides)
                            }));
                        }
                    }
                    assert_source_unchanged(&package, &file, &bytes, kind, slides);
                    cases.push(case);
                }
            }
        }
        serde_json::to_writer_pretty(std::io::stdout().lock(), &cases).expect("benchmark JSON");
        println!();
    }
}

fn main() {
    #[cfg(feature = "ppt")]
    enabled::run();
    #[cfg(not(feature = "ppt"))]
    eprintln!("enable the ppt feature");
}
