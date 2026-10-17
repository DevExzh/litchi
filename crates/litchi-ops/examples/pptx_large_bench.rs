//! Large-slide PPTX edit-and-write timing (record 0317): one slide of 1,000
//! text boxes, shape 501. A source-backed PPTX owner keeps an edited slide as
//! an overlay and audits it when it writes the package, against the physical
//! slide (records 0311, 0313 and 0317), so each leg edits and then writes the
//! deck to a sink:
//!
//! - `leaf_set_write` / `unified_set_write`: `Edit::set_shape_text` + commit
//!   on a native `litchi_pptx::source::Package`, against the unified `set` of
//!   `/slide[1]/shape[501]` on the default owner;
//! - `leaf_add_write` / `unified_add_write`: `Edit::add_text_box`, against the
//!   unified `add` of a plain `shape` to `/slide[1]`;
//! - `committed_set_write` / `committed_add_write`: the unified set and add
//!   on a handle already committed from the default owner (shape 2 changed),
//!   whose overlay differs from the physical slide in two places;
//! - `leaf_write` / `unified_write`: writing the unedited owners.
//!
//! Both legs of each pair are checked to write the same bytes before timing.
//! Each leg: two warm-ups, nine samples of `PPTX_LARGE_LOOPS` calls (default
//! 20), the median in µs per call. `PPTX_LARGE_ONLY=<leg>` runs one leg;
//! `PPTX_LARGE_PROFILE=1` loops it once for a profiler.
//!
//! `taskset -c 30 cargo run --release -p litchi-ops --features pptx --example pptx_large_bench`

#[cfg(feature = "pptx")]
mod bench {
    use litchi_core::OwnedSource;
    use litchi_ops::{File, NewNode, props};
    use std::hint::black_box;
    use std::sync::Arc;
    use std::time::Instant;

    const SHAPES: usize = 1_000;
    const BOUNDS: (i64, i64, i64, i64) = (360_000, 360_000, 3_600_000, 720_000);

    fn deck(shapes: usize) -> Vec<u8> {
        let mut builder = litchi_pptx::builder::Builder::new().expect("builder");
        let slide = builder
            .presentation()
            .expect("presentation")
            .add_slide()
            .expect("slide");
        for index in 0..shapes {
            slide.add_text_box(&format!("box {index}"), 0, 0, 914_400, 457_200);
        }
        builder.build().expect("build").to_bytes().expect("bytes")
    }

    fn time(n: usize, f: &mut dyn FnMut()) -> f64 {
        for _ in 0..2 {
            f();
        }
        let mut v: Vec<f64> = (0..9)
            .map(|_| {
                let s = Instant::now();
                for _ in 0..n {
                    f();
                }
                s.elapsed().as_secs_f64() * 1e6 / n as f64
            })
            .collect();
        v.sort_by(f64::total_cmp);
        v[4]
    }

    fn written(file: &File) -> Vec<u8> {
        let mut out = Vec::new();
        file.write_to(&mut out).expect("unified write");
        out
    }

    fn native_written(package: &litchi_pptx::source::Package) -> Vec<u8> {
        let mut out = Vec::new();
        package.write_to(&mut out).expect("native write");
        out
    }

    pub(super) fn main() {
        let only = std::env::var("PPTX_LARGE_ONLY").ok();
        let loops: usize = std::env::var("PPTX_LARGE_LOOPS")
            .ok()
            .map_or(20, |v| v.parse().unwrap());
        let bytes = deck(SHAPES);
        let native =
            litchi_pptx::source::Package::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
                .expect("native");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("open");
        let target = "/slide[1]/shape[501]";
        let text = props! {"text" => "Replaced text"};
        let node = NewNode::new("shape")
            .with_props(props! {
                "text" => "Hello", "x" => BOUNDS.0, "y" => BOUNDS.1,
                "width" => BOUNDS.2, "height" => BOUNDS.3,
            })
            .expect("node");
        let committed = file
            .set("/slide[1]/shape[2]", props! {"text" => "Committed earlier"})
            .expect("committed");

        let leaf_set = || {
            let mut edit = native.edit().expect("edit");
            edit.set_shape_text(0_usize, 500_usize, "Replaced text")
                .expect("set");
            edit.commit().expect("commit").into_package()
        };
        let leaf_add = || {
            let mut edit = native.edit().expect("edit");
            edit.add_text_box(0_usize, "Hello", BOUNDS).expect("add");
            edit.commit().expect("commit").into_package()
        };
        assert_eq!(
            native_written(&leaf_set()),
            written(&file.set(target, text.clone()).expect("set")),
            "set legs differ"
        );
        assert_eq!(
            native_written(&leaf_add()),
            written(&file.add("/slide[1]", node.clone()).expect("add").file),
            "add legs differ"
        );
        assert_eq!(native_written(&native), written(&file), "write legs differ");

        let mut legs: Vec<(&str, Box<dyn FnMut()>)> = vec![
            (
                "leaf_set_write",
                Box::new(|| {
                    let package = leaf_set();
                    package.write_to(std::io::sink()).expect("write");
                    black_box(package);
                }),
            ),
            (
                "unified_set_write",
                Box::new(|| {
                    let edited = file.set(target, text.clone()).expect("set");
                    edited.write_to(std::io::sink()).expect("write");
                    black_box(edited);
                }),
            ),
            (
                "leaf_add_write",
                Box::new(|| {
                    let package = leaf_add();
                    package.write_to(std::io::sink()).expect("write");
                    black_box(package);
                }),
            ),
            (
                "unified_add_write",
                Box::new(|| {
                    let added = file.add("/slide[1]", node.clone()).expect("add");
                    added.file.write_to(std::io::sink()).expect("write");
                    black_box(added);
                }),
            ),
            (
                "committed_set_write",
                Box::new(|| {
                    let edited = committed.set(target, text.clone()).expect("set");
                    edited.write_to(std::io::sink()).expect("write");
                    black_box(edited);
                }),
            ),
            (
                "committed_add_write",
                Box::new(|| {
                    let added = committed.add("/slide[1]", node.clone()).expect("add");
                    added.file.write_to(std::io::sink()).expect("write");
                    black_box(added);
                }),
            ),
            (
                "leaf_write",
                Box::new(|| {
                    native.write_to(std::io::sink()).expect("write");
                }),
            ),
            (
                "unified_write",
                Box::new(|| {
                    file.write_to(std::io::sink()).expect("write");
                }),
            ),
        ];
        let profile = std::env::var_os("PPTX_LARGE_PROFILE").is_some();
        for (name, f) in &mut legs {
            if only.as_deref().is_some_and(|only| only != *name) {
                continue;
            }
            if profile {
                for _ in 0..loops {
                    f();
                }
                continue;
            }
            println!("{name} {:.1} us", time(loops, f.as_mut()));
        }
    }
}

fn main() {
    #[cfg(feature = "pptx")]
    bench::main();
    #[cfg(not(feature = "pptx"))]
    eprintln!("pptx_large_bench needs the pptx feature");
}
