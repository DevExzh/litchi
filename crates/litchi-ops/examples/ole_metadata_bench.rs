//! Paired native and unified legacy metadata writes (record 0254).
//!
//! Run with `cargo run --release -p litchi-ops --no-default-features --features
//! doc,xls,ppt --example ole_metadata_bench`. The native leg is the leaf
//! `Package::edit_metadata` + `SummaryProperties::set` + `commit`; the unified
//! leg is `File::set("/metadata", {"title": ..})`, which also republishes the
//! unified handle. Fixture open and byte-equivalence checks are outside
//! timing; samples alternate ABBA. This probe is not frozen-harness
//! qualification.

#[cfg(all(feature = "doc", feature = "xls", feature = "ppt"))]
mod enabled {
    use std::hint::black_box;
    use std::time::Instant;

    fn measure<T>(repeats: usize, mut operation: impl FnMut() -> T) -> u128 {
        let start = Instant::now();
        for _ in 0..repeats {
            drop(black_box(operation()));
        }
        start.elapsed().as_nanos()
    }

    fn median(mut values: Vec<u128>) -> u128 {
        values.sort_unstable();
        values[values.len() / 2]
    }

    fn report(
        name: &str,
        bytes: usize,
        repeats: usize,
        native: impl FnMut() -> Vec<u8>,
        unified: impl FnMut() -> Vec<u8>,
    ) {
        let (mut native, mut unified) = (native, unified);
        assert_eq!(native(), unified(), "{name}: native and unified bytes");
        let (mut native_ns, mut unified_ns) = (Vec::new(), Vec::new());
        for sample in 0..21 {
            if sample % 2 == 0 {
                native_ns.push(measure(repeats, &mut native));
                unified_ns.push(measure(repeats, &mut unified));
            } else {
                unified_ns.push(measure(repeats, &mut unified));
                native_ns.push(measure(repeats, &mut native));
            }
        }
        let native = median(native_ns) as f64 / repeats as f64;
        let unified = median(unified_ns) as f64 / repeats as f64;
        println!(
            "{{\"case\":\"{name}\",\"bytes\":{bytes},\"repeats\":{repeats},\"native_us\":{:.1},\"unified_us\":{:.1},\"overhead_pct\":{:.2}}}",
            native / 1_000.0,
            unified / 1_000.0,
            (unified / native - 1.0) * 100.0
        );
    }

    fn fixture(relative: &str) -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ole")
                .join(relative),
        )
        .expect("fixture")
    }

    pub(super) fn run() {
        let props = || litchi_ops::props! { "title" => "Retitled" };
        let path = litchi_ops::path::Path::parse("/metadata").expect("path");

        let bytes = fixture("doc/FancyFoot.doc");
        let package = litchi_doc::Package::from_bytes(&bytes).expect("doc");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified doc");
        report(
            "doc",
            bytes.len(),
            50,
            || {
                let mut edit = package.edit_metadata().expect("edit");
                edit.properties_mut()
                    .expect("properties")
                    .set("title", "Retitled")
                    .expect("set");
                edit.commit()
                    .expect("commit")
                    .into_package()
                    .to_bytes()
                    .expect("bytes")
            },
            || {
                file.set(&path, props())
                    .expect("set")
                    .to_bytes()
                    .expect("bytes")
            },
        );

        let bytes = fixture("xls/Simple.xls");
        let package = litchi_xls::Package::from_bytes(&bytes).expect("xls");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified xls");
        report(
            "xls",
            bytes.len(),
            50,
            || {
                let mut edit = package.edit_metadata().expect("edit");
                edit.properties_mut()
                    .expect("properties")
                    .set("title", "Retitled")
                    .expect("set");
                edit.commit()
                    .expect("commit")
                    .into_package()
                    .to_bytes()
                    .expect("bytes")
            },
            || {
                file.set(&path, props())
                    .expect("set")
                    .to_bytes()
                    .expect("bytes")
            },
        );

        let bytes = fixture("ppt/SampleShow.ppt");
        let package = litchi_ppt::Package::from_bytes(&bytes).expect("ppt");
        let file = litchi_ops::open_bytes(bytes.clone()).expect("unified ppt");
        report(
            "ppt",
            bytes.len(),
            50,
            || {
                let mut edit = package.edit_metadata().expect("edit");
                edit.properties_mut()
                    .expect("properties")
                    .set("title", "Retitled")
                    .expect("set");
                edit.commit()
                    .expect("commit")
                    .into_package()
                    .to_bytes()
                    .expect("bytes")
            },
            || {
                file.set(&path, props())
                    .expect("set")
                    .to_bytes()
                    .expect("bytes")
            },
        );
    }
}

fn main() {
    #[cfg(all(feature = "doc", feature = "xls", feature = "ppt"))]
    enabled::run();
}
