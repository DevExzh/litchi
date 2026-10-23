//! Record 0358a: the source-backed ODF publishers, timed old route against
//! new route (binary-level A-B-B-A).
//!
//! This file drives the record 0358a family routes; its twin
//! `docs/unified-ops/results/0358a/probe_before.rs` drives the former stream
//! routes on the base commit. Everything outside `mod route` is identical in
//! both, so the fixtures, sinks, clocks and allocation counters match:
//!
//! - `ods/cells1`, `ods/cells100`: open a `SourceBackedSpreadsheet`, change
//!   1 or 100 existing cells, commit and write;
//! - `ods/dde`: replace a worksheet DDE source;
//! - `ods/sheet_metadata`: add a consolidation;
//! - `odp/tail_append`: append one slide to a source-backed ODP.
//!
//! Each row reports the median and minimum lifecycle time, the peak heap
//! growth of the lifecycle (a counting allocator; the source bytes are
//! allocated before), and the output's length and 64-bit SipHash (fixed
//! keys), which must match between the routes. With a second argument the
//! outputs are also written to that directory.
//!
//! `taskset -c 29 cargo run --release -p litchi-ops --features ods,odp --example odf_source_backed_probe -- [rounds] [dump-dir]`.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

struct Counting;

static CURRENT: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

// SAFETY: every call forwards to `System` unchanged; the counters only
// observe sizes.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: forwarded with the caller's layout.
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            let now = CURRENT.fetch_add(layout.size(), Ordering::Relaxed) + layout.size();
            PEAK.fetch_max(now, Ordering::Relaxed);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: forwarded with the caller's pointer and layout.
        unsafe { System.dealloc(pointer, layout) };
        CURRENT.fetch_sub(layout.size(), Ordering::Relaxed);
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        // SAFETY: forwarded with the caller's arguments.
        let moved = unsafe { System.realloc(pointer, layout, size) };
        if !moved.is_null() {
            if size >= layout.size() {
                let now = CURRENT.fetch_add(size - layout.size(), Ordering::Relaxed) + size
                    - layout.size();
                PEAK.fetch_max(now, Ordering::Relaxed);
            } else {
                CURRENT.fetch_sub(layout.size() - size, Ordering::Relaxed);
            }
        }
        moved
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

#[cfg(all(feature = "ods", feature = "odp"))]
mod probe {
    use std::hash::Hasher;
    use std::io::Write;
    use std::sync::Arc;
    use std::sync::atomic::Ordering;
    use std::time::Instant;

    use litchi_core::{OwnedSource, ReadAt};
    use litchi_odf_common::core::PackageWriter;

    use super::{CURRENT, PEAK};

    const ODS_ROWS: usize = 12_000;
    const ODS_COLUMNS: usize = 8;
    const ODP_SLIDES: usize = 2_000;
    const MEDIA_BYTES: usize = 4 * 1024 * 1024;

    /// A sink that keeps only a length and a SipHash of what it accepts.
    pub(super) struct Digest {
        hasher: std::collections::hash_map::DefaultHasher,
        bytes: u64,
        dump: Option<Vec<u8>>,
    }

    impl Digest {
        fn new(dump: bool) -> Self {
            Self {
                hasher: std::collections::hash_map::DefaultHasher::new(),
                bytes: 0,
                dump: dump.then(Vec::new),
            }
        }
    }

    impl Write for Digest {
        fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
            self.hasher.write(buffer);
            self.bytes += buffer.len() as u64;
            if let Some(dump) = &mut self.dump {
                dump.extend_from_slice(buffer);
            }
            Ok(buffer.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    fn media() -> Vec<u8> {
        // Incompressible, deterministic payload.
        let mut state = 0x9e37_79b9_7f4a_7c15_u64;
        (0..MEDIA_BYTES)
            .map(|_| {
                state ^= state << 13;
                state ^= state >> 7;
                state ^= state << 17;
                state as u8
            })
            .collect()
    }

    fn ods() -> Vec<u8> {
        let mut content = String::from(concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Data">"#,
            r#"<office:dde-source office:dde-application="litchi-inert-fixture" office:dde-topic="never-opened" office:dde-item="cached-values" office:automatic-update="false"/>"#,
            r#"<table:table-column table:number-columns-repeated="8"/>"#,
        ));
        for row in 0..ODS_ROWS {
            content.push_str("<table:table-row>");
            for column in 0..ODS_COLUMNS {
                content.push_str(&format!(
                    r#"<table:table-cell office:value-type="string"><text:p>r{row}c{column}</text:p></table:table-cell>"#
                ));
            }
            content.push_str("</table:table-row>");
        }
        content.push_str(
            "</table:table></office:spreadsheet></office:body></office:document-content>",
        );
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer
            .set_mimetype("application/vnd.oasis.opendocument.spreadsheet")
            .expect("mimetype");
        writer
            .add_file("content.xml", content.as_bytes())
            .expect("content");
        writer
            .add_file("Pictures/blob.bin", &media())
            .expect("media");
        writer.finish_to_bytes().expect("ods")
    }

    fn odp() -> Vec<u8> {
        let mut content = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4"><office:body><office:presentation>"#,
        );
        for slide in 0..ODP_SLIDES {
            content.push_str(&format!(
                r#"<draw:page draw:name="page{}"><draw:frame presentation:class="title"><draw:text-box><text:p>Slide {slide}</text:p></draw:text-box></draw:frame><draw:frame presentation:class="outline"><draw:text-box><text:p>Body text of slide {slide} with a few words.</text:p></draw:text-box></draw:frame></draw:page>"#,
                slide + 1
            ));
        }
        content.push_str("</office:presentation></office:body></office:document-content>");
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer
            .set_mimetype("application/vnd.oasis.opendocument.presentation")
            .expect("mimetype");
        writer
            .add_file("content.xml", content.as_bytes())
            .expect("content");
        writer
            .add_file("Pictures/blob.bin", &media())
            .expect("media");
        writer.finish_to_bytes().expect("odp")
    }

    /// The route under test (the only part that differs between the twins).
    mod route {
        use std::sync::Arc;

        use litchi_core::ReadAt;

        use super::Digest;

        pub(super) const NAME: &str = "after (record 0358a families)";

        fn cell(text: &str) -> litchi_ods::Cell {
            litchi_ods::Cell::new(litchi_ods::CellValue::Text(text.to_string()), text)
        }

        pub(super) fn ods_cells(source: Arc<dyn ReadAt>, count: usize, sink: &mut Digest) {
            let owner = litchi_ods::SourceBackedSpreadsheet::from_read_at(source).expect("open");
            let mut edit = owner.edit_cells().expect("edit");
            for index in 0..count {
                edit.set_cell("Data", index * 97, index % 8, cell("changed"))
                    .expect("set");
            }
            let commit = edit.commit().expect("commit");
            commit.package().write_to(sink).expect("write");
        }

        pub(super) fn ods_dde(source: Arc<dyn ReadAt>, sink: &mut Digest) {
            let owner = litchi_ods::SourceBackedSpreadsheet::from_read_at(source).expect("open");
            let mut edit = owner.edit_dde().expect("edit");
            edit.set_sheet_source("Data", super::dde_source())
                .expect("set");
            let commit = edit.commit().expect("commit");
            commit.package().write_to(sink).expect("write");
        }

        pub(super) fn ods_sheet_metadata(source: Arc<dyn ReadAt>, sink: &mut Digest) {
            let owner = litchi_ods::SourceBackedSpreadsheet::from_read_at(source).expect("open");
            let mut edit = owner.edit_sheet_metadata().expect("edit");
            edit.set_consolidation(Some(super::consolidation()))
                .expect("set");
            let commit = edit.commit().expect("commit");
            commit.package().write_to(sink).expect("write");
        }

        pub(super) fn odp_tail_append(source: Arc<dyn ReadAt>, sink: &mut Digest) {
            let owner = litchi_odp::SourceBackedPresentation::from_read_at(source).expect("open");
            let commit = owner
                .edit_tail_append("Appended title", "Appended body")
                .expect("edit")
                .commit()
                .expect("commit");
            commit.package().write_to(sink).expect("write");
        }
    }

    fn dde_source() -> litchi_ods::dde::Source {
        litchi_ods::dde::Source::new("litchi-inert-fixture", "never-opened", "replacement")
            .expect("source")
            .named("DataSource")
            .expect("name")
            .with_automatic_update(litchi_ods::dde::AutomaticUpdate::Disabled)
    }

    fn consolidation() -> litchi_ods::sheet_metadata::Options {
        litchi_ods::sheet_metadata::Options::new(
            "sum",
            vec!["Data.A1:Data.A2".to_owned(), "Data.B1:Data.B2".to_owned()],
            "Data.C1",
        )
        .expect("consolidation")
    }

    struct Row {
        scenario: &'static str,
        median_us: f64,
        min_us: f64,
        peak_heap_bytes: usize,
        output_bytes: u64,
        output_siphash: String,
    }

    fn measure(
        scenario: &'static str,
        rounds: usize,
        dump: Option<&std::path::Path>,
        source: &Arc<Vec<u8>>,
        run: impl Fn(Arc<dyn ReadAt>, &mut Digest),
    ) -> Row {
        let mut times = Vec::with_capacity(rounds);
        let mut peak = 0;
        let mut output = (0, String::new());
        for round in 0..=rounds {
            let reader: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(source.as_ref().clone()));
            let mut sink = Digest::new(round == 0 && dump.is_some());
            let start = CURRENT.load(Ordering::Relaxed);
            PEAK.store(start, Ordering::Relaxed);
            let clock = Instant::now();
            run(reader, &mut sink);
            let elapsed = clock.elapsed();
            peak = peak.max(PEAK.load(Ordering::Relaxed).saturating_sub(start));
            output = (sink.bytes, format!("{:016x}", sink.hasher.finish()));
            if let (Some(directory), Some(bytes)) = (dump, sink.dump.take()) {
                let name = scenario.replace('/', "_");
                std::fs::write(directory.join(name), bytes).expect("dump");
            }
            if round > 0 {
                times.push(elapsed.as_secs_f64() * 1e6);
            }
        }
        times.sort_by(f64::total_cmp);
        Row {
            scenario,
            median_us: times[times.len() / 2],
            min_us: times[0],
            peak_heap_bytes: peak,
            output_bytes: output.0,
            output_siphash: output.1,
        }
    }

    pub(super) fn run() {
        let mut arguments = std::env::args().skip(1);
        let rounds = arguments
            .next()
            .and_then(|value| value.parse().ok())
            .unwrap_or(15);
        let dump = arguments.next().map(std::path::PathBuf::from);
        if let Some(directory) = &dump {
            std::fs::create_dir_all(directory).expect("dump directory");
        }
        let dump = dump.as_deref();
        let ods = Arc::new(ods());
        let odp = Arc::new(odp());
        let rows = [
            measure("ods/cells1", rounds, dump, &ods, |source, sink| {
                route::ods_cells(source, 1, sink);
            }),
            measure("ods/cells100", rounds, dump, &ods, |source, sink| {
                route::ods_cells(source, 100, sink);
            }),
            measure("ods/dde", rounds, dump, &ods, route::ods_dde),
            measure(
                "ods/sheet_metadata",
                rounds,
                dump,
                &ods,
                route::ods_sheet_metadata,
            ),
            measure(
                "odp/tail_append",
                rounds,
                dump,
                &odp,
                route::odp_tail_append,
            ),
        ];
        let rows = rows
            .iter()
            .map(|row| {
                serde_json::json!({
                    "scenario": row.scenario,
                    "route": route::NAME,
                    "samples": rounds,
                    "median_us": row.median_us,
                    "min_us": row.min_us,
                    "peak_heap_bytes": row.peak_heap_bytes,
                    "output_bytes": row.output_bytes,
                    "output_siphash": row.output_siphash,
                    "source_bytes": if row.scenario.starts_with("odp") { odp.len() } else { ods.len() },
                })
            })
            .collect::<Vec<_>>();
        println!("{}", serde_json::to_string_pretty(&rows).expect("json"));
    }
}

fn main() {
    #[cfg(all(feature = "ods", feature = "odp"))]
    probe::run();
    #[cfg(not(all(feature = "ods", feature = "odp")))]
    eprintln!("enable the ods and odp features");
}
