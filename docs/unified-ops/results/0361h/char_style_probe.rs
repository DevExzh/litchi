//! Compile this shared probe against each protected library. Only the call
//! adapter changes under `--cfg char_style_policy`; fixture and timed bodies
//! are identical. `--cfg allocation_probe` enables a separate allocation run.

use litchi_iwa_text_wire::{self as wire, CharacterProperties};
use std::{hint::black_box, time::Instant};

#[cfg(char_style_policy)]
macro_rules! call {
    ($name:ident($($arg:expr),* $(,)?)) => {
        wire::$name($($arg,)* litchi_iwa_common::WireLimits::DEFAULT)
    };
}
#[cfg(not(char_style_policy))]
macro_rules! call {
    ($name:ident($($arg:expr),* $(,)?)) => { wire::$name($($arg),*) };
}

#[cfg(allocation_probe)]
mod allocation {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering::Relaxed};
    pub struct Counting;
    static ACTIVE: AtomicBool = AtomicBool::new(false);
    static CALLS: AtomicUsize = AtomicUsize::new(0);
    static REQUESTED: AtomicUsize = AtomicUsize::new(0);
    static LIVE: AtomicUsize = AtomicUsize::new(0);
    static PEAK: AtomicUsize = AtomicUsize::new(0);
    fn added(size: usize) {
        CALLS.fetch_add(1, Relaxed);
        REQUESTED.fetch_add(size, Relaxed);
        let live = LIVE.fetch_add(size, Relaxed) + size;
        PEAK.fetch_max(live, Relaxed);
    }
    unsafe impl GlobalAlloc for Counting {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            let result = unsafe { System.alloc(layout) };
            if ACTIVE.load(Relaxed) && !result.is_null() {
                added(layout.size());
            }
            result
        }
        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            let result = unsafe { System.alloc_zeroed(layout) };
            if ACTIVE.load(Relaxed) && !result.is_null() {
                added(layout.size());
            }
            result
        }
        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            if ACTIVE.load(Relaxed) {
                LIVE.fetch_sub(layout.size(), Relaxed);
            }
            unsafe {
                System.dealloc(pointer, layout);
            }
        }
        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
            let result = unsafe { System.realloc(pointer, layout, size) };
            if ACTIVE.load(Relaxed) && !result.is_null() {
                LIVE.fetch_sub(layout.size(), Relaxed);
                added(size);
            }
            result
        }
    }
    pub fn start() {
        ACTIVE.store(true, Relaxed);
    }
    pub fn stop() -> (usize, usize, usize, usize) {
        ACTIVE.store(false, Relaxed);
        (
            CALLS.load(Relaxed),
            REQUESTED.load(Relaxed),
            PEAK.load(Relaxed),
            LIVE.load(Relaxed),
        )
    }
}
#[cfg(allocation_probe)]
#[global_allocator]
static ALLOCATOR: allocation::Counting = allocation::Counting;

fn varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 128 {
        output.push(value as u8 | 128);
        value >>= 7;
    }
    output.push(value as u8);
}
fn scalar(output: &mut Vec<u8>, field: u64, value: u64) {
    varint(output, field << 3);
    varint(output, value);
}
fn bytes(output: &mut Vec<u8>, field: u64, value: &[u8]) {
    varint(output, field << 3 | 2);
    varint(output, value.len() as u64);
    output.extend_from_slice(value);
}
fn reference(value: u64) -> Vec<u8> {
    let mut result = Vec::new();
    scalar(&mut result, 1, value);
    result
}

struct Fixture {
    storage: Vec<u8>,
    style: Vec<u8>,
    stylesheet: Vec<u8>,
    properties: CharacterProperties,
    text_length: usize,
}

fn variation(properties: &CharacterProperties) -> Vec<u8> {
    #[cfg(char_style_policy)]
    {
        call!(character_style_variation(Some(100), 3, properties)).unwrap()
    }
    #[cfg(not(char_style_policy))]
    {
        call!(character_style_variation(Some(100), 3, properties))
    }
}

impl Fixture {
    fn new(runs: usize) -> Self {
        let mut properties = CharacterProperties::new();
        properties.bold = Some(true);
        properties.italic = Some(false);
        properties.size = Some(18.5);
        properties.font_name = Some("Helvetica".into());
        properties.color = Some([24, 139, 220]);
        properties.underline = Some(true);
        properties.strikethrough = Some(false);
        let style = variation(&properties);
        let text_length = runs * 8;
        let mut storage = Vec::new();
        bytes(&mut storage, 3, &vec![b'a'; text_length]);
        // Unknown opaque data is intentionally not recursively parsed.
        bytes(&mut storage, 100, &[0xff; 128]);
        let mut table = Vec::new();
        for index in 0..runs {
            let mut entry = Vec::new();
            scalar(&mut entry, 1, (index * 8) as u64);
            bytes(&mut entry, 2, &reference(100 + (index % 8) as u64));
            bytes(&mut table, 1, &entry);
        }
        bytes(&mut storage, 8, &table);
        let mut stylesheet = Vec::new();
        for index in 0..runs {
            bytes(&mut stylesheet, 1, &reference(100 + index as u64));
        }
        for index in 0..runs {
            let mut entry = Vec::new();
            bytes(&mut entry, 1, &reference(100 + index as u64));
            bytes(&mut entry, 2, &reference(300 + index as u64));
            bytes(&mut stylesheet, 5, &entry);
        }
        bytes(&mut stylesheet, 100, &[0xff; 128]);
        Self {
            storage,
            style,
            stylesheet,
            properties,
            text_length,
        }
    }

    fn validate(&self, output: &std::path::Path) {
        std::fs::create_dir_all(output).unwrap();
        let runs = call!(char_style_runs(&self.storage)).unwrap();
        assert_eq!(runs.len(), self.text_length / 8);
        let styles = call!(styles_in_span(
            &self.storage,
            self.text_length,
            2..self.text_length - 2
        ))
        .unwrap();
        assert_eq!(styles.len(), 8.min(runs.len()));
        let restyled = call!(restyle_span(
            &self.storage,
            self.text_length,
            2..self.text_length - 2,
            |style| style.map(|id| id + 1_000)
        ))
        .unwrap();
        let changed = call!(char_style_runs(&restyled)).unwrap();
        assert_eq!(changed.first().unwrap().style, Some(100));
        assert_eq!(changed[1].start, 2);
        assert_eq!(changed[1].style, Some(1_100));
        assert_eq!(changed.last().unwrap().start, self.text_length - 2);
        let properties = call!(character_style_properties(&self.style)).unwrap();
        assert_eq!(properties, self.properties);
        let parent = call!(character_style_parent(&self.style)).unwrap();
        assert_eq!(parent, Some(100));
        let sheet = call!(stylesheet_with_style(
            &self.stylesheet,
            900,
            Some(100 + runs.len() as u64 / 2)
        ))
        .unwrap();
        for (name, value) in [
            ("storage.bin", self.storage.clone()),
            ("style.bin", self.style.clone()),
            ("stylesheet.bin", self.stylesheet.clone()),
            ("runs.txt", format!("{runs:?}").into_bytes()),
            ("styles.txt", format!("{styles:?}").into_bytes()),
            ("restyled.bin", restyled),
            ("variation.bin", variation(&self.properties)),
            ("properties.txt", format!("{properties:?}").into_bytes()),
            ("parent.txt", format!("{parent:?}").into_bytes()),
            ("registered.bin", sheet),
        ] {
            std::fs::write(output.join(name), value).unwrap();
        }
    }
}

fn run<T>(count: usize, mut operation: impl FnMut() -> T) -> u128 {
    let start = Instant::now();
    for _ in 0..count {
        drop(black_box(operation()));
    }
    start.elapsed().as_nanos()
}

fn main() {
    let arguments: Vec<String> = std::env::args().collect();
    let route = arguments.get(1).expect("route or validate");
    let runs: usize = arguments
        .get(2)
        .expect("fixture run count")
        .parse()
        .unwrap();
    assert!((2..=512).contains(&runs));
    let fixture = Fixture::new(runs);
    if route == "validate" {
        fixture.validate(std::path::Path::new(
            arguments.get(3).expect("output directory"),
        ));
        return;
    }
    let count: usize = arguments.get(3).expect("call count").parse().unwrap();
    assert!(count > 0);
    #[cfg(allocation_probe)]
    allocation::start();
    let elapsed = match route.as_str() {
        "runs" => run(count, || {
            call!(char_style_runs(black_box(&fixture.storage))).unwrap()
        }),
        "styles" => run(count, || {
            call!(styles_in_span(
                black_box(&fixture.storage),
                fixture.text_length,
                2..fixture.text_length - 2
            ))
            .unwrap()
        }),
        "restyle" => run(count, || {
            call!(restyle_span(
                black_box(&fixture.storage),
                fixture.text_length,
                2..fixture.text_length - 2,
                |style| style.map(|id| id + 1_000)
            ))
            .unwrap()
        }),
        "variation" => run(count, || variation(black_box(&fixture.properties))),
        "properties" => run(count, || {
            call!(character_style_properties(black_box(&fixture.style))).unwrap()
        }),
        "parent" => run(count, || {
            call!(character_style_parent(black_box(&fixture.style))).unwrap()
        }),
        "stylesheet" => run(count, || {
            call!(stylesheet_with_style(
                black_box(&fixture.stylesheet),
                900,
                Some(100 + runs as u64 / 2)
            ))
            .unwrap()
        }),
        _ => panic!("unknown route"),
    };
    #[cfg(allocation_probe)]
    {
        let (calls, requested, peak, live) = allocation::stop();
        assert_eq!(live, 0);
        println!(
            "{{\"route\":\"{route}\",\"runs\":{runs},\"count\":{count},\"allocation_calls\":{calls},\"requested_bytes\":{requested},\"peak_live_bytes\":{peak},\"retained_bytes\":{live}}}"
        );
        black_box(elapsed);
    }
    #[cfg(not(allocation_probe))]
    println!(
        "{{\"route\":\"{route}\",\"runs\":{runs},\"count\":{count},\"elapsed_ns\":{elapsed}}}"
    );
}
