//! Paired Graph owner probe; synthetic bytes are generated identically in both builds.
use litchi_biff::{Encoder, Kind, Records};
use litchi_cfb::OleWriter;
use litchi_ograph::{Limits, Package, chart};
use std::{hint::black_box, io::Cursor, time::Instant};
const BOF: Kind = Kind::from_wire(0x0809);
const EOF: Kind = Kind::from_wire(0x000a);
const BOF_BYTES: usize = 16;
const OGRAPH_VERSION: u16 = 0x0680;
const GLOBALS: u16 = 5;
const CHART_SHEET: u16 = 0x8000;
const REQUIRED_PLATFORM_FLAGS: u32 = 9;
#[cfg(feature = "before")]
type Owner = litchi_ograph::Snapshot;
#[cfg(not(feature = "before"))]
type Owner = Package;
fn open(bytes: Vec<u8>) -> Owner {
    #[cfg(feature = "before")]
    {
        Owner::from_bytes(bytes, Limits::default()).unwrap()
    }
    #[cfg(not(feature = "before"))]
    {
        Owner::open(bytes).unwrap()
    }
}
fn bytes(owner: &Owner) -> &[u8] {
    #[cfg(feature = "before")]
    {
        owner.bytes()
    }
    #[cfg(not(feature = "before"))]
    {
        owner.as_bytes()
    }
}
fn edit(owner: &Owner) -> litchi_ograph::Edit {
    #[cfg(feature = "before")]
    {
        owner.edit().unwrap()
    }
    #[cfg(not(feature = "before"))]
    {
        owner.edit()
    }
}
fn apply(owner: &Owner, patch: &litchi_ograph::Patch) -> Owner {
    #[cfg(feature = "before")]
    {
        patch.apply(owner).unwrap()
    }
    #[cfg(not(feature = "before"))]
    {
        owner.apply(patch).unwrap().into_package()
    }
}
fn fixture(extra: usize) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &editable_workbook(true))
        .unwrap();
    writer
        .create_stream(&["\u{1}CompObj"], &vec![0xa7; extra])
        .unwrap();
    writer
        .create_stream(&["\u{1}Ole"], &[3, 1, 4, 1, 5])
        .unwrap();
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    out.into_inner()
}
fn main() {
    let args: Vec<_> = std::env::args().collect();
    let route = &args[1];
    let iterations: usize = args[2].parse().unwrap();
    let extra: usize = args[3].parse().unwrap();
    let input = fixture(extra);
    // This existing ownership-transfer API is identical on both revisions.
    // Keep its setup independent of the replaced Snapshot API: that API's
    // extra whole-file copies change the large-allocation cache before timing.
    if route == "open_payload" {
        let candidate = input.clone();
        let pointer = candidate.as_ptr();
        let capacity = candidate.capacity();
        let payload = Package::open(candidate).unwrap().finish().into_bytes();
        assert_eq!(payload, input);
        assert_eq!(payload.as_ptr(), pointer);
        assert_eq!(payload.capacity(), capacity);
        drop(payload);
        let run = || {
            Package::open(black_box(input.clone()))
                .unwrap()
                .finish()
                .into_bytes()
                .len()
        };
        for _ in 0..100 {
            black_box(run());
        }
        let start = Instant::now();
        let mut checksum = 0usize;
        for _ in 0..iterations {
            checksum = checksum.wrapping_add(black_box(run()));
        }
        let ns = start.elapsed().as_nanos();
        assert_eq!(checksum, input.len().wrapping_mul(iterations));
        println!(
            "{{\"route\":\"{}\",\"iterations\":{},\"ns\":{},\"checksum\":{},\"input_bytes\":{}}}",
            route,
            iterations,
            ns,
            checksum,
            input.len()
        );
        return;
    }
    let base = open(input.clone());
    let rect = chart::Rect {
        width: 4_200 << 16,
        height: 3_100 << 16,
        ..chart::Rect::default()
    };
    let mut initial = edit(&base);
    initial.set_rect(rect).unwrap();
    let commit = initial.commit().unwrap();
    let patch = commit.patch().clone();
    let inverse = patch.inverse();
    let after = commit.into_package();
    assert_ne!(bytes(&after), input);
    assert_eq!(bytes(&apply(&open(input.clone()), &patch)), bytes(&after));
    assert_eq!(
        bytes(&apply(&open(bytes(&after).to_vec()), &inverse)),
        input
    );
    assert_eq!(bytes(&edit(&base).commit().unwrap().into_package()), input);
    let run = || -> usize {
        match route.as_str() {
            "open_owner" => bytes(&open(black_box(input.clone()))).len(),
            "open_payload" => Package::open(black_box(input.clone()))
                .unwrap()
                .finish()
                .into_bytes()
                .len(),
            "begin_edit" => {
                let e = edit(black_box(&base));
                black_box(&e);
                e.len()
            },
            "noop_commit" => bytes(&edit(black_box(&base)).commit().unwrap().into_package()).len(),
            "edit_commit" => {
                let mut e = edit(black_box(&base));
                e.set_rect(black_box(rect)).unwrap();
                bytes(&e.commit().unwrap().into_package()).len()
            },
            "apply" => bytes(&apply(black_box(&base), black_box(&patch))).len(),
            "inverse" => bytes(&apply(black_box(&after), black_box(&inverse))).len(),
            "history128" => {
                let mut current = open(input.clone());
                let mut history = Vec::with_capacity(128);
                for index in 0..128 {
                    let mut e = edit(&current);
                    e.set_rect(chart::Rect {
                        width: rect.width + index,
                        ..rect
                    })
                    .unwrap();
                    let c = e.commit().unwrap();
                    history.push(c.patch().clone());
                    current = c.into_package();
                }
                black_box(&history);
                bytes(&current).len() + history.len()
            },
            _ => panic!("unknown route"),
        }
    };
    let expected = run();
    for _ in 0..3 {
        black_box(run());
    }
    let start = Instant::now();
    let mut checksum = 0usize;
    for _ in 0..iterations {
        checksum = checksum.wrapping_add(black_box(run()));
    }
    let ns = start.elapsed().as_nanos();
    assert_eq!(checksum, expected.wrapping_mul(iterations));
    println!(
        "{{\"route\":\"{}\",\"iterations\":{},\"ns\":{},\"checksum\":{},\"input_bytes\":{}}}",
        route,
        iterations,
        ns,
        checksum,
        input.len()
    );
}

fn bof(doc_type: u16) -> [u8; BOF_BYTES] {
    let mut payload = [0; BOF_BYTES];
    payload[0..2].copy_from_slice(&OGRAPH_VERSION.to_le_bytes());
    payload[2..4].copy_from_slice(&doc_type.to_le_bytes());
    payload[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
    payload[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
    payload[8..12].copy_from_slice(&(REQUIRED_PLATFORM_FLAGS | (6 << 14)).to_le_bytes());
    payload[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    payload
}

fn editable_chart(with_cache: bool) -> Vec<u8> {
    let mut out = Encoder::new(Limits::default().biff).expect("BIFF limits");
    out.push(BOF, &bof(CHART_SHEET)).expect("chart BOF");

    let mut rect = [0u8; 16];
    rect[8..12].copy_from_slice(&(4_000_i32 << 16).to_le_bytes());
    rect[12..16].copy_from_slice(&(3_000_i32 << 16).to_le_bytes());
    out.push(Kind::from_wire(0x1002), &rect)
        .expect("Chart record");
    out.push(Kind::from_wire(0x1033), &[]).expect("chart Begin");
    out.push(Kind::from_wire(0x00A0), &[1, 0, 1, 0])
        .expect("Scl");
    out.push(Kind::from_wire(0x1064), &[0, 0, 1, 0, 0, 0, 1, 0])
        .expect("PlotGrowth");
    out.push(Kind::from_wire(0x1044), &[2, 0, 0, 0])
        .expect("ShtProps");
    out.push(Kind::from_wire(0x1046), &[1, 0])
        .expect("AxesUsed");
    out.push(Kind::from_wire(0x1041), &[0; 18])
        .expect("AxisParent");
    out.push(Kind::from_wire(0x1033), &[])
        .expect("parent Begin");
    let mut pos = [0u8; 20];
    pos[0..2].copy_from_slice(&2_u16.to_le_bytes());
    pos[2..4].copy_from_slice(&2_u16.to_le_bytes());
    pos[12..14].copy_from_slice(&4_000_i16.to_le_bytes());
    pos[16..18].copy_from_slice(&4_000_i16.to_le_bytes());
    out.push(Kind::from_wire(0x104F), &pos).expect("Pos");
    out.push(Kind::from_wire(0x1034), &[]).expect("parent End");
    out.push(Kind::from_wire(0x1035), &[]).expect("PlotArea");
    out.push(Kind::from_wire(0x1034), &[]).expect("chart End");

    let mut dimensions = [0u8; 14];
    if with_cache {
        dimensions[4..8].copy_from_slice(&1_u32.to_le_bytes());
        dimensions[10..12].copy_from_slice(&1_u16.to_le_bytes());
    }
    out.push(Kind::from_wire(0x0200), &dimensions)
        .expect("Dimensions");
    if with_cache {
        let mut number = [0u8; 15];
        number[5..7].copy_from_slice(&4_u16.to_le_bytes());
        number[7..15].copy_from_slice(&1.25_f64.to_le_bytes());
        out.push(Kind::from_wire(0x0003), &number)
            .expect("Graph Number");
    }
    out.push(Kind::from_wire(0x7777), &[0xA1, 0xB2, 0xC3])
        .expect("opaque chart record");
    out.push(EOF, &[]).expect("chart EOF");
    out.finish()
}

fn editable_workbook(with_cache: bool) -> Vec<u8> {
    let chart = editable_chart(with_cache);
    let mut out = Encoder::new(Limits::default().biff).expect("BIFF limits");
    out.push(BOF, &bof(GLOBALS)).expect("globals BOF");
    out.push(Kind::from_wire(0x7776), &[0x10, 0x20])
        .expect("opaque globals record");
    out.push(EOF, &[]).expect("globals EOF");
    for item in Records::new(&chart, litchi_biff::Limits::default()).expect("valid BIFF limits") {
        out.push_ref(item.expect("chart record"))
            .expect("chart replay");
    }
    out.finish()
}
