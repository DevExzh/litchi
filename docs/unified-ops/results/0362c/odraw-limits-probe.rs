//! Required-limit ODraw entry points. The Python runner adapts baseline spellings.
use litchi_odraw::{image, shape, Limits, Parser, Record, RecordKind};
use std::{hint::black_box, time::Instant};

fn record(header: u16, kind: u16, body: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&header.to_le_bytes());
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&(body.len() as u32).to_le_bytes());
    out.extend_from_slice(body);
    out
}
fn png(size: usize) -> Vec<u8> {
    let mut body = vec![0; 16];
    body.push(255);
    body.extend(std::iter::repeat_n(0xa5, size));
    record(0x6e00, 0xf01e, &body)
}
fn shape_bytes() -> Vec<u8> {
    let mut body = Vec::new();
    use litchi_odraw::builder::ShapeBuilder;
    use litchi_odraw::write::{self, Atom, Container};
    ShapeBuilder::new(shape::Native::RECTANGLE, 7)
        .with_flags(shape::Flags::HAVE_ANCHOR | shape::Flags::HAVE_SPT)
        .write(&mut body).unwrap();
    let anchor: Vec<u8> = [10_i32, 20, 110, 70].into_iter().flat_map(i32::to_le_bytes).collect();
    write::atom(&mut body, 0, Atom::ClientAnchor, &anchor).unwrap();
    write::atom(&mut body, 0, Atom::unknown(0xf123, 0).unwrap(), &[0x33; 32]).unwrap();
    let mut result = Vec::new();
    write::container(&mut result, 0, Container::Sp, &body).unwrap();
    result
}
fn main() {
    let mut args = std::env::args().skip(1);
    let route = args.next().unwrap();
    let iterations: u64 = args.next().unwrap().parse().unwrap();
    let small = png(64);
    let large = png(2 * 1024 * 1024);
    let shape = shape_bytes();
    let (shape_record, _) = Record::parse(&shape, 0).unwrap();
    let mut entry = vec![6, 6];
    entry.extend_from_slice(&[0; 16]);
    entry.extend_from_slice(&255_u16.to_le_bytes());
    entry.extend_from_slice(&(small.len() as u32).to_le_bytes());
    entry.extend_from_slice(&1_u32.to_le_bytes());
    entry.extend_from_slice(&0_u32.to_le_bytes());
    entry.extend_from_slice(&[0; 4]);
    let entry = record(0x62, 0xf007, &entry);
    let store_bytes = record(0x1f, 0xf001, &entry);
    let limits = image::Limits::default();
    let store = image::Store::parse(&store_bytes, limits).unwrap();
    let context = image::Context::new().with_delay(image::Delay::new(&small, limits));
    let id = image::Id::new(1).unwrap();
    let drawing = record(0xf, 0xf002, &shape.repeat(128));
    let root = Parser::new(&drawing).root().unwrap().unwrap();
    let operation = || match route.as_str() {
        "blip-small" | "blip-large" => {
            let bytes = if route == "blip-small" { &small } else { &large };
            let blip = black_box(image::Blip::parse(black_box(bytes), black_box(limits)).unwrap());
            assert_eq!(blip.data().as_ptr(), bytes[25..].as_ptr());
            blip.data().len() as u64 + u64::from(blip.data()[0])
        }
        "delayed" => {
            let blip = black_box(black_box(&store).resolve(black_box(id), black_box(context)).unwrap().unwrap());
            assert_eq!(blip.data().as_ptr(), small[25..].as_ptr());
            blip.data().len() as u64 + u64::from(blip.data()[0])
        }
        "shape-parse" => {
            let shapes = black_box(shape::parse(black_box(&shape), black_box(Limits::default())).unwrap());
            assert_eq!(shapes[0].meta().find(RecordKind::Unknown(0xf123)).unwrap().unwrap().data(), &[0x33; 32]);
            u64::from(shapes[0].id())
        }
        "shape-record" => {
            let s = black_box(shape::Shape::from_record(black_box(shape_record.clone()), Limits::default()).unwrap());
            u64::from(s.id())
        }
        "recursive" => {
            let records = black_box(black_box(&root).find_recursive(RecordKind::SpContainer, black_box(Limits::default())).unwrap());
            records.len() as u64 + records.iter().map(|r| u64::from(r.len())).sum::<u64>()
        }
        _ => panic!("unknown route"),
    };
    let expected = operation();
    for _ in 0..8 { assert_eq!(operation(), expected); }
    let start = Instant::now();
    let mut checksum = 0_u64;
    for _ in 0..iterations { checksum = checksum.wrapping_add(black_box(operation())); }
    let ns = start.elapsed().as_nanos();
    assert_eq!(checksum, expected.wrapping_mul(iterations));
    println!("{{\"route\":\"{route}\",\"iterations\":{iterations},\"ns\":{ns},\"checksum\":{checksum},\"store_size\":{},\"shape_size\":{}}}", std::mem::size_of::<image::Store<'_>>(), std::mem::size_of::<shape::Shape<'_>>());
}
