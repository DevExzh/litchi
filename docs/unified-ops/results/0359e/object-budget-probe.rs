//! Identical public object operations before and after typed budget errors.
use litchi_cfb::OleWriter;
use litchi_ole_common::object::{Edit, Limits, Target, Targets};
use std::{hint::black_box, io::Cursor, time::Instant};

fn targets() -> Targets {
    Targets::one(Target::new("object", ["Object"]).unwrap())
}

fn fixture(opaque: usize, streams: usize) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Object", "Payload"], b"base")
        .unwrap();
    writer
        .create_stream(&["Opaque"], &vec![0xa7; opaque])
        .unwrap();
    for index in 0..streams {
        writer
            .create_stream(&[&format!("Stream{index:04}")], b"inert")
            .unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn main() {
    let args: Vec<_> = std::env::args().collect();
    let route = &args[1];
    let iterations: usize = args[2].parse().unwrap();
    let opaque: usize = args[3].parse().unwrap();
    let streams: usize = args[4].parse().unwrap();
    let input = fixture(opaque, streams);
    let source = Edit::open(input.clone(), targets(), Limits::default())
        .unwrap()
        .staged();
    let path = vec!["Object".to_string(), "Payload".to_string()];
    let mut initial = source.edit();
    initial.put_stream(&path, b"next".to_vec()).unwrap();
    let committed = initial.commit().unwrap();
    let patch = committed.patch().clone();
    let after = committed.package().finish().unwrap();
    assert_eq!(source.finish().unwrap(), input);
    assert_eq!(patch.apply(&input).unwrap(), after);
    assert_eq!(patch.inverse().apply(&after).unwrap(), input);
    assert!(patch.apply(b"wrong source").is_err());
    assert_eq!(
        source.edit().commit().unwrap().package().finish().unwrap(),
        input
    );
    let run = || -> usize {
        match route.as_str() {
            "open" => black_box(Edit::open(black_box(input.clone()), targets(), Limits::default())
                .unwrap())
                .objects()
                .len(),
            "begin_edit" => {
                let edit = source.edit();
                black_box(&edit);
                edit.objects().len()
            },
            "noop_commit" => black_box(source
                .edit()
                .commit()
                .unwrap()
                .package()
                .finish()
                .unwrap())
                .len(),
            "replace_stream" => {
                let mut edit = source.edit();
                edit.put_stream(&path, black_box(b"next".to_vec())).unwrap();
                black_box(edit.commit().unwrap().package().finish().unwrap()).len()
            },
            "apply_patch" => black_box(patch.apply(black_box(&input)).unwrap()).len(),
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
        "{{\"ns\":{ns},\"iterations\":{iterations},\"checksum\":{checksum},\"input_bytes\":{}}}",
        input.len()
    );
}
