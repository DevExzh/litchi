//! One-shot: `once <two|one> <path>` performs exactly one open, with or
//! without a preceding detect_file_format, for syscall counting under strace.
use std::hint::black_box;
fn main() {
    let mode = std::env::args().nth(1).unwrap();
    let path = std::env::args().nth(2).unwrap();
    let kind = std::env::args().nth(3).unwrap();
    if mode == "two" {
        black_box(litchi::detect_file_format(&path));
    }
    match kind.as_str() {
        "doc" => {
            black_box(litchi::Document::open(&path).unwrap());
        },
        "prs" => {
            black_box(litchi::Presentation::open(&path).unwrap());
        },
        "wb" => {
            black_box(litchi::sheet::Workbook::open(&path).unwrap());
        },
        _ => panic!(),
    }
}
