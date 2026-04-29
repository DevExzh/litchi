//! Minimal PPT test - single empty slide

use litchi::ole::ppt::writer::PptWriter;

fn main() {
    println!("Creating minimal PPT with empty slide...");

    let mut writer = PptWriter::new();
    writer.add_slide().unwrap();
    writer.save("output_minimal.ppt").unwrap();

    println!("Created output_minimal.ppt");
}
