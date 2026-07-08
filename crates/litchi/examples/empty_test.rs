//! Empty PPT test - no slides at all

use litchi::ole::ppt::writer::PptWriter;

fn main() {
    println!("Creating empty PPT with no slides...");

    let mut writer = PptWriter::new();
    writer.save("output_empty.ppt").unwrap();

    println!("Created output_empty.ppt");
}
