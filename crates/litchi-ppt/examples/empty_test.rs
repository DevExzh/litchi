//! Empty PPT test - no slides at all

use litchi_ppt::writer::Writer;

fn main() {
    println!("Creating empty PPT with no slides...");

    let mut writer = Writer::new();
    writer.save("output_empty.ppt").unwrap();

    println!("Created output_empty.ppt");
}
