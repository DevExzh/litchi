// Small helper to inspect a generated minimal.ppt with litchi's PPT reader
// This is for debugging the writer: it opens minimal.ppt and prints basic info

use litchi_ppt::Package;

fn main() {
    match Package::open("minimal.ppt") {
        Ok(mut pkg) => match pkg.presentation() {
            Ok(pres) => {
                println!("Opened presentation successfully");
                println!("Slide count: {}", pres.slide_count());
                match pres.text() {
                    Ok(text) => {
                        println!("Total text length: {}", text.len());
                    },
                    Err(e) => {
                        eprintln!("Failed to extract text: {e}");
                    },
                }
            },
            Err(e) => {
                eprintln!("Failed to build Presentation: {e}");
            },
        },
        Err(e) => {
            eprintln!("Failed to open minimal.ppt as Package: {e}");
        },
    }
}
