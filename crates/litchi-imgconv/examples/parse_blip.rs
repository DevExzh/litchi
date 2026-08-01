//! Parse an OfficeArt BLIP record and convert it to PNG.
//!
//! OfficeArt grammar is provided by `litchi-odraw`; `litchi-imgconv` consumes
//! that borrowed view with explicit decoding limits.

use std::path::PathBuf;

use litchi_imgconv::{Options, to_png};
use litchi_odraw::image::{Blip, Kind, write::BlipBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let blip_input = args.next();
    let output: PathBuf = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("blip_out.png"));

    let blip_bytes = match blip_input {
        Some(path) => std::fs::read(path)?,
        None => {
            let png = std::fs::read("test-data/images/png/lena.png")?;
            let builder = BlipBuilder::bitmap(Kind::Png, png)?;
            let mut bytes = Vec::new();
            builder.write(&mut bytes)?;
            bytes
        },
    };

    let blip = Blip::parse(&blip_bytes)?;
    println!("kind: {:?}, data bytes: {}", blip.kind(), blip.data().len());
    let png = to_png(&blip, Options::default())?;
    std::fs::write(&output, png)?;
    println!("wrote {}", output.display());
    Ok(())
}
