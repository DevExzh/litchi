//! Open a Numbers package and atomically save it to a filesystem path.

use std::error::Error;
use std::path::PathBuf;

use litchi_numbers::Package;

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let source = PathBuf::from(
        arguments
            .next()
            .ok_or("usage: save_package <source.numbers> <destination.numbers>")?,
    );
    let destination = PathBuf::from(arguments.next().ok_or("missing destination Numbers path")?);
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }

    let package = Package::open(&source)?;
    package.save(&destination)?;
    println!("saved {} -> {}", source.display(), destination.display());
    Ok(())
}
