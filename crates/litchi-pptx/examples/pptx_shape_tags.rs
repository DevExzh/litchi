//! Replace or add one inert tag on a named shape on the first slide.

use std::io;

use litchi_pptx::Package;
use litchi_pptx::tag::{List, Tag};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let input = args.next().ok_or_else(usage)?;
    let output = args.next().ok_or_else(usage)?;
    let shape = args.next().ok_or_else(usage)?;
    let name = args.next().ok_or_else(usage)?;
    let value = args.next().ok_or_else(usage)?;
    if args.next().is_some() {
        return Err(usage().into());
    }

    let package = Package::open(input)?;
    let mut tags = package
        .shape_tags(0_usize, shape.as_str())?
        .unwrap_or_else(List::new);
    let outcome = match tags.set(name.as_str(), value.as_str()) {
        Ok(_) => "replaced",
        Err(litchi_pptx::Error::NameNotFound(_)) => {
            tags.add(Tag::new(name, value)?)?;
            "added"
        },
        Err(error) => return Err(error.into()),
    };
    let mut edit = package.edit_shape_tags()?;
    let _ = edit.set(0_usize, shape.as_str(), tags)?;
    let package = edit.commit()?.into_package();
    package.save(output)?;
    println!("{outcome} shape tag");
    Ok(())
}

fn usage() -> io::Error {
    io::Error::new(
        io::ErrorKind::InvalidInput,
        "usage: pptx_shape_tags <input.pptx> <output.pptx> <shape-name> <tag-name> <value>",
    )
}
