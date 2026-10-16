//! List, move, reorder, remove and create the drawables of one Keynote slide
//! through the selector-first slide-drawable transaction.
//!
//! ```text
//! edit_slide_drawables <input.key> <slide>
//! edit_slide_drawables <input.key> <output.key> <slide> <operation>...
//! ```
//!
//! Operations (selectors address the input slide's drawables):
//!
//! ```text
//! geometry:<drawable>:<x>,<y>,<width>,<height>[,<rotation>]
//! remove:<drawable>
//! z:<drawable>:<position>
//! text:<x>,<y>,<width>,<height>:<text>
//! shape:<rectangle|rounded|ellipse>:<x>,<y>,<width>,<height>:<text>
//! image:<file>:<x>,<y>,<width>,<height>
//! ```

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports the slide inventory and its committed change"
)]

use std::error::Error;
use std::fs::OpenOptions;
use std::path::Path;

use litchi_keynote::slide::drawable::{Edit, Geometry, Preset};
use litchi_keynote::{DrawableSelector, Package, Position, SlideSelector};

const USAGE: &str = "usage: edit_slide_drawables <input.key> <slide> \
                     | edit_slide_drawables <input.key> <output.key> <slide> <operation>...";

fn main() -> Result<(), Box<dyn Error>> {
    let arguments = std::env::args().skip(1).collect::<Vec<_>>();
    match arguments.as_slice() {
        [input, slide] => list(&Package::open(input)?, slide.parse()?),
        [input, output, slide, operations @ ..] if !operations.is_empty() => edit(
            Path::new(input),
            Path::new(output),
            slide.parse()?,
            operations,
        ),
        _ => Err(USAGE.into()),
    }
}

fn list(package: &Package, slide: usize) -> Result<(), Box<dyn Error>> {
    for drawable in package
        .slide_drawable_list(SlideSelector::index(slide))?
        .iter()
    {
        println!(
            "{}: {:?} z={:?} geometry={:?} text={:?}",
            drawable.selector().as_index(),
            drawable.kind(),
            drawable.z_position().map(Position::get),
            drawable.geometry(),
            drawable.text(),
        );
    }
    Ok(())
}

fn geometry(text: &str) -> Result<Geometry, Box<dyn Error>> {
    let values = text
        .split(',')
        .map(str::parse::<f32>)
        .collect::<Result<Vec<_>, _>>()?;
    match values.as_slice() {
        [x, y, width, height] => Ok(Geometry::new(*x, *y, *width, *height)),
        [x, y, width, height, rotation] => {
            Ok(Geometry::new(*x, *y, *width, *height).with_rotation(*rotation))
        },
        _ => Err("geometry is x,y,width,height[,rotation]".into()),
    }
}

fn stage(edit: &mut Edit<'_>, slide: usize, operation: &str) -> Result<(), Box<dyn Error>> {
    let slide = SlideSelector::index(slide);
    let parts = operation.splitn(3, ':').collect::<Vec<_>>();
    match parts.as_slice() {
        ["geometry", drawable, frame] => {
            edit.set_geometry(
                slide,
                DrawableSelector::index(drawable.parse()?),
                geometry(frame)?,
            )?;
        },
        ["remove", drawable] => {
            edit.remove(slide, DrawableSelector::index(drawable.parse()?))?;
        },
        ["z", drawable, position] => {
            edit.move_to(
                slide,
                DrawableSelector::index(drawable.parse()?),
                Position::new(position.parse()?),
            )?;
        },
        ["text", frame, text] => {
            edit.add_text_box(slide, geometry(frame)?, text)?;
        },
        ["shape", preset, rest] => {
            let (frame, text) = rest.split_once(':').ok_or(USAGE)?;
            let preset = match *preset {
                "rectangle" => Preset::Rectangle,
                "rounded" => Preset::RoundedRectangle { radius: 12.0 },
                "ellipse" => Preset::Ellipse,
                _ => return Err("preset is rectangle, rounded or ellipse".into()),
            };
            edit.add_shape(slide, preset, geometry(frame)?, text)?;
        },
        ["image", file, frame] => {
            let bytes = std::fs::read(file)?;
            let name = Path::new(file)
                .file_name()
                .and_then(|name| name.to_str())
                .ok_or("image file name")?;
            edit.add_image(slide, &bytes, name, geometry(frame)?)?;
        },
        _ => return Err(format!("unknown operation {operation:?}").into()),
    }
    Ok(())
}

fn edit(
    input: &Path,
    output: &Path,
    slide: usize,
    operations: &[String],
) -> Result<(), Box<dyn Error>> {
    if input == output {
        return Err("input and output paths must differ".into());
    }
    let package = Package::open(input)?;
    let mut edit = package.edit_slide_drawables();
    for operation in operations {
        stage(&mut edit, slide, operation)?;
    }
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_drawables(&commit.patch().inverse())?;
    if restored
        .package()
        .slide_drawable_list(SlideSelector::index(slide))?
        != package.slide_drawable_list(SlideSelector::index(slide))?
    {
        return Err("inverse patch did not restore the input slide".into());
    }
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(output)?;
    commit.package().write_to(&mut file)?;
    file.sync_all()?;
    let diagnostics = commit.diagnostics();
    println!(
        "operations={}, created={}, removed={}, touched_components={}",
        diagnostics.operations(),
        diagnostics.created_drawables(),
        diagnostics.removed_drawables(),
        diagnostics.touched_components(),
    );
    list(commit.package(), slide)
}
