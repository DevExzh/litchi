//! Generate the small ODG candidate used by the headless native metadata receipt.

use litchi_odg::{Drawing, Transition};
use std::{env, fs, io};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args_os();
    let _program = arguments.next();
    let source_path = arguments.next().ok_or_else(|| {
        io::Error::new(io::ErrorKind::InvalidInput, "source ODG path is required")
    })?;
    let destination_path = arguments.next().ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "candidate ODG path is required",
        )
    })?;
    if arguments.next().is_some() {
        return Err(io::Error::new(io::ErrorKind::InvalidInput, "too many arguments").into());
    }

    let source = Drawing::from_bytes(fs::read(&source_path)?)?;
    let mut transition = Transition::new();
    transition.set_transition_type(Some("automatic"))?;
    transition.set_style(Some("dissolve"))?;
    transition.set_speed(Some("medium"))?;
    transition.set_smil_type(Some("fade"))?;
    transition.set_smil_subtype(Some("crossfade"))?;
    transition.set_direction(Some("reverse"))?;
    transition.set_fade_color(Some("#AABBCC"))?;
    transition.set_duration(Some("PT3S"))?;

    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition))?;
    let commit = edit.commit()?;
    let output = commit.snapshot();
    let reopened = Drawing::from_bytes(output.as_bytes().to_vec())?;
    println!("source={}", source_path.to_string_lossy());
    println!("candidate={}", destination_path.to_string_lossy());
    println!("bytes={}", output.as_bytes().len());
    println!("transition={:?}", reopened.pages()[0].transition());
    println!("active={:?}", reopened.active_content());
    fs::write(destination_path, output.as_bytes())?;
    Ok(())
}
