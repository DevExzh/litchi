//! Probe one source-backed PPTX slide copy for the change-0454 inventory.
//!
//! The command opens the source and destination independently, plans a copy at
//! the destination position's following slot, and reports whether the bounded
//! native PPTX path publishes or refuses the selected slide.

use litchi_pptx::SourceBackedPresentation;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() != 4 {
        return Err(
            "expected source-path source-position destination-path destination-position".into(),
        );
    }
    let source = SourceBackedPresentation::open(&args[0])?;
    let editor_view = litchi_pptx::SourceBackedPresentation::open(&args[2])?;
    let editor = litchi_pptx::source::Package::from_presentation(&editor_view)?;
    let source_position: usize = args[1].parse()?;
    let destination_position: usize = args[3].parse()?;
    let insertion_position = destination_position.checked_add(1).ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "destination position is too large for following-slot insertion",
        )
    })?;
    match editor.plan_cross_slide_copy(
        &source,
        source_position,
        destination_position,
        insertion_position,
    ) {
        Ok(plan) => {
            let mut output = Vec::new();
            let mut editor_copy = editor.edit_cross_slide_copy()?;
            let copied = editor_copy.stage(&plan)?;
            editor_copy.commit()?.package().write_to(&mut output)?;
            println!(
                "PUBLISHED bytes={} slides={} name={:?}",
                output.len(),
                copied.destination_slide_count(),
                copied.name()
            );
        },
        Err(error) => println!("REFUSED {error:?}"),
    }
    Ok(())
}
