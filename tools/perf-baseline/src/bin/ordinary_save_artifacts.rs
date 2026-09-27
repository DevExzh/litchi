//! Export untimed ordinary-save source and policy artifacts.

use std::path::PathBuf;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let mut output = None;
    let mut filesystem_root = None;
    let mut ooxml_files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.to_string_lossy().as_ref() {
            "--output" => {
                if output.is_some() {
                    return Err("--output was specified more than once".into());
                }
                output = Some(PathBuf::from(
                    arguments.next().ok_or("--output requires NEWDIR")?,
                ));
            },
            "--filesystem-root" => {
                if filesystem_root.is_some() {
                    return Err("--filesystem-root was specified more than once".into());
                }
                filesystem_root = Some(PathBuf::from(
                    arguments.next().ok_or("--filesystem-root requires ROOT")?,
                ));
            },
            "--ooxml-file" => {
                ooxml_files.push(PathBuf::from(
                    arguments.next().ok_or("--ooxml-file requires PATH")?,
                ));
            },
            value => return Err(format!("unknown argument {value:?}").into()),
        }
    }
    let output = output.ok_or("--output NEWDIR is required")?;
    litchi_perf_baseline::write_ordinary_save_artifacts(
        &output,
        filesystem_root.as_deref(),
        &ooxml_files,
    )
}
