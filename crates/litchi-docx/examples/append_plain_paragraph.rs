//! Append one plain paragraph to an admitted source-backed DOCX document.
//!
//! Usage: `append_plain_paragraph INPUT.docx NEW_OUTPUT.docx TEXT`
//! The output must not already exist. A sequential output failure can leave a
//! partial new file; callers needing atomic replacement must supply that
//! publication boundary themselves.

#[cfg(any(unix, windows))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::fs::OpenOptions;
    use std::io;

    let mut args = std::env::args_os().skip(1);
    let usage = || {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "usage: append_plain_paragraph INPUT.docx NEW_OUTPUT.docx TEXT",
        )
    };
    let input = args.next().ok_or_else(usage)?;
    let output = args.next().ok_or_else(usage)?;
    let text = args.next().ok_or_else(usage)?;
    if args.next().is_some() {
        return Err(usage().into());
    }
    let text = text
        .into_string()
        .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "paragraph text must be UTF-8"))?;
    let package = litchi_docx::source_backed::Package::from_path(input)?;
    let commit = package.tail_append_plain_paragraph(text).commit()?;
    // Preparation and semantic readback finish before even creating the file.
    let mut sink = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(output)?;
    let publication = commit.write_to_stream(&mut sink)?;
    sink.sync_all()?;
    println!(
        "Appended one paragraph: {} -> {} paragraphs",
        publication.source_proof().paragraph_count,
        publication.candidate_proof().paragraph_count,
    );
    Ok(())
}

#[cfg(not(any(unix, windows)))]
fn main() {
    eprintln!("This filesystem example requires positional file-source support.");
}
