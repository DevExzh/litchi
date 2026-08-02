//! Generate a small encrypted DOCX for interoperability verification.

#[cfg(feature = "encryption")]
use std::path::PathBuf;

#[cfg(feature = "encryption")]
use litchi_ooxml::docx::Package;
#[cfg(feature = "encryption")]
use litchi_ooxml::encryption::Mode;

#[cfg(feature = "encryption")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let output = arguments
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("target/office-crypto/encrypted.docx"));
    let password = arguments
        .next()
        .and_then(|value| value.into_string().ok())
        .unwrap_or_else(|| "Litchi-Office-42!".to_string());
    let mode = match arguments.next().and_then(|value| value.into_string().ok()) {
        None => Mode::Agile,
        Some(value) if value.eq_ignore_ascii_case("agile") => Mode::Agile,
        Some(value) if value.eq_ignore_ascii_case("standard") => Mode::Standard,
        Some(value) => {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                format!("unknown encryption mode {value:?}; expected `agile` or `standard`"),
            )
            .into());
        },
    };
    if let Some(parent) = output.parent() {
        std::fs::create_dir_all(parent)?;
    }

    let mut package = Package::new()?;
    package
        .document_mut()?
        .add_heading("Litchi encrypted DOCX verification", 1)?;
    package
        .document_mut()?
        .add_paragraph()
        .add_run_with_text("This package was encrypted by litchi-crypto.");
    package.save_encrypted(&output, &password, mode)?;

    let reopened = Package::open_with_password(&output, &password)?;
    if !reopened
        .document()?
        .text()?
        .contains("This package was encrypted by litchi-crypto.")
    {
        return Err(std::io::Error::other("encrypted DOCX marker did not round-trip").into());
    }

    println!("{}", output.canonicalize()?.display());
    Ok(())
}

#[cfg(not(feature = "encryption"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    Err(std::io::Error::other("enable the litchi-ooxml `encryption` feature").into())
}
