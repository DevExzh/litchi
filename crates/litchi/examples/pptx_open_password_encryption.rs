use litchi::pptx::{Package, encryption::Mode};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let password = std::env::var("LITCHI_PPTX_PASSWORD")?;

    println!("Generating PPTX files with open-password encryption...\n");

    generate_standard_2007(&password)?;
    generate_agile(&password)?;

    println!("\nDone. Files written:");
    println!("  - pptx_open_password_standard2007.pptx");
    println!("  - pptx_open_password_agile.pptx");

    Ok(())
}

fn generate_standard_2007(password: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide = pres.add_slide()?;
        slide.set_title("Standard 2007 Encryption");
        slide.add_text_box(
            "This presentation is encrypted with Standard 2007.",
            914_400,
            1_828_800,
            7_315_200,
            914_400,
        );
    }

    let output = "pptx_open_password_standard2007.pptx";
    println!(
        "Saving Standard 2007-encrypted presentation to {}...",
        output
    );
    write_encrypted(pkg, output, password, Mode::Standard)?;
    Ok(())
}

fn generate_agile(password: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide = pres.add_slide()?;
        slide.set_title("Agile Encryption");
        slide.add_text_box(
            "This presentation is encrypted with Agile encryption.",
            914_400,
            1_828_800,
            7_315_200,
            914_400,
        );
    }

    let output = "pptx_open_password_agile.pptx";
    println!("Saving Agile-encrypted presentation to {}...", output);
    write_encrypted(pkg, output, password, Mode::Agile)?;
    Ok(())
}

fn write_encrypted(
    mut package: Package,
    output: &str,
    password: &str,
    mode: Mode,
) -> Result<(), Box<dyn std::error::Error>> {
    package.save_encrypted(output, password, mode)?;

    // Exercise the managed read path as well. The package retains its source
    // encryption profile so subsequent writes cannot silently downgrade it.
    let reopened = Package::open_with_password(output, password)?;
    assert_eq!(reopened.encryption(), Some(mode));
    Ok(())
}
