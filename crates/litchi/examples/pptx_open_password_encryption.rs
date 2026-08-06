use litchi::crypto::ooxml::{self, Mode};
use litchi::pptx::Package;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    const PASSWORD: &str = "Secret123";

    println!("Generating PPTX files with open-password encryption...\n");

    generate_standard_2007(PASSWORD)?;
    generate_agile(PASSWORD)?;

    println!("\nDone. Files written:");
    println!("  - pptx_open_password_standard2007.pptx");
    println!("  - pptx_open_password_agile.pptx");
    println!("Password: {}", PASSWORD);

    Ok(())
}

fn generate_standard_2007(password: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide = pres.add_slide()?;
        slide.set_title("Standard 2007 Encryption");
        slide.add_text_box(
            "This presentation is encrypted with Standard 2007.\nPassword: Secret123",
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
            "This presentation is encrypted with Agile encryption.\nPassword: Secret123",
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
    let clear_package = package.to_bytes()?;
    let encrypted_package = ooxml::encrypt(clear_package, password, mode)?;
    std::fs::write(output, &encrypted_package)?;

    // Exercise the canonical read path as well: decrypt the package through
    // the bounded crypto service, then validate the clear OPC graph with the
    // typed PPTX facade.
    let opened = ooxml::open(encrypted_package, password)?;
    assert_eq!(opened.mode(), Some(mode));
    Package::from_vec(opened.into_bytes())?;
    Ok(())
}
