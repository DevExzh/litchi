use litchi::ooxml::pptx::{OpenPasswordEncryption, Package, PresentationProtection};

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

        let mut protection = PresentationProtection::new();
        protection.set_open_password(password)?;
        protection.set_open_password_encryption(OpenPasswordEncryption::Standard2007);
        pres.set_protection(protection);
    }

    let output = "pptx_open_password_standard2007.pptx";
    println!(
        "Saving Standard 2007-encrypted presentation to {}...",
        output
    );
    pkg.save(output)?;
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

        let mut protection = PresentationProtection::new();
        protection.set_open_password(password)?;
        protection.set_open_password_encryption(OpenPasswordEncryption::Agile);
        pres.set_protection(protection);
    }

    let output = "pptx_open_password_agile.pptx";
    println!("Saving Agile-encrypted presentation to {}...", output);
    pkg.save(output)?;
    Ok(())
}
