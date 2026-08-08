//! Presentation protection example - demonstrates security settings.

use litchi::pptx::presentation_properties::metadata::protection::{
    Algorithm, Settings, Slide, Type,
};
use litchi::pptx::{Package, encryption::Mode};

fn main() {
    println!("=== Presentation Protection Example ===\n");

    let password = match std::env::var("LITCHI_PPTX_PASSWORD") {
        Ok(password) => password,
        Err(error) => {
            eprintln!("LITCHI_PPTX_PASSWORD is required: {error}");
            return;
        },
    };

    // Create unprotected settings
    let unprotected = Settings::new();
    println!("Unprotected Presentation:");
    println!("  Is protected: {}", unprotected.is_protected());
    println!("  Protection type: {:?}", unprotected.protection_type());
    assert_eq!(unprotected.protection_type(), Type::None);

    // Create read-only recommended
    let read_only = Settings::new().with_read_only_recommended(true);

    println!("\nRead-Only Recommended:");
    println!("  Is protected: {}", read_only.is_protected());
    println!("  Protection type: {:?}", read_only.protection_type());
    assert_eq!(read_only.protection_type(), Type::ReadOnlyRecommended);

    // Create settings with structure protection
    let structure_protected = Settings::new()
        .with_structure_protection(true)
        .with_window_protection(true);

    println!("\nStructure Protected:");
    println!(
        "  Protect structure: {}",
        structure_protected.protect_structure
    );
    println!("  Protect windows: {}", structure_protected.protect_windows);

    // Test password protection
    let mut password_protected = Settings::new();
    if let Err(e) = password_protected.set_modify_password(&password) {
        println!("\nError setting modify password: {e}");
        return;
    }

    println!("\nPassword Protected:");
    let Some(verifier) = password_protected.modify() else {
        println!("  Internal error: modify verifier was not retained");
        return;
    };
    println!("  Modify protected: true");
    println!("  Algorithm: {:?}", verifier.algorithm());
    println!("  Spin count: {}", verifier.spins());
    println!(
        "  Protection type: {:?}",
        password_protected.protection_type()
    );
    assert_eq!(password_protected.protection_type(), Type::ModifyPassword);

    // Clear password
    password_protected.clear_modify_password();
    println!("\nAfter clearing password:");
    println!(
        "  Modify protected: {}",
        password_protected.modify().is_some()
    );

    // Test crypto algorithms
    println!("\n--- Crypto Algorithms ---");
    let algorithms = [
        Algorithm::Sha1,
        Algorithm::Sha256,
        Algorithm::Sha384,
        Algorithm::Sha512,
    ];

    for algo in algorithms {
        let uri = algo.uri();
        let Ok(parsed) = Algorithm::from_uri(uri) else {
            println!("  Unexpected unsupported algorithm URI: {uri}");
            return;
        };
        println!("  {:?} -> {}", algo, uri);
        assert_eq!(algo, parsed);
    }

    // Slide-level protection
    println!("\n--- Slide Protection ---");
    let slide_prot = Slide::new().protect_all();

    println!("Full slide protection:");
    println!("  No select: {}", slide_prot.no_select);
    println!("  No move: {}", slide_prot.no_move);
    println!("  No resize: {}", slide_prot.no_resize);
    println!("  No edit text: {}", slide_prot.no_edit_text);
    println!("  No ungroup: {}", slide_prot.no_ungroup);
    println!("  No change z-order: {}", slide_prot.no_change_z_order);
    println!("  Is protected: {}", slide_prot.is_protected());
    assert!(slide_prot.is_protected());

    // Partial slide protection
    let partial = Slide {
        no_edit_text: true,
        no_resize: true,
        ..Default::default()
    };
    println!("\nPartial slide protection:");
    println!("  No edit text: {}", partial.no_edit_text);
    println!("  No resize: {}", partial.no_resize);
    println!("  No move: {}", partial.no_move);

    println!(
        "\nOpen-password encryption is applied by the standalone OOXML crypto service, \
         independently of presentation modification protection."
    );

    // Generate XML
    let mut with_password = Settings::new();
    if let Err(error) = with_password.set_modify_password(&password) {
        println!("\nError setting XML example password: {error}");
        return;
    }

    let xml = with_password.to_xml();
    println!("\nGenerated protection XML:");
    println!("  Length: {} bytes", xml.len());
    assert!(xml.contains("modifyVerifier"));
    // We now emit the legacy SID-based form (hashData/saltData) to match
    // PowerPoint's own modifyVerifier output.
    assert!(xml.contains("hashData"));

    if let Err(e) = generate_protection_pptx(&password) {
        println!("\nError generating protection PPTX: {e}");
    }

    println!("\n✅ Presentation protection example completed successfully!");
}

fn generate_protection_pptx(password: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide = pres.add_slide()?;
        slide.set_title("Protection and Encryption Demo");
        slide.add_text_box(
            concat!(
                "This package demonstrates open-password encryption. ",
                "Modify-password settings are modeled and serialized separately ",
                "as PresentationML protection metadata."
            ),
            914400,
            1828800,
            7315200,
            914400,
        );
    }
    pkg.save_plain("pptx_protection_clear.pptx")?;
    pkg.save_encrypted("pptx_protection_open_password.pptx", password, Mode::Agile)?;

    let opened = Package::open_with_password("pptx_protection_open_password.pptx", password)?;
    assert_eq!(opened.encryption(), Some(Mode::Agile));
    Ok(())
}
