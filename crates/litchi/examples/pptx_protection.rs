//! Presentation protection example - demonstrates security settings.

use litchi::ooxml::pptx::{CryptoAlgorithm, Package, Protection, ProtectionType, SlideProtection};

fn main() {
    println!("=== Presentation Protection Example ===\n");

    // Create unprotected presentation
    let unprotected = Protection::new();
    println!("Unprotected Presentation:");
    println!("  Is protected: {}", unprotected.is_protected());
    println!("  Protection type: {:?}", unprotected.protection_type());
    assert_eq!(unprotected.protection_type(), ProtectionType::None);

    // Create read-only recommended
    let read_only = Protection::new().with_read_only_recommended(true);

    println!("\nRead-Only Recommended:");
    println!("  Is protected: {}", read_only.is_protected());
    println!("  Protection type: {:?}", read_only.protection_type());
    assert_eq!(
        read_only.protection_type(),
        ProtectionType::ReadOnlyRecommended
    );

    // Create with structure protection
    let structure_protected = Protection::new()
        .with_structure_protection(true)
        .with_window_protection(true);

    println!("\nStructure Protected:");
    println!(
        "  Protect structure: {}",
        structure_protected.protect_structure
    );
    println!("  Protect windows: {}", structure_protected.protect_windows);

    // Test password protection
    let mut password_protected = Protection::new();
    if let Err(e) = password_protected.set_modify_password("secret123") {
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
    assert_eq!(
        password_protected.protection_type(),
        ProtectionType::ModifyPassword
    );

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
        CryptoAlgorithm::Sha1,
        CryptoAlgorithm::Sha256,
        CryptoAlgorithm::Sha384,
        CryptoAlgorithm::Sha512,
    ];

    for algo in algorithms {
        let uri = algo.uri();
        let Ok(parsed) = CryptoAlgorithm::from_uri(uri) else {
            println!("  Unexpected unsupported algorithm URI: {uri}");
            return;
        };
        println!("  {:?} -> {}", algo, uri);
        assert_eq!(algo, parsed);
    }

    // Slide-level protection
    println!("\n--- Slide Protection ---");
    let slide_prot = SlideProtection::new().protect_all();

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
    let partial = SlideProtection {
        no_edit_text: true,
        no_resize: true,
        ..Default::default()
    };
    println!("\nPartial slide protection:");
    println!("  No edit text: {}", partial.no_edit_text);
    println!("  No resize: {}", partial.no_resize);
    println!("  No move: {}", partial.no_move);

    println!(
        "\nOpen-password encryption is selected explicitly on Package::save_encrypted, \
         independently of presentation modification protection."
    );

    // Generate XML
    let mut with_password = Protection::new();
    if let Err(error) = with_password.set_modify_password("secret123") {
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

    if let Err(e) = generate_protection_pptx() {
        println!("\nError generating protection PPTX: {e}");
    }

    println!("\n✅ Presentation protection example completed successfully!");
}

fn generate_protection_pptx() -> Result<(), Box<dyn std::error::Error>> {
    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;
        let slide = pres.add_slide()?;
        slide.set_title("Protection Demo");
        slide.add_text_box(
            "This presentation is modify-protected. Try modifying it in PowerPoint.",
            914400,
            1828800,
            7315200,
            914400,
        );
        let mut protection = Protection::new();
        protection.set_modify_password("secret123")?;
        pres.set_protection(protection);
    }
    pkg.save("pptx_protection_modify_password.pptx")?;
    Ok(())
}
