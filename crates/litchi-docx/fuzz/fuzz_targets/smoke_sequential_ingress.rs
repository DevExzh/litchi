mod sequential_ingress_harness;

fn main() {
    let mut bytes = Vec::new();
    litchi_docx::Package::new()
        .unwrap()
        .write_to(&mut bytes)
        .unwrap();
    let mut opened = 0;
    for flags in 0..=255 {
        let mut input = vec![flags];
        input.extend_from_slice(&bytes);
        let accepted = sequential_ingress_harness::exercise(&input);
        assert_eq!(accepted, flags & 12 == 0);
        opened += usize::from(accepted);
    }
    for bytes in [&b""[..], &b"PK"[..], &b"not a ZIP"[..]] {
        for flags in 0..=255 {
            let mut input = vec![flags];
            input.extend_from_slice(bytes);
            assert!(!sequential_ingress_harness::exercise(&input));
        }
    }
    println!(
        "sequential DOCX smoke passed: valid_modes={opened} rejected_valid_modes={} malformed_modes=768; no retained memory charges",
        256 - opened
    );
}
