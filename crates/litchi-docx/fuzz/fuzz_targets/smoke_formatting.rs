mod formatting_harness;
fn main() {
    let seeds: [&[u8]; 5] = [
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>"#;
        5
    ];
    let mut cases = 0;
    let mut accepted = 0;
    for (selector, seed) in seeds.iter().enumerate() {
        let mut input = vec![selector as u8];
        input.extend_from_slice(seed);
        assert!(formatting_harness::exercise(&input));
        cases += 1;
        accepted += 1;
        for length in 0..seed.len() {
            cases += 1;
            accepted += usize::from(formatting_harness::exercise(&input[..=length]));
        }
        for position in 1..input.len() {
            for mask in [1, 2, 4, 8, 16, 32, 64, 128] {
                input[position] ^= mask;
                cases += 1;
                accepted += usize::from(formatting_harness::exercise(&input));
                input[position] ^= mask;
            }
        }
    }
    println!(
        "formatting smoke: {cases} cases, {accepted} accepted, {} refused",
        cases - accepted
    );
}
