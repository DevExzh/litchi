mod direct_properties_harness;
fn main() {
    let seeds = [
        br#"<w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>x</w:t></w:r>"#.as_slice(),
        br#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#,
        br#"<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr></w:tbl>"#,
        br#"<w:tr><w:trPr><w:trHeight w:val="240"/></w:trPr></w:tr>"#,
        br#"<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc>"#,
    ];
    let mut cases = 0;
    let mut accepted = 0;
    for (selector, seed) in seeds.iter().enumerate() {
        let mut input = vec![selector as u8];
        input.extend_from_slice(seed);
        assert!(direct_properties_harness::exercise(&input));
        cases += 1;
        accepted += 1;
        for length in 0..seed.len() {
            cases += 1;
            accepted += usize::from(direct_properties_harness::exercise(&input[..=length]));
        }
        for position in 1..input.len() {
            for mask in [1, 2, 4, 8, 16, 32, 64, 128] {
                input[position] ^= mask;
                cases += 1;
                accepted += usize::from(direct_properties_harness::exercise(&input));
                input[position] ^= mask;
            }
        }
    }
    println!(
        "direct properties smoke: {cases} cases, {accepted} accepted, {} refused",
        cases - accepted
    );
}
