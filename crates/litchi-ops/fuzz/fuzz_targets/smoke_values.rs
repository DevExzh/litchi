mod values_harness;
fn main() {
    let seeds = [
        "1in",
        "4294967295",
        "0.5pt",
        "accent1+lumMod75",
        "rgb(255,0,1)",
        "#01234567",
        "true",
        "1,234.5",
        "-1.25e10",
        "PRIVATE",
        "\u{1f34d}",
    ];
    let mut cases = 0;
    let mut accepted = 0;
    for seed in seeds {
        let mut data = seed.as_bytes().to_vec();
        cases += 1;
        accepted += values_harness::exercise(&data);
        for length in 0..data.len() {
            cases += 1;
            accepted += values_harness::exercise(&data[..length]);
        }
        for index in 0..data.len() {
            for mask in [1, 2, 4, 8, 16, 32, 64, 128] {
                data[index] ^= mask;
                cases += 1;
                accepted += values_harness::exercise(&data);
                data[index] ^= mask;
            }
        }
    }
    println!("value smoke: {cases} inputs, {accepted} successful coercions");
}
