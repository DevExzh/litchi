mod paths_harness;
fn main() {
    let seeds = [
        "/body/p[1]",
        "Sheet1!b:b",
        "'It''s'!A1",
        "'My Data'!3:3",
        "/body/p[@paraId=ABC]",
        "Sheet1!0:0",
        "PRIVATE!",
        "é!A1",
        "'bad\nname'!B:B",
    ];
    let (mut inputs, mut accepted) = (0, 0);
    for seed in seeds {
        let mut data = seed.as_bytes().to_vec();
        for length in 0..=data.len() {
            inputs += 1;
            accepted += usize::from(paths_harness::exercise(&data[..length]));
        }
        for index in 0..data.len() {
            for mask in [1, 2, 4, 8, 16, 32, 64, 128] {
                data[index] ^= mask;
                inputs += 1;
                accepted += usize::from(paths_harness::exercise(&data));
                data[index] ^= mask;
            }
        }
    }
    println!("path smoke: {inputs} inputs, {accepted} accepted");
}
