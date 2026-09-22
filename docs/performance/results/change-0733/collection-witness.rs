#![forbid(unsafe_code)]

#[inline(never)]
fn leaf() -> u64 {
    let mut sum = 0_u64;
    for value in 0..1000_u64 {
        sum = sum.wrapping_add(std::hint::black_box(value));
    }
    std::hint::black_box(sum)
}

#[inline(never)]
fn work() -> u64 {
    leaf()
}

#[inline(never)]
fn owner() -> u64 {
    work()
}

fn main() {
    let before = work();
    let collected = owner();
    let after = work();
    assert_eq!([before, collected, after], [499_500; 3]);
    println!("{{\"results\":[{before},{collected},{after}]}}");
}
