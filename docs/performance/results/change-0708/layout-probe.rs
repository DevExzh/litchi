#[allow(dead_code)]
enum ElementName {
    Static(&'static [u8]),
    Owned(Box<[u8]>),
}
fn main() { println!("{} {}", std::mem::size_of::<ElementName>(), std::mem::size_of::<Box<[u8]>>()); }
