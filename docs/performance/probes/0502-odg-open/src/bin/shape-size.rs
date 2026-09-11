use litchi_odg::shape::Shape;

fn main() {
    println!("shape_size_bytes={}", std::mem::size_of::<Shape>());
    println!("shape_align_bytes={}", std::mem::align_of::<Shape>());
}
