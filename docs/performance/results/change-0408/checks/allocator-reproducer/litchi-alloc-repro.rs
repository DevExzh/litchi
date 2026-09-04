use std::alloc::{GlobalAlloc, Layout, System};
fn main() {
    let size = std::hint::black_box(isize::MAX as usize);
    let layout = Layout::from_size_align(size, 1).unwrap();
    let pointer = unsafe { System.alloc(layout) };
    println!("size={size} pointer={pointer:p} null={}", pointer.is_null());
    if !pointer.is_null() {
        unsafe { System.dealloc(pointer, layout) };
    }
}
