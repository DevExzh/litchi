use std::alloc::{GlobalAlloc, Layout, System};
fn main() {
    let layout = Layout::from_size_align(isize::MAX as usize, 1).unwrap();
    let pointer = unsafe { System.alloc(layout) };
    println!("pointer={pointer:p} null={}", pointer.is_null());
    if !pointer.is_null() {
        unsafe { System.dealloc(pointer, layout) };
    }
}
