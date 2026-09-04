use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};
static FAILED: AtomicUsize = AtomicUsize::new(0);
static SUCCESS: AtomicUsize = AtomicUsize::new(0);
struct A;
unsafe impl GlobalAlloc for A {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let p = unsafe { System.alloc(layout) };
        if p.is_null() { FAILED.fetch_add(1, Ordering::SeqCst); }
        else { SUCCESS.fetch_add(layout.size(), Ordering::SeqCst); }
        p
    }
    unsafe fn dealloc(&self, p: *mut u8, l: Layout) { unsafe { System.dealloc(p, l) }; }
}
fn main() {
    let l = Layout::from_size_align(isize::MAX as usize, 1).unwrap();
    let p = unsafe { A.alloc(l) };
    println!("p={p:p} null={} failed={} success={}", p.is_null(), FAILED.load(Ordering::SeqCst), SUCCESS.load(Ordering::SeqCst));
}
