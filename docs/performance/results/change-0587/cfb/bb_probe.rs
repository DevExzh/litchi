use std::io::{BorrowedBuf, Read};
fn main() {
    let mut backing = [std::mem::MaybeUninit::<u8>::uninit(); 8];
    let mut buf: BorrowedBuf<'_> = (&mut backing[..]).into();
    let mut cursor = buf.unfilled();
    let src: &[u8] = b"abc";
    (&src[..]).read_buf(cursor.reborrow()).unwrap();
    println!("{}", buf.filled().len());
}
