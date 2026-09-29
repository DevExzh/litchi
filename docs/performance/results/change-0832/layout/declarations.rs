#![allow(dead_code)]

mod before {
use bitflags::bitflags;
use litchi_xlsx::{ColumnIndex as Index, Width, Outline};

bitflags! {
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub(crate) struct Flags: u8 {
        const HIDDEN = 1 << 0;
        const BEST_FIT = 1 << 1;
        const CUSTOM_WIDTH = 1 << 2;
        const PHONETIC = 1 << 3;
        const COLLAPSED = 1 << 4;
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Properties {
    pub(crate) width: Option<Width>,
    pub(crate) style: Option<u32>,
    pub(crate) outline: Outline,
    pub(crate) flags: Flags,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Node<T> {
    Unset,
    Value(T),
    Split,
}
#[derive(Debug)]
pub(crate) struct Assignments<T> {
    nodes: Vec<Node<T>>,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Assigned<T> {
    pub(crate) first: Index,
    pub(crate) last: Index,
    pub(crate) value: T,
}
pub(super) fn print() {

println!("before\tProperties\t{}\t{}", std::mem::size_of::<Properties>(), std::mem::align_of::<Properties>());

println!("before\tNode<Properties>\t{}\t{}", std::mem::size_of::<Node<Properties>>(), std::mem::align_of::<Node<Properties>>());

println!("before\tAssigned<Properties>\t{}\t{}", std::mem::size_of::<Assigned<Properties>>(), std::mem::align_of::<Assigned<Properties>>());

println!("before\tAssignments<Properties>\t{}\t{}", std::mem::size_of::<Assignments<Properties>>(), std::mem::align_of::<Assignments<Properties>>());

println!("before\tNode<usize>\t{}\t{}", std::mem::size_of::<Node<usize>>(), std::mem::align_of::<Node<usize>>());

println!("before\tAssigned<usize>\t{}\t{}", std::mem::size_of::<Assigned<usize>>(), std::mem::align_of::<Assigned<usize>>());

println!("before\tAssignments<usize>\t{}\t{}", std::mem::size_of::<Assignments<usize>>(), std::mem::align_of::<Assignments<usize>>());

}
}

mod after {
use bitflags::bitflags;
use litchi_xlsx::{ColumnIndex as Index, Width, Outline};

bitflags! {
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub(crate) struct Flags: u8 {
        const HIDDEN = 1 << 0;
        const BEST_FIT = 1 << 1;
        const CUSTOM_WIDTH = 1 << 2;
        const PHONETIC = 1 << 3;
        const COLLAPSED = 1 << 4;
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Properties {
    pub(crate) width: Option<Width>,
    pub(crate) style: Option<u32>,
    pub(crate) outline: Outline,
    pub(crate) flags: Flags,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Node<T> {
    Unset,
    Value(T),
    Split,
}
#[derive(Debug)]
pub(crate) struct Assignments<T> {
    nodes: Vec<Node<T>>,
    first: Option<Assigned<T>>,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Assigned<T> {
    pub(crate) first: Index,
    pub(crate) last: Index,
    pub(crate) value: T,
}
pub(super) fn print() {

println!("after\tProperties\t{}\t{}", std::mem::size_of::<Properties>(), std::mem::align_of::<Properties>());

println!("after\tNode<Properties>\t{}\t{}", std::mem::size_of::<Node<Properties>>(), std::mem::align_of::<Node<Properties>>());

println!("after\tAssigned<Properties>\t{}\t{}", std::mem::size_of::<Assigned<Properties>>(), std::mem::align_of::<Assigned<Properties>>());

println!("after\tAssignments<Properties>\t{}\t{}", std::mem::size_of::<Assignments<Properties>>(), std::mem::align_of::<Assignments<Properties>>());

println!("after\tNode<usize>\t{}\t{}", std::mem::size_of::<Node<usize>>(), std::mem::align_of::<Node<usize>>());

println!("after\tAssigned<usize>\t{}\t{}", std::mem::size_of::<Assigned<usize>>(), std::mem::align_of::<Assigned<usize>>());

println!("after\tAssignments<usize>\t{}\t{}", std::mem::size_of::<Assignments<usize>>(), std::mem::align_of::<Assignments<usize>>());

}
}

fn main() { before::print(); after::print(); }
