/// Shape anchor (position and size).
#[derive(Debug, Clone, Copy)]
pub struct Anchor {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

impl Anchor {
    #[inline]
    pub const fn new(left: i32, top: i32, right: i32, bottom: i32) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    #[inline]
    pub const fn width(&self) -> Option<i32> {
        self.right.checked_sub(self.left)
    }

    #[inline]
    pub const fn height(&self) -> Option<i32> {
        self.bottom.checked_sub(self.top)
    }
}
