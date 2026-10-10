//! Inline spreadsheet coordinate labels shared by format adapters.

/// Short text kept inline in a node, so a view borrows it without
/// allocating: an A1 cell or range, `row N`, or column letters.
#[derive(Clone, Copy)]
pub(crate) struct Label {
    bytes: [u8; 24],
    len: u8,
}

impl Label {
    fn empty() -> Self {
        Self {
            bytes: [0; 24],
            len: 0,
        }
    }

    pub(crate) fn cell(column: u32, row: u32) -> Self {
        let mut label = Self::empty();
        label.push_cell(column, row);
        label
    }

    pub(crate) fn range(first: (u32, u32), last: (u32, u32)) -> Self {
        let mut label = Self::empty();
        label.push_cell(first.0, first.1);
        if last != first {
            label.push(b':');
            label.push_cell(last.0, last.1);
        }
        label
    }

    pub(crate) fn row(index: u32) -> Self {
        let mut label = Self::empty();
        for &byte in b"row " {
            label.push(byte);
        }
        label.push_number(u64::from(index) + 1);
        label
    }

    pub(crate) fn column(index: u32) -> Self {
        let mut label = Self::empty();
        label.push_column(index);
        label
    }

    fn push(&mut self, byte: u8) {
        if let Some(slot) = self.bytes.get_mut(usize::from(self.len)) {
            *slot = byte;
            self.len += 1;
        }
    }

    pub(crate) fn push_column(&mut self, column: u32) {
        let mut letters = [0_u8; 7];
        let mut start = letters.len();
        let mut number = u64::from(column) + 1;
        while number > 0 && start > 0 {
            start -= 1;
            // `(number - 1) % 26` is below 26.
            letters[start] = b'A' + ((number - 1) % 26) as u8;
            number = (number - 1) / 26;
        }
        for &letter in &letters[start..] {
            self.push(letter);
        }
    }

    fn push_number(&mut self, mut number: u64) {
        let mut digits = [0_u8; 20];
        let mut start = digits.len();
        loop {
            start -= 1;
            // `number % 10` is a single digit.
            digits[start] = b'0' + (number % 10) as u8;
            number /= 10;
            if number == 0 {
                break;
            }
        }
        for &digit in &digits[start..] {
            self.push(digit);
        }
    }

    fn push_cell(&mut self, column: u32, row: u32) {
        self.push_column(column);
        self.push_number(u64::from(row) + 1);
    }

    pub(crate) fn as_str(&self) -> &str {
        // Letters, digits, spaces and colons are always UTF-8.
        std::str::from_utf8(&self.bytes[..usize::from(self.len)]).unwrap_or_default()
    }
}
