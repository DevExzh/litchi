use crate::sheet::CellValue;

use super::{to_number, to_text};

pub(crate) enum CriteriaOperator {
    Eq,
    Ne,
    Gt,
    Ge,
    Lt,
    Le,
}

pub(crate) enum CriteriaValue {
    Number(f64),
    Text(String),
}

pub(crate) struct Criteria {
    pub(crate) op: CriteriaOperator,
    pub(crate) rhs: CriteriaValue,
}

pub(crate) fn parse_criteria(spec: &str) -> Option<Criteria> {
    let s = spec.trim();
    let (op, rest) = if s.len() >= 2 {
        let prefix = &s[..2];
        match prefix {
            ">=" => (CriteriaOperator::Ge, &s[2..]),
            "<=" => (CriteriaOperator::Le, &s[2..]),
            "<>" => (CriteriaOperator::Ne, &s[2..]),
            _ => {
                let first = s.chars().next().unwrap();
                match first {
                    '>' => (CriteriaOperator::Gt, &s[1..]),
                    '<' => (CriteriaOperator::Lt, &s[1..]),
                    '=' => (CriteriaOperator::Eq, &s[1..]),
                    _ => (CriteriaOperator::Eq, s),
                }
            },
        }
    } else if let Some(first) = s.chars().next() {
        match first {
            '>' => (CriteriaOperator::Gt, &s[1..]),
            '<' => (CriteriaOperator::Lt, &s[1..]),
            '=' => (CriteriaOperator::Eq, &s[1..]),
            _ => (CriteriaOperator::Eq, s),
        }
    } else {
        (CriteriaOperator::Eq, s)
    };

    let rhs_str = rest.trim();
    if rhs_str.is_empty() {
        return Some(Criteria {
            op,
            rhs: CriteriaValue::Text(String::new()),
        });
    }

    if let Ok(n) = rhs_str.parse::<f64>() {
        Some(Criteria {
            op,
            rhs: CriteriaValue::Number(n),
        })
    } else {
        Some(Criteria {
            op,
            rhs: CriteriaValue::Text(rhs_str.to_string()),
        })
    }
}

fn wildcard_match(pattern: &str, text: &str) -> bool {
    let p: Vec<char> = pattern.chars().collect();
    let t: Vec<char> = text.chars().collect();
    let m = p.len();
    let n = t.len();
    let mut dp = vec![vec![false; n + 1]; m + 1];
    dp[0][0] = true;

    for i in 1..=m {
        if p[i - 1] == '*' {
            dp[i][0] = dp[i - 1][0];
        }
    }

    for i in 1..=m {
        for j in 1..=n {
            match p[i - 1] {
                '*' => {
                    dp[i][j] = dp[i - 1][j] || dp[i][j - 1];
                },
                '?' => {
                    dp[i][j] = dp[i - 1][j - 1];
                },
                c => {
                    dp[i][j] = dp[i - 1][j - 1] && c == t[j - 1];
                },
            }
        }
    }

    dp[m][n]
}

pub(crate) fn matches_criteria(value: &CellValue, criteria: &Criteria) -> bool {
    match &criteria.rhs {
        CriteriaValue::Number(target) => {
            if let Some(v) = to_number(value) {
                match criteria.op {
                    CriteriaOperator::Eq => v == *target,
                    CriteriaOperator::Ne => v != *target,
                    CriteriaOperator::Gt => v > *target,
                    CriteriaOperator::Ge => v >= *target,
                    CriteriaOperator::Lt => v < *target,
                    CriteriaOperator::Le => v <= *target,
                }
            } else {
                false
            }
        },
        CriteriaValue::Text(pattern) => {
            let text = to_text(value);
            let has_wildcard = pattern.contains('*') || pattern.contains('?');
            match criteria.op {
                CriteriaOperator::Eq => {
                    if has_wildcard {
                        wildcard_match(pattern, &text)
                    } else {
                        text == *pattern
                    }
                },
                CriteriaOperator::Ne => {
                    if has_wildcard {
                        !wildcard_match(pattern, &text)
                    } else {
                        text != *pattern
                    }
                },
                CriteriaOperator::Gt => text > *pattern,
                CriteriaOperator::Ge => text >= *pattern,
                CriteriaOperator::Lt => text < *pattern,
                CriteriaOperator::Le => text <= *pattern,
            }
        },
    }
}
