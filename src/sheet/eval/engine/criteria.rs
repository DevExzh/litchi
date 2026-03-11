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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_criteria_empty() {
        let result = parse_criteria("");
        assert!(result.is_some());
        let crit = result.unwrap();
        match crit.rhs {
            CriteriaValue::Text(s) => assert_eq!(s, ""),
            _ => panic!("Expected empty text"),
        }
    }

    #[test]
    fn test_parse_criteria_number() {
        let crit = parse_criteria("42").unwrap();
        match crit.rhs {
            CriteriaValue::Number(n) => assert_eq!(n, 42.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_parse_criteria_text() {
        let crit = parse_criteria("hello").unwrap();
        match crit.rhs {
            CriteriaValue::Text(s) => assert_eq!(s, "hello"),
            _ => panic!("Expected text"),
        }
    }

    #[test]
    fn test_parse_criteria_eq() {
        let crit = parse_criteria("=5").unwrap();
        match crit.op {
            CriteriaOperator::Eq => {},
            _ => panic!("Expected Eq operator"),
        }
        match crit.rhs {
            CriteriaValue::Number(n) => assert_eq!(n, 5.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_parse_criteria_gt() {
        let crit = parse_criteria(">10").unwrap();
        match crit.op {
            CriteriaOperator::Gt => {},
            _ => panic!("Expected Gt operator"),
        }
        match crit.rhs {
            CriteriaValue::Number(n) => assert_eq!(n, 10.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_parse_criteria_lt() {
        let crit = parse_criteria("<5").unwrap();
        match crit.op {
            CriteriaOperator::Lt => {},
            _ => panic!("Expected Lt operator"),
        }
    }

    #[test]
    fn test_parse_criteria_ge() {
        let crit = parse_criteria(">=100").unwrap();
        match crit.op {
            CriteriaOperator::Ge => {},
            _ => panic!("Expected Ge operator"),
        }
        match crit.rhs {
            CriteriaValue::Number(n) => assert_eq!(n, 100.0),
            _ => panic!("Expected number"),
        }
    }

    #[test]
    fn test_parse_criteria_le() {
        let crit = parse_criteria("<=50").unwrap();
        match crit.op {
            CriteriaOperator::Le => {},
            _ => panic!("Expected Le operator"),
        }
    }

    #[test]
    fn test_parse_criteria_ne() {
        let crit = parse_criteria("<>0").unwrap();
        match crit.op {
            CriteriaOperator::Ne => {},
            _ => panic!("Expected Ne operator"),
        }
    }

    #[test]
    fn test_wildcard_match_exact() {
        assert!(wildcard_match("hello", "hello"));
        assert!(!wildcard_match("hello", "world"));
    }

    #[test]
    fn test_wildcard_match_star() {
        assert!(wildcard_match("*", "anything"));
        assert!(wildcard_match("h*o", "hello"));
        assert!(wildcard_match("h*", "hello"));
        assert!(wildcard_match("*o", "hello"));
        assert!(!wildcard_match("h*z", "hello"));
    }

    #[test]
    fn test_wildcard_match_question() {
        assert!(wildcard_match("h?llo", "hello"));
        assert!(wildcard_match("?????", "hello"));
        assert!(!wildcard_match("????", "hello"));
    }

    #[test]
    fn test_wildcard_match_mixed() {
        assert!(wildcard_match("h*l?o", "hello"));
        assert!(wildcard_match("*e?lo", "hello"));
    }

    #[test]
    fn test_matches_criteria_number_eq() {
        let crit = Criteria {
            op: CriteriaOperator::Eq,
            rhs: CriteriaValue::Number(10.0),
        };
        assert!(matches_criteria(&CellValue::Int(10), &crit));
        assert!(matches_criteria(&CellValue::Float(10.0), &crit));
        assert!(!matches_criteria(&CellValue::Int(5), &crit));
    }

    #[test]
    fn test_matches_criteria_number_gt() {
        let crit = Criteria {
            op: CriteriaOperator::Gt,
            rhs: CriteriaValue::Number(10.0),
        };
        assert!(matches_criteria(&CellValue::Int(15), &crit));
        assert!(!matches_criteria(&CellValue::Int(10), &crit));
        assert!(!matches_criteria(&CellValue::Int(5), &crit));
    }

    #[test]
    fn test_matches_criteria_text_eq() {
        let crit = Criteria {
            op: CriteriaOperator::Eq,
            rhs: CriteriaValue::Text("apple".to_string()),
        };
        assert!(matches_criteria(
            &CellValue::String("apple".to_string()),
            &crit
        ));
        assert!(!matches_criteria(
            &CellValue::String("banana".to_string()),
            &crit
        ));
    }

    #[test]
    fn test_matches_criteria_text_wildcard() {
        let crit = Criteria {
            op: CriteriaOperator::Eq,
            rhs: CriteriaValue::Text("a*e".to_string()),
        };
        assert!(matches_criteria(
            &CellValue::String("apple".to_string()),
            &crit
        ));
        assert!(matches_criteria(
            &CellValue::String("axe".to_string()),
            &crit
        ));
        assert!(!matches_criteria(
            &CellValue::String("banana".to_string()),
            &crit
        ));
    }

    #[test]
    fn test_matches_criteria_text_gt() {
        let crit = Criteria {
            op: CriteriaOperator::Gt,
            rhs: CriteriaValue::Text("m".to_string()),
        };
        assert!(matches_criteria(&CellValue::String("z".to_string()), &crit));
        assert!(!matches_criteria(
            &CellValue::String("a".to_string()),
            &crit
        ));
    }

    #[test]
    fn test_matches_criteria_non_number() {
        let crit = Criteria {
            op: CriteriaOperator::Eq,
            rhs: CriteriaValue::Number(10.0),
        };
        assert!(!matches_criteria(
            &CellValue::String("hello".to_string()),
            &crit
        ));
        assert!(!matches_criteria(&CellValue::Bool(true), &crit));
    }
}
