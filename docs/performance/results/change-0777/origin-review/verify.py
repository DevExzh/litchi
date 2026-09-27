#!/usr/bin/env python3
"""Mechanical verification of `checked_attributes()` call sites (record 0770 review).

For each `checked` site: tokenize the Rust file (comments/strings/chars/lifetimes aware),
locate the `checked_attributes` token on the line, and accept automatically only the shape

    [label:] for VAR in <receiver>.checked_attributes() {
        <statements that neither mention VAR nor contain `continue`>
        let PAT = VAR?;                          |
        let PAT = VAR.map_err(<balanced>)?[<.field/.method(...) chain>];  |
        let PAT = VAR.ok()?[...];
        ...
    }

where the VAR token of the first mention is at brace/paren depth 0 of the loop body, the
statement starts at a statement boundary, and nothing (no closure bar) sits between `=` and VAR.
Everything else is written out for hand review.
"""
import json
import re
import sys

W = '/home/zhuhe/code/litchi-worktrees/0770-quick-xml-fail-fast-attribute-checks/'
SITES = '/home/zhuhe/code/litchi-worktrees/scratch/0770/review-sites.tsv'


def lex(text):
    """Return a list of (kind, value, start, end). Comments are dropped."""
    toks = []
    i = 0
    n = len(text)
    ident_re = re.compile(r'[A-Za-z_][A-Za-z0-9_]*')
    num_re = re.compile(r'[0-9][0-9A-Za-z_]*(\.[0-9][0-9A-Za-z_]*)?')
    while i < n:
        c = text[i]
        if c in ' \t\r\n':
            i += 1
            continue
        if text.startswith('//', i):
            j = text.find('\n', i)
            i = n if j < 0 else j
            continue
        if text.startswith('/*', i):
            depth = 0
            j = i
            while j < n:
                if text.startswith('/*', j):
                    depth += 1
                    j += 2
                elif text.startswith('*/', j):
                    depth -= 1
                    j += 2
                    if depth == 0:
                        break
                else:
                    j += 1
            i = j
            continue
        # raw strings: r"..", r#".."#, br"..", br#".."#, cr".."
        m = re.match(r'(b|c)?r(#*)"', text[i:i + 300])
        if m and (i == 0 or not (text[i - 1].isalnum() or text[i - 1] == '_')):
            hashes = m.group(2)
            close = '"' + hashes
            j = text.find(close, i + m.end())
            assert j >= 0
            toks.append(('str', text[i:j + len(close)], i, j + len(close)))
            i = j + len(close)
            continue
        # strings: "..", b"..", c".."
        if c == '"' or (c in 'bc' and i + 1 < n and text[i + 1] == '"' and
                        (i == 0 or not (text[i - 1].isalnum() or text[i - 1] == '_'))):
            j = i + (1 if c == '"' else 2)
            while text[j] != '"':
                if text[j] == '\\':
                    j += 1
                j += 1
            toks.append(('str', text[i:j + 1], i, j + 1))
            i = j + 1
            continue
        # byte char b'x'
        if c == 'b' and i + 1 < n and text[i + 1] == "'" and (i == 0 or not (text[i - 1].isalnum() or text[i - 1] == '_')):
            j = i + 2
            if text[j] == '\\':
                j += 2
                while text[j] != "'":
                    j += 1
            else:
                j += 1
            assert text[j] == "'", (text[i:i + 10])
            toks.append(('char', text[i:j + 1], i, j + 1))
            i = j + 1
            continue
        if c == "'":
            # char literal or lifetime/label
            if i + 2 < n and text[i + 1] == '\\':
                j = i + 3
                while text[j] != "'":
                    j += 1
                toks.append(('char', text[i:j + 1], i, j + 1))
                i = j + 1
                continue
            # one (possibly multi-byte) char followed by '
            if i + 2 < n and text[i + 2] == "'":
                toks.append(('char', text[i:i + 3], i, i + 3))
                i += 3
                continue
            m = ident_re.match(text, i + 1)
            if m:
                toks.append(('lifetime', text[i:m.end()], i, m.end()))
                i = m.end()
                continue
            toks.append(('punct', c, i, i + 1))
            i += 1
            continue
        m = ident_re.match(text, i)
        if m:
            toks.append(('ident', m.group(0), i, m.end()))
            i = m.end()
            continue
        m = num_re.match(text, i)
        if m:
            toks.append(('num', m.group(0), i, m.end()))
            i = m.end()
            continue
        toks.append(('punct', c, i, i + 1))
        i += 1
    return toks


OPEN = {'(': ')', '[': ']', '{': '}'}
CLOSE = {')', ']', '}'}


def match_close(toks, k):
    """toks[k] is an opener; return index of its matching closer."""
    depth = 0
    for j in range(k, len(toks)):
        v = toks[j][1]
        if toks[j][0] != 'punct':
            continue
        if v in OPEN:
            depth += 1
        elif v in CLOSE:
            depth -= 1
            if depth == 0:
                return j
    raise ValueError('unbalanced')


def match_open(toks, k):
    """toks[k] is a closer; return index of its matching opener."""
    depth = 0
    for j in range(k, -1, -1):
        v = toks[j][1]
        if toks[j][0] != 'punct':
            continue
        if v in CLOSE:
            depth += 1
        elif v in OPEN:
            depth -= 1
            if depth == 0:
                return j
    raise ValueError('unbalanced')


def line_starts(text):
    starts = [0]
    for idx, ch in enumerate(text):
        if ch == '\n':
            starts.append(idx + 1)
    return starts


def analyse(path, line, cache):
    if path not in cache:
        text = open(W + path).read()
        cache[path] = (text, lex(text), line_starts(text))
    text, toks, starts = cache[path]
    lo = starts[line - 1]
    hi = starts[line] if line < len(starts) else len(text)
    idxs = [k for k, t in enumerate(toks) if t[0] == 'ident' and t[1] == 'checked_attributes' and lo <= t[2] < hi]
    if len(idxs) != 1:
        return ('HAND', f'{len(idxs)} checked_attributes tokens on line')
    k = idxs[0]
    if toks[k - 1][1] != '.' or toks[k + 1][1] != '(' or toks[k + 2][1] != ')':
        return ('HAND', 'not a .checked_attributes() method call')
    after = toks[k + 3]
    if after[1] != '{':
        return ('HAND', f'iterator followed by {after[1]!r} (adaptor or not a for header)')
    # walk back to `in` at depth 0 and the `for` before it
    j = k - 2
    while j >= 0:
        t = toks[j]
        if t[0] == 'punct' and t[1] in CLOSE:
            j = match_open(toks, j) - 1
            continue
        if t[0] == 'ident' and t[1] == 'in':
            break
        if t[0] == 'punct' and t[1] in (';', '{', '}', '=', ','):
            return ('HAND', 'no `in` before receiver')
        if t[0] == 'ident' and t[1] in ('let', 'for', 'match', 'if', 'while', 'return'):
            return ('HAND', f'hit {t[1]} before `in`')
        j -= 1
    in_idx = j
    receiver = text[toks[in_idx + 1][2]:toks[k - 1][2]]
    if not re.fullmatch(r'[&*(]*[A-Za-z_][A-Za-z0-9_]*(\.[A-Za-z_][A-Za-z0-9_]*)*\)?', receiver.strip()):
        return ('HAND', f'receiver {receiver!r}')
    # pattern between `for` and `in`
    if toks[in_idx - 2][1] != 'for' or toks[in_idx - 1][0] != 'ident':
        return ('HAND', 'for pattern is not a single identifier')
    var = toks[in_idx - 1][1]
    if var in ('mut', '_'):
        return ('HAND', 'for pattern mut/_')
    body_open = k + 3
    body_close = match_close(toks, body_open)
    body = toks[body_open + 1:body_close]
    # first mention of var anywhere in body
    first = None
    for bi, t in enumerate(body):
        if t[0] == 'ident' and t[1] == var:
            first = bi
            break
    if first is None:
        return ('HAND', 'loop variable never mentioned')
    pre = body[:first]
    if any(t[0] == 'ident' and t[1] == 'continue' for t in pre):
        return ('HAND', '`continue` before first mention')
    # depth of first mention
    depth = 0
    stmt_start = 0
    for bi in range(first):
        v = body[bi][1]
        if body[bi][0] != 'punct':
            continue
        if v in OPEN:
            depth += 1
        elif v in CLOSE:
            depth -= 1
            if depth == 0 and v == '}':
                stmt_start = bi + 1
        elif v == ';' and depth == 0:
            stmt_start = bi + 1
    if depth != 0:
        return ('HAND', 'first mention is nested (inside (), [] or {})')
    if body[stmt_start][1] != 'let':
        return ('HAND', 'first mention statement is not `let ... = VAR...`: ' + ' '.join(t[1] for t in body[stmt_start:stmt_start + 12]))
    # find `=` at depth 0 after `let`
    d = 0
    eq = None
    for bi in range(stmt_start + 1, len(body)):
        v = body[bi][1]
        if body[bi][0] == 'punct':
            if v in OPEN:
                d += 1
            elif v in CLOSE:
                d -= 1
            elif v == '=' and d == 0 and body[bi + 1][1] not in ('=', '>') and body[bi - 1][1] not in ('=', '!', '<', '>'):
                eq = bi
                break
            elif v == ';' and d == 0:
                break
    if eq is None:
        return ('HAND', 'let without =')
    pat = body[stmt_start + 1:eq]
    if any(t[1] in ('|', 'else') for t in pat):
        return ('HAND', 'odd let pattern')
    if eq < first and not (body[eq + 1][0] == 'ident' and body[eq + 1][1] == var):
        return ('HAND', 'RHS does not start with VAR: ' + ' '.join(t[1] for t in body[eq + 1:eq + 8]))
    if eq > first:
        # first mention is inside the pattern (shadowing): the RHS must start with VAR
        if not (body[eq + 1][0] == 'ident' and body[eq + 1][1] == var):
            return ('HAND', 'pattern shadows VAR but RHS does not start with VAR: ' + ' '.join(t[1] for t in body[eq + 1:eq + 8]))
        # make sure the pattern mention is just a binding (let VAR / let mut VAR / let VAR: T)
    first = eq + 1
    # after VAR
    rest = body[first + 1:]
    r = 0
    shape = None
    if rest[r][1] == '?':
        shape = 'VAR?'
        r += 1
    elif rest[r][1] == '.' and rest[r + 1][1] == 'map_err' and rest[r + 2][1] == '(':
        close = match_close(rest, r + 2)
        if rest[close + 1][1] != '?':
            return ('HAND', 'map_err not followed by ?')
        # map_err argument must not mention var (paranoia)
        shape = 'VAR.map_err(..)?'
        r = close + 2
    elif rest[r][1] == '.' and rest[r + 1][1] == 'ok' and rest[r + 2][1] == '(' and rest[r + 3][1] == ')' and rest[r + 4][1] == '?':
        shape = 'VAR.ok()?'
        r += 5
    else:
        return ('HAND', 'VAR not followed by ? / .map_err(..)? / .ok()?: ' + ' '.join(t[1] for t in rest[:8]))
    # remainder of the statement up to `;` at depth 0 (a trailing field/method chain is fine)
    d = 0
    tail = []
    while r < len(rest):
        v = rest[r][1]
        if rest[r][0] == 'punct':
            if v in OPEN:
                d += 1
            elif v in CLOSE:
                d -= 1
            elif v == ';' and d == 0:
                break
        if d == 0 and v in ('else',):
            return ('HAND', 'let-else after ?')
        tail.append(v)
        r += 1
    else:
        return ('HAND', 'statement not terminated by ;')
    if tail and tail[0] not in ('.',):
        return ('HAND', 'unexpected tail after ?: ' + ' '.join(tail[:8]))
    pre_txt = ' '.join(t[1] for t in body[:stmt_start])
    if any(t[0] == 'ident' and t[1] == 'continue' for t in body[:stmt_start]):
        return ('HAND', 'continue before the let statement')
    return ('AUTO', f'{shape}; tail={"".join(tail)[:60]!r}; pre-stmts={len(pre)} toks', pre_txt)


def main():
    cache = {}
    out = []
    for raw in open(SITES):
        loc, kind, src = raw.rstrip('\n').split('\t', 2)
        if kind != 'checked':
            continue
        path, line = loc.rsplit(':', 1)
        res = analyse(path, int(line), cache)
        out.append((loc, res))
    auto = [o for o in out if o[1][0] == 'AUTO']
    hand = [o for o in out if o[1][0] != 'AUTO']
    print('auto', len(auto), 'hand', len(hand))
    json.dump(out, open('/home/zhuhe/code/litchi-worktrees/scratch/0770/review/auto.json', 'w'), indent=1)
    for loc, res in hand:
        print('HAND', loc, res[1])


if __name__ == '__main__':
    main()
