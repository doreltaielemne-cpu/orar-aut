"""Parse the UTCN Automatica timetable workbook (xlsx) into structured entries.

Model: each timetable sheet is a grid of small cells. A time slot (e.g. 8-10) spans
two rows; a semigroup spans two columns. Activities are text cells whose extent is
given by merges and borders. Alternating (odd/even week) activities are drawn with
diagonal cell borders: text above the diagonal = odd weeks, below = even weeks.
"""
import re, sys, json
from collections import defaultdict, deque
import openpyxl
from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL

TIME_RE = re.compile(r'^\s*(\d{1,2})\s*[:.]?\s*-\s*(\d{1,2})\s*[:.]?\s*$')
GROUP_RE = re.compile(r'^\s*(3\d{4})')


def hidden_sets(ws):
    hc, hr = set(), set()
    for k, d in ws.column_dimensions.items():
        if d.hidden:
            lo = d.min or CI(k); hi = d.max or lo
            for c in range(lo, hi + 1):
                hc.add(c)
    for k, d in ws.row_dimensions.items():
        if d.hidden:
            hr.add(k)
    return hc, hr


class Sheet:
    def __init__(self, ws):
        self.ws = ws
        self.hc, self.hr = hidden_sets(ws)
        self.anchor = {}   # (r,c) -> (r0,c0,r1,c1) merged range
        for m in ws.merged_cells.ranges:
            rng = (m.min_row, m.min_col, m.max_row, m.max_col)
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1):
                    self.anchor[(r, c)] = rng

    def rng(self, r, c):
        return self.anchor.get((r, c), (r, c, r, c))

    def val(self, r, c):
        r0, c0, _, _ = self.rng(r, c)
        v = self.ws.cell(r0, c0).value
        if v is None:
            return ''
        if isinstance(v, str) and v.startswith('='):
            return ''
        return str(v).strip() if not isinstance(v, float) else str(int(v)) if v == int(v) else str(v)

    def cell(self, r, c):
        return self.ws.cell(r, c)

    def diag(self, r, c):
        """'/' , '\\' or None for the cell (merged ranges: take anchor's border)."""
        r0, c0, _, _ = self.rng(r, c)
        b = self.ws.cell(r0, c0).border
        if b.diagonal is not None and b.diagonal.style:
            if b.diagonalUp:
                return '/'
            if b.diagonalDown:
                return '\\'
        return None

    def edge_right(self, r, c, c2):
        """border between (r,c) and (r,c2) where c2 is next visible column."""
        if self.rng(r, c) == self.rng(r, c2) and (r, c) in self.anchor:
            return False
        b1 = self.ws.cell(r, c).border.right
        b2 = self.ws.cell(r, c2).border.left
        s = bool((b1 and b1.style) or (b2 and b2.style))
        # borders set on hidden columns in between
        for cc in range(c + 1, c2):
            bb = self.ws.cell(r, cc).border
            if (bb.left and bb.left.style) or (bb.right and bb.right.style):
                s = True
        return s

    def edge_down(self, r, c, r2):
        if self.rng(r, c) == self.rng(r2, c) and (r, c) in self.anchor:
            return False
        b1 = self.ws.cell(r, c).border.bottom
        b2 = self.ws.cell(r2, c).border.top
        return bool((b1 and b1.style) or (b2 and b2.style))


ROOM_TOK = re.compile(r'^(?:s?\d[\d.]*[A-Za-z]?|[A-Z]{1,2}\s?\.?\d{1,3}[A-Z]?|BT|Bt|bt|\d\.\d{2}|Amf\.?|sala|OBS)$')


def ROOM_ONLY(text):
    s = ' '.join(text.split())
    room = r'(?:[A-Za-z]{1,2}\.?\s?\d{2,3}[A-Z]?|\d{3}[A-Z]?|[A-Za-z]\d|s\s?\d\.\d|BT\s?\d\.\d{2}|\d\.\d{2})'
    if re.fullmatch(room + r'(?:\s+(?:OBS|Obs|obs))?', s):
        return True
    toks = s.split()
    return bool(toks) and len(toks) <= 3 and all(ROOM_TOK.match(t) for t in toks)


def CONT(text):
    s = ' '.join(text.split()).lower()
    return ROOM_ONLY(text) or bool(re.match(r'^(s[aă]?pt|spt|st\.|in\s+s[aă]?pt|intre|\d{1,2}\s*[,.(]|\d{1,2}\s+\d|\(|\d{1,2}$)', s))


def parse_sheet(ws, sheet_key):
    S = Sheet(ws)
    maxr, maxc = ws.max_row, ws.max_column

    def fill_of(r, c):
        f = ws.cell(r, c).fill
        if f is None or f.fill_type != 'solid' or f.fgColor is None:
            return None
        col = f.fgColor.rgb if isinstance(f.fgColor.rgb, str) else ('theme%s_%s' % (f.fgColor.theme, f.fgColor.tint) if f.fgColor.type == 'theme' else None)
        if col in (None, 'FFFFFFFF', '00FFFFFF', 'theme0_0.0'):
            return None
        return col
    vis_cols = [c for c in range(1, maxc + 1) if c not in S.hc]
    vis_rows = [r for r in range(1, maxr + 1) if r not in S.hr]

    # ---- group header row
    best = None
    for r in range(1, 12):
        hits = [(c, S.val(r, c)) for c in vis_cols if GROUP_RE.match(S.val(r, c)) and S.rng(r, c)[:2] == (r, c)]
        if len(hits) >= 2 and (best is None or len(hits) > len(best[1])):
            best = (r, hits)
    grow, ghits = best
    pa_max = None
    try:
        pa = ws.print_area
        if pa:
            last = pa.split(':')[-1].replace('$', '')
            pa_max = CI(re.match(r'[A-Z]+', last).group(0))
    except Exception:
        pa_max = None
    if pa_max:
        ghits = [(c, t) for c, t in ghits if c <= pa_max]
    groups = []
    for c, txt in ghits:
        r0, c0, r1, c1 = S.rng(grow, c)
        cols = [x for x in range(c0, c1 + 1) if x not in S.hc]
        m = re.search(r'\(\s*(\d)\s*semigrup', txt)
        groups.append(dict(code=GROUP_RE.match(txt).group(1), label=txt, cols=cols, nsg=int(m.group(1)) if m else None, row=r1))
    # duplicate labels
    seen = defaultdict(int)
    for g in groups:
        seen[g['code']] += 1
        if seen[g['code']] > 1:
            g['code'] = g['code'] + '#' + str(seen[g['code']])
    # ---- semigroup columns
    sgcols = []   # (group code, sg, [cols])
    for g in groups:
        srow = g['row'] + 1
        labs = []
        for c in g['cols']:
            v = S.val(srow, c)
            if v in ('1', '2') and S.rng(srow, c)[1] == c:
                r0, c0, r1, c1 = S.rng(srow, c)
                labs.append((int(v), [x for x in range(c0, c1 + 1) if x not in S.hc and x in g['cols']]))
        if labs:
            # a label row may give "1","1" (single sg split in two sub-columns) -> merge same numbers
            by = defaultdict(list)
            for n, cs in labs:
                by[n] += cs
            # columns of the group not covered by any label -> attach to nearest label
            covered = {c for cs in by.values() for c in cs}
            for c in g['cols']:
                if c not in covered:
                    near = min(by, key=lambda n: min(abs(c - x) for x in by[n]))
                    by[near].append(c)
            for n in sorted(by):
                sgcols.append((g['code'], n, sorted(by[n])))
        else:
            cs = g['cols']
            n = g['nsg'] or (2 if len(cs) == 4 else 1)
            if n == 2 and len(cs) >= 2:
                h = len(cs) // 2
                sgcols.append((g['code'], 1, cs[:h])); sgcols.append((g['code'], 2, cs[h:]))
            else:
                sgcols.append((g['code'], 1, cs))
    grid_cols = sorted({c for _, _, cs in sgcols for c in cs})
    col_owner = {}
    for code, n, cs in sgcols:
        for c in cs:
            col_owner[c] = (code, n)

    # ---- time rows: find the time-label column (left of grid)
    tcol = None
    for c in range(1, grid_cols[0]):
        n = sum(1 for r in vis_rows if TIME_RE.match(S.val(r, c)) and S.rng(r, c)[:2] == (r, c))
        if n >= 10:
            tcol = c
    slots = []  # dict(day, s, e, rows)
    day = -1; prev = 99
    for r in vis_rows:
        if r <= grow:
            continue
        v = S.val(r, tcol)
        m = TIME_RE.match(v)
        if not m or S.rng(r, tcol)[:2] != (r, tcol):
            continue
        s, e = int(m.group(1)), int(m.group(2))
        if s < prev:
            day += 1
        prev = s
        r0, _, r1, _ = S.rng(r, tcol)
        rows = [x for x in range(r0, r1 + 1) if x not in S.hr]
        slots.append(dict(day=day, s=s, e=e, rows=rows))
    row_time = {}
    for sl in slots:
        n = len(sl['rows'])
        for i, r in enumerate(sl['rows']):
            h0 = sl['s'] + (sl['e'] - sl['s']) * i / n
            h1 = sl['s'] + (sl['e'] - sl['s']) * (i + 1) / n
            row_time[r] = (sl['day'], h0, h1)
    grid_rows = sorted(row_time)
    rowset = set(grid_rows); colset = set(grid_cols)
    nextc = {grid_cols[i]: grid_cols[i + 1] for i in range(len(grid_cols) - 1)}
    nextr = {grid_rows[i]: grid_rows[i + 1] for i in range(len(grid_rows) - 1)}

    # ---- flood regions (connected sub-cells without borders between them)
    comp = {}
    cid = 0
    for r in grid_rows:
        for c in grid_cols:
            if (r, c) in comp:
                continue
            cid += 1
            q = deque([(r, c)]); comp[(r, c)] = cid
            while q:
                a, b = q.popleft()
                nbrs = []
                if b in nextc and not S.edge_right(a, b, nextc[b]):
                    nbrs.append((a, nextc[b]))
                pb = [x for x in grid_cols if nextc.get(x) == b]
                if pb and not S.edge_right(a, pb[0], b):
                    nbrs.append((a, pb[0]))
                if a in nextr and row_time[a][0] == row_time[nextr[a]][0] and not S.edge_down(a, b, nextr[a]):
                    nbrs.append((nextr[a], b))
                pa = [x for x in grid_rows if nextr.get(x) == a]
                if pa and row_time[pa[0]][0] == row_time[a][0] and not S.edge_down(pa[0], b, a):
                    nbrs.append((pa[0], b))
                for nb in nbrs:
                    if nb not in comp:
                        comp[nb] = cid; q.append(nb)
    regions = defaultdict(list)
    for k, v in comp.items():
        regions[v].append(k)

    # ---- diagonal units and chains (whole sheet)
    units = []
    useen = set()
    for r in grid_rows:
        for c in grid_cols:
            d = S.diag(r, c)
            if d:
                a = S.rng(r, c)
                if a in useen:
                    continue
                useen.add(a)
                rs = [x for x in range(a[0], a[2] + 1) if x in rowset]
                cs = [x for x in range(a[1], a[3] + 1) if x in colset]
                if rs and cs:
                    units.append(dict(r0=min(rs), r1=max(rs), c0=min(cs), c1=max(cs), d=d))
    parent = list(range(len(units)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]; i = parent[i]
        return i
    for i, a in enumerate(units):
        for j, b in enumerate(units):
            if i == j or b['c0'] != nextc.get(a['c1']) or row_time[b['r1']][0] != row_time[a['r0']][0]:
                continue
            same_region = comp.get((a['r0'], a['c0'])) == comp.get((b['r0'], b['c0']))
            step = nextr.get(b['r1']) == a['r0']
            flat = b['r1'] == a['r0']
            if same_region and (step or flat):
                parent[find(i)] = find(j)
    cg = defaultdict(list)
    for i, u in enumerate(units):
        cg[find(i)].append(u)
    slot_rows = {}
    for sl in slots:
        for rr in sl['rows']:
            slot_rows[rr] = sl['rows']
    ALLCHAINS = []
    for us in cg.values():
        r0 = min(u['r0'] for u in us); r1 = max(u['r1'] for u in us)
        ALLCHAINS.append(dict(units=us, r0=min(slot_rows[r0]), r1=max(slot_rows[r1]),
                              c0=min(u['c0'] for u in us), c1=max(u['c1'] for u in us)))

    entries = []
    for rid, cells in regions.items():
        cellset = set(cells)
        # text anchors inside region
        texts = []
        seen_anchor = set()
        for (r, c) in sorted(cells):
            a = S.rng(r, c)
            if a in seen_anchor:
                continue
            seen_anchor.add(a)
            v = S.val(r, c)
            if v and v.strip() and v.strip() not in ('`', "'"):
                al = S.ws.cell(a[0], a[1]).alignment
                texts.append(dict(r=a[0], c=a[1], r1=a[2], c1=a[3], v=re.sub(r'[ \t]+', ' ', v.strip()), raw=v, h=al.horizontal, fill=fill_of(a[0], a[1])))
        if not texts:
            continue
        chains = [ch for ch in ALLCHAINS if any((r, c) in cellset for u in ch['units'] for r in range(u['r0'], u['r1'] + 1) for c in range(u['c0'], u['c1'] + 1))]
        units = [u for ch in chains for u in ch['units']]
        rows_all = sorted({r for r, _ in cells}); cols_all = sorted({c for _, c in cells})

        def emit(txts, rows, cols, parity):
            if not txts or not rows:
                return
            txts = sorted(txts, key=lambda t: (t['r'], t['c']))
            keep = []
            for t in txts:
                n_ = re.sub(r'\s+', '', t['v']).lower()
                if any(n_ != re.sub(r'\s+', '', o['v']).lower() and n_ in re.sub(r'\s+', '', o['v']).lower() for o in txts):
                    continue
                if any(n_ == re.sub(r'\s+', '', o['v']).lower() for o in keep):
                    continue
                keep.append(t)
            txts = keep
            text = '\n'.join(t['v'] for t in txts)
            d0 = row_time[rows[0]][0]
            h0 = min(row_time[r][1] for r in rows); h1 = max(row_time[r][2] for r in rows)
            sgs = sorted({col_owner[c] for c in cols})
            entries.append(dict(sheet=sheet_key, day=d0, s=h0, e=h1, parity=parity, cols=[list(x) for x in sgs], text=text,
                                ytop=txts[0]['r']))

        def rect_emit(texts):
            # text-centric rectangles: start from the text's own merged range, grow to whole
            # slots / semigroup blocks, then to further slots/blocks only if the region covers
            # them completely and they hold no other text.
            text_cells = set()
            for t in texts:
                for rr in range(t['r'], t['r1'] + 1):
                    for cc in range(t['c'], t['c1'] + 1):
                        text_cells.add((rr, cc))
            slot_of = {}
            for sl in slots:
                for rr in sl['rows']:
                    slot_of[rr] = sl
            rects = []
            for t in texts:
                rows = [x for x in grid_rows if t['r'] <= x <= t['r1']]
                cols = [x for x in grid_cols if t['c'] <= x <= t['c1']]
                if not rows or not cols:
                    continue
                own = {(rr, cc) for rr in rows for cc in cols}

                def free(cells_):
                    return all(x in cellset and (x in own or x not in text_cells) for x in cells_)
                srows = sorted({rr for x in rows for rr in slot_of[x]['rows']})
                exact = bool(t.get('fill')) and len(rows) >= 1 and (t['r1'] > t['r'] or len(slot_of[rows[0]]['rows']) > 2)
                if not exact and free({(rr, cc) for rr in srows for cc in cols}):
                    rows = srows
                blocks = [cs for _, _, cs in sgcols if set(cs) & set(cols)]
                bcols = sorted({c for cs in blocks for c in cs})
                if free({(rr, cc) for rr in rows for cc in bcols}):
                    cols = bcols
                changed = True
                while changed:
                    changed = False
                    for direction in ((1, -1) if not t.get('fill') else ()):
                        if direction < 0:
                            idx = [i for i, sl in enumerate(slots) if rows[0] in sl['rows']][0]
                        else:
                            idx = [i for i, sl in enumerate(slots) if rows[-1] in sl['rows']][0]
                        j = idx + direction
                        if 0 <= j < len(slots) and slots[j]['day'] == slots[idx]['day']:
                            cand = {(rr, cc) for rr in slots[j]['rows'] for cc in cols}
                            if cand and free(cand):
                                rows = sorted(set(rows) | set(slots[j]['rows'])); changed = True
                    own_groups = {col_owner[c][0] for c in cols}
                    for g_, _, cs in sgcols:
                        if set(cs) & set(cols):
                            continue
                        if cs[0] == nextc.get(cols[-1]) or nextc.get(cs[-1]) == cols[0]:
                            cand = {(rr, cc) for rr in rows for cc in cs}
                            same_fill = t['fill'] is not None and all(fill_of(rr, cc) == t['fill'] for rr, cc in cand)
                            if free(cand) and (g_ in own_groups or same_fill):
                                cols = sorted(set(cols) | set(cs)); changed = True
                rects.append((tuple(rows), tuple(cols), t))
            byrect = defaultdict(list)
            for rows, cols, t in rects:
                byrect[(rows, cols)].append(t)
            for (rows, cols), ts in byrect.items():
                emit(ts, list(rows), list(cols), 'all')

        if not units:
            rect_emit(texts)
            continue
        chains.sort(key=lambda ch: ch['r0'])
        def touches_chain(t):
            return any(not (t['r1'] < ch['r0'] or t['r'] > ch['r1']) and not (t['c1'] < ch['c0'] or t['c'] > ch['c1']) for ch in chains)
        plain = [t for t in texts if not touches_chain(t)]
        if plain:
            rect_emit(plain)
        texts = [t for t in texts if touches_chain(t)]
        split_texts = []
        for t in texts:
            parts = [p.strip() for p in re.split(r'[ \t]{4,}|\n\s{4,}', t['raw'].strip()) if p.strip()]
            if len(parts) == 2 and S.diag(t['r'], t['c']):
                a_ = dict(t, v=re.sub(r'\s+', ' ', parts[0]), h='left')
                b_ = dict(t, v=re.sub(r'\s+', ' ', parts[1]), h='right')
                split_texts += [a_, b_]
            else:
                split_texts.append(t)
        texts = split_texts
        buckets = defaultdict(list)
        for t in texts:
            tc = list(range(t['c'], t['c1'] + 1))
            def score(i):
                ch = chains[i]
                colhit = any(ch['c0'] <= x <= ch['c1'] for x in tc)
                rowhit = ch['r0'] <= t['r'] <= ch['r1'] or ch['r0'] <= t['r1'] <= ch['r1']
                dist = 0 if rowhit else min(abs(ch['r0'] - t['r']), abs(ch['r1'] - t['r']))
                return (0 if colhit else 1, 0 if rowhit else 1, dist)
            ci = min(range(len(chains)), key=score)
            ch = chains[ci]
            cov = [u for u in ch['units'] if u['c0'] <= t['c'] <= u['c1']] or [u for u in ch['units'] if u['c0'] <= t['c1'] <= u['c1']]
            if cov:
                u = cov[0]
                if t['r1'] < u['r0']:
                    side = 'odd'
                elif t['r'] > u['r1']:
                    side = 'even'
                elif t['h'] == 'right':
                    side = 'even'
                elif t['h'] == 'left':
                    side = 'odd'
                else:
                    side = 'even' if t['c'] > ch['c0'] and t.get('fill') else 'odd'
            else:
                side = 'odd' if t['c'] < ch['c0'] else 'even'
            buckets[(ci, side)].append(t)
        for (ci, side) in list(buckets):
            other = (ci, 'even' if side == 'odd' else 'odd')
            if (ci, side) in buckets and all(ROOM_ONLY(t['v']) for t in buckets[(ci, side)]) and other in buckets \
                    and not any(ROOM_ONLY(t['v']) is False and re.search(r'\d{2,3}', t['v']) for t in buckets[other]):
                buckets[other] = buckets[(ci, side)] + buckets[other]
                del buckets[(ci, side)]
        slot_of = {}
        for sl in slots:
            for rr in sl['rows']:
                slot_of[rr] = sl
        for (ci, side), ts in buckets.items():
            ch = chains[ci]
            crow = [r for r in grid_rows if ch['r0'] <= r <= ch['r1']]
            ccol = [c for c in grid_cols if ch['c0'] <= c <= ch['c1']]
            # coloured box smaller than the diagonal cell -> the activity uses only the box's rows
            def own_rows(t):
                rows_t = [r for r in grid_rows if t['r'] <= r <= t['r1']]
                if t.get('fill') and rows_t and len(rows_t) < len(crow):
                    H = lambda r: (ws.row_dimensions[r].height or 15.0)
                    out_rows = set()
                    for sl in {id(slot_of[x]): slot_of[x] for x in rows_t}.values():
                        cov = sum(H(r) for r in sl['rows'] if r in rows_t)
                        tot = sum(H(r) for r in sl['rows'])
                        if cov >= 0.5 * tot:
                            out_rows |= set(sl['rows'])
                    return sorted(out_rows) or sorted({rr for x in rows_t for rr in slot_of[x]['rows']})
                return None
            real = [t for t in ts if not CONT(t['v'])]
            slots_used = {id(slot_of[t['r']]) for t in real}
            fixed = [(t, own_rows(t)) for t in ts]
            if any(fr for _, fr in fixed):
                for t, fr in fixed:
                    emit([t], fr or crow, ccol, side)
                    if fr:
                        extra = [c for c in grid_cols if t['c'] <= c <= t['c1'] and c not in ccol]
                        # whole semigroup blocks of the box that lie outside the diagonal: every week
                        blocks_out = [cs for _, _, cs in sgcols if set(cs) <= set(extra)]
                        if blocks_out:
                            emit([t], fr, sorted({c for cs in blocks_out for c in cs}), 'all')
            elif len(real) >= 2 and len(slots_used) >= 2 and all((t['r1'] - t['r']) >= 0 for t in real):
                # several activities on one side, each in its own slot
                groups_by_slot = defaultdict(list)
                for t in ts:
                    groups_by_slot[id(slot_of[t['r']])].append(t)
                # attach room-only fragments to the slot above them
                for k in list(groups_by_slot):
                    if all(ROOM_ONLY(t['v']) for t in groups_by_slot[k]) and len(groups_by_slot) > 1:
                        pass
                for k, g in groups_by_slot.items():
                    emit(g, list(slot_of[g[0]['r']]['rows']), ccol, side)
            else:
                rows_use = crow
                if len(real) == 1 and len({id(slot_of[r]) for r in crow}) >= 2 and re.search(r'(?<![\d.])2\s?[LPS]\b', real[0]['v']):
                    rows_use = list(slot_of[real[0]['r']]['rows'])
                emit(ts, rows_use, ccol, side)
    # merge room-only fragments into the activity above them (same day, columns, parity)
    out = []
    for e in sorted(entries, key=lambda e: (e['day'], e['ytop'])):
        if CONT(e['text']):
            cand = [o for o in out if o['day'] == e['day'] and o['cols'] == e['cols'] and o['parity'] == e['parity'] and o['ytop'] < e['ytop'] and o['e'] >= e['s']]
            if cand:
                cand[-1]['text'] += '\n' + e['text']
                if not ROOM_ONLY(e['text']):
                    cand[-1]['e'] = max(cand[-1]['e'], e['e'])
                continue
        out.append(e)
    entries = out
    return dict(groups=groups, sgcols=sgcols, slots=slots, entries=entries)


SHEETS = ['AN I', 'AN I En', 'AN II', 'AN II En', 'AN III', 'AN III En', 'AN IV ', 'AN IV En', 'MASTER AN I', 'MASTER AN II']

if __name__ == '__main__':
    wb = openpyxl.load_workbook(sys.argv[1])
    out = {}
    for name in SHEETS:
        res = parse_sheet(wb[name], name.strip())
        out[name.strip()] = res
        print(name, 'groups', [(g['code'], len(g['cols'])) for g in res['groups']], 'sg', [(a, b, len(c)) for a, b, c in res['sgcols']], 'slots', len(res['slots']), 'entries', len(res['entries']))
    json.dump(out, open(sys.argv[2], 'w'), ensure_ascii=False, indent=0)
