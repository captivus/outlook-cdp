"""A side-by-side review page for the contacts a deduplication plan leaves for a person to judge.

    from outlook_cdp import dedupe, dedupe_review

    page = dedupe_review.render(dedupe.load("backup.json"), dedupe.load("plan.json"))
    open("review.html", "w", encoding="utf-8").write(page)

One self-contained HTML file, nothing fetched: the contacts in it are personal data, so it is meant to be
opened from disk.  Three tabs, one per kind of decision the plan could not make:

- **Each holds something** - contacts sharing an email or phone where no copy holds all of another, so
  none was deleted (``kept_with_differences``).
- **Same name only** - contacts sharing only a name, which are never deleted without approval.
- **In a contact group** - copies that would have been deleted but are members of a group.

Each group is a table with one column per contact and one row per field.  A value found on only one
contact is marked as such, a value that differs between contacts is marked as differing, and under the
table each contact lists exactly what `dedupe.coverage_gaps` found it holds that the first one lacks.

In the first two tabs each column also carries a decision: delete this contact, in favour of which other
one, and - when that other one lacks something - an explicit acceptance of exactly what will be lost.  The
page saves nothing anywhere but the viewer's own browser; "Download decisions" writes a file for
``contacts dedupe-apply --decisions``, which re-checks every decision (`dedupe.checked_decisions`).
"""
from __future__ import annotations

import html
import json

from .contacts import extended
from .dedupe import coverage_gaps, display, ews_gaps, rest_gaps

_NAME_FIELDS = [("Title", None, "DisplayNamePrefix"), ("Given name", "GivenName", None),
                ("Middle name", "MiddleName", None), ("Surname", "Surname", None),
                ("Suffix", "Generation", None), ("Nickname", "NickName", None),
                ("Display name", "DisplayName", None), ("File as", "FileAs", None)]
_WORK_FIELDS = [("Company", "CompanyName"), ("Job title", "JobTitle"), ("Department", "Department"),
                ("Office", "OfficeLocation"), ("Profession", "Profession"), ("Manager", "Manager"),
                ("Assistant", "AssistantName"), ("Spouse", "SpouseName")]
_PHONE_LABELS = {"MobilePhone": "Mobile", "HomePhone": "Home phone", "HomePhone2": "Home phone 2",
                 "BusinessPhone": "Business phone", "BusinessPhone2": "Business phone 2",
                 "OtherTelephone": "Other phone", "CompanyMainPhone": "Company main", "PrimaryPhone": "Primary",
                 "BusinessFax": "Business fax", "HomeFax": "Home fax", "OtherFax": "Other fax", "Pager": "Pager",
                 "CarPhone": "Car phone", "Callback": "Callback", "AssistantPhone": "Assistant phone",
                 "RadioPhone": "Radio", "Telex": "Telex", "Isdn": "ISDN", "TtyTddPhone": "TTY/TDD"}


def _s(v) -> str:
    return " ".join(str(v or "").split())


def _rows(entry: dict, ignore=()) -> list[tuple[str, str, str]]:
    """One contact as (section, label, value) rows; empty values and ignored categories left out."""
    ignored = {c.lower() for c in ignore}
    r, x = entry["record"], entry.get("ews") or {}
    out: list[tuple[str, str, str]] = []
    add = lambda sec, label, v: out.append((sec, label, v)) if _s(v) else None  # noqa: E731
    for label, rest, ews in _NAME_FIELDS:
        add("Name", label, r.get(rest) if rest else x.get(ews))
    emails = []
    for e in r.get("EmailAddresses") or []:
        a, n = e.get("Address") or "", e.get("Name") or ""
        if not _s(a):
            continue
        emails.append(a + (f"  ({n})" if _s(n) and _s(n).lower() != _s(a).lower() else ""))
    for i, e in enumerate(sorted(emails), 1):
        add("Email", f"Email {i}", e)
    for p in x.get("PhoneNumbers") or []:
        add("Phone", _PHONE_LABELS.get(p.get("Key"), p.get("Key") or "Phone"), p.get("PhoneNumber"))
    if not x:                                          # no second view: fall back to the REST slots
        add("Phone", "Mobile", r.get("MobilePhone1"))
        for i, p in enumerate(r.get("HomePhones") or [], 1):
            add("Phone", f"Home phone {i}", p)
        for i, p in enumerate(r.get("BusinessPhones") or [], 1):
            add("Phone", f"Business phone {i}", p)
    for key, label in (("HomeAddress", "Home"), ("BusinessAddress", "Business"), ("OtherAddress", "Other")):
        a = r.get(key) or {}
        add("Address", label, ", ".join(_s(a.get(k)) for k in ("Street", "City", "State", "PostalCode",
                                                             "CountryOrRegion") if _s(a.get(k))))
    for label, key in _WORK_FIELDS:
        add("Work & family", label, r.get(key))
    add("Work & family", "Children", ", ".join(r.get("Children") or []))
    add("Other", "Business web page", r.get("BusinessHomePage"))
    for name, v in sorted(extended(r).items()):
        if not name.endswith(("phone", "fax")) and name not in ("pager", "telex", "isdn"):
            add("Other", name.replace("_", " ").capitalize(), v)
    add("Other", "IM", ", ".join(r.get("ImAddresses") or []))
    add("Other", "Birthday", (r.get("Birthday") or "")[:10])
    add("Other", "Anniversary", (x.get("WeddingAnniversary") or "")[:10])
    add("Other", "Categories", ", ".join(sorted((c for c in r.get("Categories") or [] if c.lower() not in ignored),
                                                key=str.lower)))
    if (x.get("Sensitivity") or "Normal") != "Normal":
        add("Other", "Sensitivity", x.get("Sensitivity"))
    if (x.get("Flag") or {}).get("FlagStatus") not in (None, "NotFlagged"):
        add("Other", "Follow-up flag", x["Flag"]["FlagStatus"])
    notes = r.get("PersonalNotes") or ((x.get("Body") or {}).get("Value") or "")
    if _s(notes):
        fmt = (x.get("Body") or {}).get("BodyType")
        add("Notes", "Notes" + (" (HTML)" if fmt == "HTML" else ""), notes.replace("\r\n", "\n").strip())
    return out


def _gaps(a: dict, b: dict, ignore) -> list[str]:
    """What `a` holds that `b` lacks, for reading: the REST view's report, which names fields plainly, or
    the service.svc view's when REST finds nothing - the two mostly report the same difference twice."""
    return rest_gaps(a, b, ignore) or ews_gaps(a, b, ignore)


def _decide_row(entries: list[dict]) -> str:
    cells = "".join(
        f"<td class='decide' data-col='{i}'><label class='del'><input type='checkbox' class='d-del'> Delete this "
        f"contact</label><div class='d-more' hidden><label class='kp'>keeping <select class='d-keep'></select>"
        f"</label><div class='d-loss'></div><label class='acc' hidden><input type='checkbox' class='d-acc'> "
        f"I accept losing this</label></div></td>" for i in range(len(entries)))
    return f"<tr class='decrow'><th scope='row'>Decision</th>{cells}</tr>"


def _decision_data(entries: list[dict], ignore, planned: dict) -> str:
    """What the page's script needs for one group: ids, names, and the loss of every possible deletion."""
    loss = {}
    for i, a in enumerate(entries):
        for j, b in enumerate(entries):
            if i != j:
                loss[f"{i}>{j}"] = {"show": _gaps(a, b, ignore), "full": coverage_gaps(a, b, ignore)}
    ids = [e["record"]["Id"] for e in entries]
    data = {"ids": ids, "names": [display(e) for e in entries], "loss": loss,
            "planned": {ids.index(d): ids.index(k) for d, k in planned.items() if d in ids and k in ids}}
    return json.dumps(data, ensure_ascii=False).replace("</", "<\\/")


def _table(entries: list[dict], folders: dict, ignore, note: str = "", decide: bool = False) -> str:
    cards = [dict(((sec, label), v) for sec, label, v in _rows(e, ignore)) for e in entries]
    order: list[tuple[str, str]] = []
    for e in entries:
        for sec, label, _ in _rows(e, ignore):
            if (sec, label) not in order:
                order.append((sec, label))
    sections = ["Name", "Email", "Phone", "Address", "Work & family", "Other", "Notes"]
    order.sort(key=lambda k: sections.index(k[0]))
    head = "".join(
        f"<th scope='col'><div class='who'>{_photo(e)}<div><div class='nm'>{html.escape(display(e))}</div>"
        f"<div class='meta'>{html.escape(folders.get(e.get('folder'), 'Contacts'))}<br>"
        f"modified {html.escape((e['record'].get('LastModifiedDateTime') or '')[:10])}</div></div></div></th>"
        for e in entries)
    body, last = [], None
    for sec, label in order:
        if sec != last:
            body.append(f"<tr class='sec'><th colspan='{len(entries) + 1}'>{html.escape(sec)}</th></tr>")
            last = sec
        vals = [c.get((sec, label), "") for c in cards]
        filled = [v for v in vals if _s(v)]
        distinct = {_s(v) for v in filled}
        cells = []
        for v in vals:
            cls = ""
            if _s(v) and len(filled) == 1 and len(entries) > 1:
                cls = "only"
            elif _s(v) and len(distinct) > 1:
                cls = "diff"
            elif not _s(v) and filled:
                cls = "none"
            text = html.escape(v) if _s(v) else "—"
            cells.append(f"<td class='{cls}'>{'<div class=notes>' + text + '</div>' if sec == 'Notes' else text}</td>")
        body.append(f"<tr><th scope='row'>{html.escape(label)}</th>{''.join(cells)}</tr>")
    first = entries[0]
    holds = []
    for i, e in enumerate(entries):
        if i == 0:
            gaps = sorted({g for other in entries[1:] for g in _gaps(first, other, ignore)})
            title = f"{html.escape(display(e))} (first column) has, that some other column lacks"
        else:
            gaps = _gaps(e, first, ignore)
            title = f"Column {i + 1} has, that the first column lacks"
        items = "".join(f"<li>{html.escape(g)}</li>" for g in gaps[:40]) or "<li class='nothing'>nothing</li>"
        more = f"<li class='nothing'>…and {len(gaps) - 40} more</li>" if len(gaps) > 40 else ""
        holds.append(f"<div class='holds'><h4>{title}</h4><ul>{items}{more}</ul></div>")
    decision = _decide_row(entries) if decide else ""
    return (f"{note}<div class='scroll'><table><thead><tr><th class='corner'></th>{head}</tr>{decision}</thead>"
            f"<tbody>{''.join(body)}</tbody></table></div><div class='holdsgrid'>{''.join(holds)}</div>")


def _photo(e: dict) -> str:
    p = e.get("photo")
    if not p:
        return "<div class='ph none' aria-hidden='true'></div>"
    return f"<img class='ph' alt='' src='data:{html.escape(p['type'])};base64,{p['b64']}'>"


def _group(n: int, kind: str, entries: list[dict], folders: dict, ignore, note: str = "",
           planned: dict | None = None) -> str:
    names = sorted({display(e) for e in entries})
    search = html.escape(" ".join(names).lower())
    title = html.escape(" · ".join(names[:3]) + (f" +{len(names) - 3}" if len(names) > 3 else ""))
    decide = kind != "group"
    data = (f"<script type='application/json' class='gdata'>{_decision_data(entries, ignore, planned or {})}"
            f"</script>") if decide else ""
    return (f"<details class='grp' data-kind='{kind}' data-search='{search}'><summary><span class='n'>{n}</span>"
            f"<span class='t'>{title}</span><span class='st'></span><span class='c'>{len(entries)} contacts</span>"
            f"</summary>{data}{_table(entries, folders, ignore, note, decide)}</details>")


def render(backup: dict, plan: dict) -> str:
    """The review page for `plan` (made from `backup`), as one HTML document."""
    by = {e["record"]["Id"]: e for e in backup["contacts"]}
    folders = {f["id"]: f["name"] for f in backup["folders"]}
    ignore = plan.get("ignore_categories") or ()
    parts = {"differ": [], "name": [], "group": []}
    for x in plan["kept_with_differences"]:
        parts["differ"].append([by[c["id"]] for c in x["contacts"] if c["id"] in by])
    for g in plan["name_only"]:
        ids = [c["id"] for c in g["kept"]] + [d["delete"]["id"] for d in g["deletable"]]
        parts["name"].append([by[i] for i in ids if i in by])
    for d in plan["in_a_contact_group"]:
        parts["group"].append([by[d["keep"]["id"]], by[d["delete"]["id"]]])
    planned = {d["delete"]["id"]: d["keep"]["id"] for g in plan["name_only"] for d in g["deletable"]}
    notes = {"name": "<p class='why'>These share only a name. Where one copy is fully held by another, "
                     "ticking Delete on it loses nothing; otherwise the page shows what would be lost.</p>",
             "group": "<p class='why'>The second column would be deleted - the first holds everything in it - "
                      "but its address is a member of a contact group, so it is left alone.</p>"}
    sections = []
    for kind in ("differ", "name", "group"):
        groups = sorted(parts[kind], key=lambda es: display(es[0]).lower())
        sections.append("".join(_group(i, kind, es, folders, ignore, notes.get(kind, ""), planned)
                                for i, es in enumerate(groups, 1)))
    counts = {k: len(v) for k, v in parts.items()}
    return _PAGE.format(
        mailbox=html.escape(backup["mailbox"]), taken=html.escape(backup["taken_at"]),
        differ=counts["differ"], name=counts["name"], group=counts["group"],
        ignored=html.escape(", ".join(ignore) or "none"),
        body="".join(f"<section id='{k}' class='tab' {'hidden' if k != 'differ' else ''}>{s}</section>"
                     for k, s in zip(("differ", "name", "group"), sections)),
        summary=html.escape(json.dumps(plan.get("summary", {}))),
        meta=json.dumps({"backup_taken_at": backup["taken_at"], "mailbox": backup["mailbox"]}))


_PAGE = """<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Contact Duplicates Review</title>
<style>
:root{{--bg:#f7f7f5;--panel:#fff;--ink:#1d1d1b;--muted:#6b6b66;--line:#e3e2dd;--only:#e6f4ea;--only-ink:#14532d;
--diff:#fff4d6;--diff-ink:#7a4b00;--none:#fafaf8;--accent:#2f5bd3;--sec:#f1f0ec;--bad:#b42318}}
@media (prefers-color-scheme:dark){{:root:not([data-theme="light"]){{--bg:#151514;--panel:#1e1e1c;--ink:#ecebe6;
--muted:#9c9b94;--line:#33322f;--only:#163a24;--only-ink:#b7ebc6;--diff:#3d3213;--diff-ink:#f3d58b;--none:#1a1a18;
--accent:#8fb0ff;--sec:#262624;--bad:#ff8a80}}}}
:root[data-theme="dark"]{{--bg:#151514;--panel:#1e1e1c;--ink:#ecebe6;--muted:#9c9b94;--line:#33322f;--only:#163a24;
--only-ink:#b7ebc6;--diff:#3d3213;--diff-ink:#f3d58b;--none:#1a1a18;--accent:#8fb0ff;--sec:#262624;--bad:#ff8a80}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 system-ui,-apple-system,Segoe UI,sans-serif}}
header{{padding:20px 16px 8px;max-width:1400px;margin:0 auto}}
h1{{font-size:20px;margin:0 0 4px}} .sub{{color:var(--muted);font-size:13px}}
.bar{{position:sticky;top:0;z-index:5;background:var(--bg);border-bottom:1px solid var(--line)}}
.bar-in{{max-width:1400px;margin:0 auto;padding:8px 16px;display:flex;flex-wrap:wrap;gap:8px;align-items:center}}
.tabs button{{font:inherit;border:1px solid var(--line);background:var(--panel);color:var(--ink);padding:6px 12px;
border-radius:999px;cursor:pointer}} .tabs button[aria-selected=true]{{border-color:var(--accent);color:var(--accent);font-weight:600}}
input[type=search]{{font:inherit;flex:1;min-width:160px;max-width:320px;padding:6px 10px;border:1px solid var(--line);
border-radius:8px;background:var(--panel);color:var(--ink)}}
.ctl button{{font:inherit;background:none;border:none;color:var(--accent);cursor:pointer;padding:4px}}
.legend{{display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:var(--muted)}}
.sw{{display:inline-block;width:12px;height:12px;border-radius:3px;vertical-align:-2px;margin-right:4px;border:1px solid var(--line)}}
main{{max-width:1400px;margin:0 auto;padding:12px 16px 48px}}
.grp{{background:var(--panel);border:1px solid var(--line);border-radius:10px;margin:0 0 8px}}
.grp summary{{display:flex;gap:10px;align-items:baseline;padding:10px 14px;cursor:pointer;list-style:none}}
.grp summary::-webkit-details-marker{{display:none}}
.grp summary .n{{color:var(--muted);font-variant-numeric:tabular-nums;min-width:2.5em}}
.grp summary .t{{font-weight:600;flex:1;overflow-wrap:anywhere}} .grp summary .c{{color:var(--muted);font-size:12px;white-space:nowrap}}
.grp[open] summary{{border-bottom:1px solid var(--line)}}
.scroll{{overflow-x:auto}}
table{{border-collapse:collapse;min-width:100%}}
th,td{{text-align:left;vertical-align:top;padding:6px 10px;border-bottom:1px solid var(--line)}}
td{{min-width:200px;max-width:340px;overflow-wrap:anywhere}}
thead th{{background:var(--panel);font-weight:400}} th.corner{{min-width:120px}}
tbody th[scope=row]{{color:var(--muted);font-weight:500;white-space:nowrap;position:sticky;left:0;background:var(--panel)}}
tr.sec th{{background:var(--sec);font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);padding:4px 10px}}
td.only{{background:var(--only);color:var(--only-ink)}} td.diff{{background:var(--diff);color:var(--diff-ink)}}
td.none{{background:var(--none);color:var(--muted)}}
.notes{{white-space:pre-wrap;max-height:12em;overflow:auto;font-size:13px}}
.who{{display:flex;gap:8px;align-items:center}} .nm{{font-weight:600}} .meta{{color:var(--muted);font-size:12px}}
.ph{{width:40px;height:40px;border-radius:50%;object-fit:cover;flex:none}} .ph.none{{background:var(--sec)}}
.holdsgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:8px;padding:10px 14px 14px}}
.holds h4{{margin:4px 0;font-size:12px;color:var(--muted);font-weight:600}}
.holds ul{{margin:0;padding-left:18px;font-size:13px}} .holds li{{overflow-wrap:anywhere}} .nothing{{color:var(--muted);list-style:none;margin-left:-18px}}
.why{{margin:10px 14px 0;color:var(--muted);font-size:13px}}
tr.decrow th,tr.decrow td{{background:var(--sec);border-bottom:2px solid var(--line);font-size:13px}}
td.decide label{{display:block;cursor:pointer}} td.decide select{{font:inherit;max-width:100%;margin-top:4px;
background:var(--panel);color:var(--ink);border:1px solid var(--line);border-radius:6px}}
td.decide.on{{background:var(--diff)}} td.decide.on.bad{{outline:2px solid var(--bad);outline-offset:-2px}}
.d-loss{{margin:6px 0;font-size:12px}} .d-loss ul{{margin:2px 0;padding-left:16px}} .d-loss .ok{{color:var(--only-ink)}}
.d-loss .warn{{color:var(--bad);font-weight:600}} .acc{{font-weight:600}}
.grp summary .st{{font-size:12px;white-space:nowrap}} .st.ok{{color:var(--only-ink)}} .st.bad{{color:var(--bad)}}
.dl{{display:flex;gap:8px;align-items:center;flex-wrap:wrap}} .dl button{{font:inherit;border:1px solid var(--accent);
background:var(--accent);color:var(--panel);padding:6px 12px;border-radius:8px;cursor:pointer}}
.dl button.ghost{{background:none;color:var(--accent)}} .dl button:disabled{{opacity:.45;cursor:not-allowed}}
.dl .msg{{font-size:12px;color:var(--muted)}}
code{{font-size:12px}}
</style></head><body>
<header><h1>Contact duplicates to review</h1>
<div class="sub">{mailbox} · backup {taken} · categories ignored as information: {ignored}</div></header>
<div class="bar"><div class="bar-in">
<div class="tabs" role="tablist">
<button role="tab" aria-selected="true" data-tab="differ">Each holds something ({differ})</button>
<button role="tab" aria-selected="false" data-tab="name">Same name only ({name})</button>
<button role="tab" aria-selected="false" data-tab="group">In a contact group ({group})</button></div>
<input type="search" placeholder="Filter by name" aria-label="Filter by name">
<span class="ctl"><button data-act="open">Expand all</button><button data-act="close">Collapse all</button></span>
<span class="legend"><span><span class="sw" style="background:var(--only)"></span>only on this contact</span>
<span><span class="sw" style="background:var(--diff)"></span>differs between contacts</span>
<span><span class="sw" style="background:var(--none)"></span>missing here</span></span>
<span class="dl"><button id="dl" disabled>Download decisions</button><button id="clr" class="ghost">Clear all</button>
<span class="msg" id="msg">No decisions yet.</span></span>
</div></div>
<main>{body}</main>
<script>
const tabs=[...document.querySelectorAll('[data-tab]')], q=document.querySelector('input[type=search]');
function current(){{return document.querySelector('.tab:not([hidden])')}}
tabs.forEach(b=>b.onclick=()=>{{tabs.forEach(x=>x.setAttribute('aria-selected',x===b));
  document.querySelectorAll('.tab').forEach(s=>s.hidden=s.id!==b.dataset.tab);filter()}});
function filter(){{const t=q.value.trim().toLowerCase();
  current().querySelectorAll('.grp').forEach(g=>g.hidden=t&&!g.dataset.search.includes(t))}}
q.oninput=filter;
document.querySelectorAll('[data-act]').forEach(b=>b.onclick=()=>current().querySelectorAll('.grp:not([hidden])')
  .forEach(g=>g.open=b.dataset.act==='open'));
</script>
<script>
const META={meta}, KEY='dedupe-decisions:'+META.mailbox+':'+META.backup_taken_at;
const groups=[...document.querySelectorAll('.grp')].filter(g=>g.querySelector('.gdata'))
  .map((el,gi)=>({{el,gi,data:JSON.parse(el.querySelector('.gdata').textContent)}}));
let state={{}};
try{{state=JSON.parse(localStorage.getItem(KEY)||'{{}}')}}catch(e){{state={{}}}}
function save(){{try{{localStorage.setItem(KEY,JSON.stringify(state))}}catch(e){{}}}}
function esc(t){{const d=document.createElement('div');d.textContent=t;return d.innerHTML}}
function deletedIds(){{const s=new Set();groups.forEach(g=>Object.entries(state[g.gi]||{{}}).forEach(([c,v])=>v.del&&s.add(g.data.ids[c])));return s}}
function render(g){{
  const st=state[g.gi]||(state[g.gi]={{}}), n=g.data.ids.length, gone=deletedIds();
  const kept=[...Array(n).keys()].filter(c=>!(st[c]&&st[c].del));
  let bad=0, dels=0;
  g.el.querySelectorAll('td.decide').forEach(td=>{{
    const c=+td.dataset.col, v=st[c]||(st[c]={{}}), more=td.querySelector('.d-more'), sel=td.querySelector('.d-keep');
    td.querySelector('.d-del').checked=!!v.del; more.hidden=!v.del; td.classList.toggle('on',!!v.del);
    if(!v.del){{td.classList.remove('bad');return}}
    dels++;
    const opts=kept.filter(k=>!gone.has(g.data.ids[k]));
    if(!opts.includes(v.keep)) v.keep=opts.includes(g.data.planned[c])?g.data.planned[c]:opts[0];
    sel.innerHTML=opts.map(k=>`<option value="${{k}}">column ${{k+1}}: ${{esc(g.data.names[k])}}</option>`).join('');
    const box=td.querySelector('.d-loss'), acc=td.querySelector('.acc');
    if(v.keep===undefined){{box.innerHTML='<span class="warn">Keep at least one contact in this group.</span>';
      acc.hidden=true;td.classList.add('bad');bad++;return}}
    sel.value=v.keep;
    const L=g.data.loss[c+'>'+v.keep];
    if(!L.full.length){{box.innerHTML='<span class="ok">Nothing is lost: that contact holds everything in this one.</span>';
      acc.hidden=true;v.acc=false}}
    else{{box.innerHTML='<span class="warn">Would be lost:</span><ul>'+L.show.map(x=>'<li>'+esc(x)+'</li>').join('')+'</ul>';
      acc.hidden=false;td.querySelector('.d-acc').checked=!!v.acc}}
    const need=L.full.length&&!v.acc; td.classList.toggle('bad',!!need); if(need) bad++;
  }});
  const chip=g.el.querySelector('.st');
  chip.className='st'+(dels?(bad?' bad':' ok'):'');
  chip.textContent=dels?(bad?`${{bad}} need attention`:`deleting ${{dels}}`):'';
}}
function decisions(){{const out=[];let pending=0;
  groups.forEach(g=>Object.entries(state[g.gi]||{{}}).forEach(([c,v])=>{{if(!v.del)return;
    const L=v.keep===undefined?null:g.data.loss[c+'>'+v.keep];
    if(!L||(L.full.length&&!v.acc)){{pending++;return}}
    out.push({{delete:g.data.ids[c],keep:g.data.ids[v.keep],delete_name:g.data.names[c],keep_name:g.data.names[v.keep],
      accepted_loss:L.full}})}}));return {{out,pending}}}}
function refresh(){{groups.forEach(render);save();const {{out,pending}}=decisions();
  document.getElementById('dl').disabled=!out.length||pending>0;
  document.getElementById('msg').textContent=!out.length&&!pending?'No decisions yet.':
    `${{out.length}} deletion${{out.length==1?'':'s'}} decided`+(pending?` · ${{pending}} need you to accept a loss or keep a contact`:'')}}
groups.forEach(g=>g.el.querySelectorAll('td.decide').forEach(td=>{{const c=+td.dataset.col;
  const v=()=>((state[g.gi]||(state[g.gi]={{}}))[c]||(state[g.gi][c]={{}}));
  td.querySelector('.d-del').onchange=e=>{{v().del=e.target.checked;v().acc=false;refresh()}};
  td.querySelector('.d-keep').onchange=e=>{{v().keep=+e.target.value;v().acc=false;refresh()}};
  td.querySelector('.d-acc').onchange=e=>{{v().acc=e.target.checked;refresh()}};}}));
document.getElementById('clr').onclick=()=>{{if(confirm('Clear every decision on this page?')){{state={{}};refresh()}}}};
document.getElementById('dl').onclick=()=>{{const {{out}}=decisions();
  const doc={{format:1,backup_taken_at:META.backup_taken_at,mailbox:META.mailbox,made_at:new Date().toISOString(),decisions:out}};
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify(doc,null,1)],{{type:'application/json'}}));
  a.download='decisions-'+META.backup_taken_at+'.json';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}};
refresh();
</script>
<!-- plan summary: {summary} -->
</body></html>
"""
