"""Track changes mixin: insert, delete, accept, reject."""

from __future__ import annotations

from lxml import etree

from .base import W, _now_iso, _preserve


class TracksMixin:
    """Insert/delete/accept/reject text with tracked changes markup."""

    def insert_text(
        self,
        para_id: str,
        text: str,
        *,
        position: str = "end",
        author: str = "Claude",
    ) -> dict:
        """Insert text with Word track-changes markup (w:ins).

        Args:
            para_id: Target paragraph paraId.
            text: Text to insert.
            position: 'end', 'start', or a substring after which to insert.
            author: Author name for the revision.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cid = self._next_markup_id(doc)
        now = _now_iso()

        ins = etree.Element(f"{W}ins")
        ins.set(f"{W}id", str(cid))
        ins.set(f"{W}author", author)
        ins.set(f"{W}date", now)

        r = etree.SubElement(ins, f"{W}r")
        t = etree.SubElement(r, f"{W}t")
        _preserve(t, text)

        # The inserted words take the formatting of the run they're placed in:
        # the matched run, the first text run for 'start', the last otherwise.
        text_runs = [el for el in para.findall(f"{W}r") if el.find(f"{W}t") is not None]
        anchor = text_runs[-1] if text_runs else None

        if position == "start":
            anchor = text_runs[0] if text_runs else None
            ppr = para.find(f"{W}pPr")
            if ppr is not None:
                ppr.addnext(ins)
            else:
                para.insert(0, ins)
        elif position == "end":
            para.append(ins)
        else:
            placed = False
            # Search regular runs — split if match is mid-run
            for run_el in list(para.findall(f"{W}r")):
                t_el = run_el.find(f"{W}t")
                if t_el is None or t_el.text is None:
                    continue
                if position not in t_el.text:
                    continue

                full = t_el.text
                end = full.index(position) + len(position)

                if end < len(full):
                    # Match is mid-run: split, insert between halves
                    rpr = run_el.find(f"{W}rPr")
                    rpr_bytes = etree.tostring(rpr) if rpr is not None else None
                    _preserve(t_el, full[:end])
                    after_run = self._make_run(full[end:], rpr_bytes)
                    run_el.addnext(ins)
                    ins.addnext(after_run)
                else:
                    run_el.addnext(ins)
                anchor = run_el
                placed = True
                break

            # Fallback: search w:del elements (delete-then-insert pattern)
            if not placed:
                for del_el in para.findall(f"{W}del"):
                    del_text = "".join(t.text for t in del_el.iter(f"{W}delText") if t.text)
                    if position in del_text:
                        del_el.addnext(ins)
                        anchor = del_el.find(f"{W}r")
                        placed = True
                        break

            if not placed:
                para.append(ins)

        anchor_rpr = anchor.find(f"{W}rPr") if anchor is not None else None
        if anchor_rpr is not None:
            r.insert(0, etree.fromstring(etree.tostring(anchor_rpr)))

        self._mark("word/document.xml")
        return {"change_id": cid, "type": "insertion", "author": author, "date": now}

    def delete_text(
        self,
        para_id: str,
        text: str,
        *,
        author: str = "Claude",
    ) -> dict:
        """Mark text as deleted with Word track-changes markup (w:del).

        Finds the text within runs of the target paragraph and wraps the
        matching portion in <w:del><w:r><w:delText>...</w:delText></w:r></w:del>,
        splitting the run if the match is a substring.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        cid = self._next_markup_id(doc)
        now = _now_iso()

        for run_el in list(para.findall(f"{W}r")):
            t_el = run_el.find(f"{W}t")
            if t_el is None or t_el.text is None:
                continue
            full = t_el.text
            if text not in full:
                continue

            idx = full.index(text)
            rpr = run_el.find(f"{W}rPr")
            rpr_bytes = etree.tostring(rpr) if rpr is not None else None
            parent = run_el.getparent()
            pos = list(parent).index(run_el)
            parent.remove(run_el)

            insert_at = pos

            # Text before
            if idx > 0:
                before = self._make_run(full[:idx], rpr_bytes)
                parent.insert(insert_at, before)
                insert_at += 1

            # Deletion
            del_el = etree.Element(f"{W}del")
            del_el.set(f"{W}id", str(cid))
            del_el.set(f"{W}author", author)
            del_el.set(f"{W}date", now)
            del_run = etree.SubElement(del_el, f"{W}r")
            if rpr_bytes:
                del_run.append(etree.fromstring(rpr_bytes))
            dt = etree.SubElement(del_run, f"{W}delText")
            _preserve(dt, text)
            parent.insert(insert_at, del_el)
            insert_at += 1

            # Text after
            end = idx + len(text)
            if end < len(full):
                after = self._make_run(full[end:], rpr_bytes)
                parent.insert(insert_at, after)

            self._mark("word/document.xml")
            return {"change_id": cid, "type": "deletion", "author": author, "date": now}

        raise ValueError(
            f"Text '{text}' not found in a single run of paragraph '{para_id}'. "
            "If the text spans multiple runs, try searching for a shorter substring."
        )

    # ── Accept / Reject ───────────────────────────────────────────────────

    def _matches_author(self, el: etree._Element, author: str | None) -> bool:
        """Check if element's author matches filter (None = match all)."""
        if author is None:
            return True
        return el.get(f"{W}author") == author

    def _tracked_rows(self, doc: etree._Element, tag: str, author: str | None) -> list:
        """Rows whose ``w:trPr`` carries a ``tag`` revision (w:ins / w:del) matching ``author``."""
        return [
            mark.getparent().getparent()
            for mark in doc.iter(tag)
            if mark.getparent().tag == f"{W}trPr" and self._matches_author(mark, author)
        ]

    def accept_changes(
        self,
        *,
        author: str | None = None,
    ) -> dict:
        """Accept tracked changes: keep insertions, remove deletions, keep new formatting.

        Args:
            author: If set, only accept changes by this author.
        """
        doc = self._require("word/document.xml")
        count = 0

        # Accept row deletions: remove the whole row, not just its w:trPr marker
        for tr in self._tracked_rows(doc, f"{W}del", author):
            tr.getparent().remove(tr)
            count += 1

        # Accept insertions: unwrap w:ins (promote children)
        for ins in list(doc.iter(f"{W}ins")):
            if not self._matches_author(ins, author):
                continue
            parent = ins.getparent()
            idx = list(parent).index(ins)
            for child in list(ins):
                ins.remove(child)
                parent.insert(idx, child)
                idx += 1
            parent.remove(ins)
            count += 1

        # Accept deletions: remove w:del entirely
        for del_el in list(doc.iter(f"{W}del")):
            if not self._matches_author(del_el, author):
                continue
            del_el.getparent().remove(del_el)
            count += 1

        # Accept formatting revisions: the new run / paragraph properties stay, and the record of
        # the old ones goes (set_formatting writes w:rPrChange; Word writes w:pPrChange too).
        for tag in ("rPrChange", "pPrChange"):
            for change in list(doc.iter(f"{W}{tag}")):
                if not self._matches_author(change, author):
                    continue
                change.getparent().remove(change)
                count += 1

        if count > 0:
            self._mark("word/document.xml")
        scope = "by_author" if author else "all"
        return {"accepted": count, "scope": scope}

    def reject_changes(
        self,
        *,
        author: str | None = None,
    ) -> dict:
        """Reject tracked changes: remove insertions, restore deletions and old formatting.

        Args:
            author: If set, only reject changes by this author.
        """
        doc = self._require("word/document.xml")
        count = 0

        # Reject row insertions: remove the whole row, not just its w:trPr marker
        for tr in self._tracked_rows(doc, f"{W}ins", author):
            tr.getparent().remove(tr)
            count += 1

        # Reject insertions: remove w:ins and its content entirely
        for ins in list(doc.iter(f"{W}ins")):
            if not self._matches_author(ins, author):
                continue
            ins.getparent().remove(ins)
            count += 1

        # Reject deletions: unwrap w:del, convert w:delText back to w:t
        for del_el in list(doc.iter(f"{W}del")):
            if not self._matches_author(del_el, author):
                continue
            parent = del_el.getparent()
            idx = list(parent).index(del_el)
            for child in list(del_el):
                # Convert w:delText -> w:t inside each run
                for dt in child.findall(f"{W}delText"):
                    dt.tag = f"{W}t"
                del_el.remove(child)
                parent.insert(idx, child)
                idx += 1
            parent.remove(del_el)
            count += 1

        # Reject formatting revisions: put back the properties the revision recorded. A run's
        # w:rPr becomes the old w:rPr whole; a paragraph's w:pPr takes the old properties back but
        # keeps its paragraph-mark w:rPr and its w:sectPr, which a w:pPrChange never records.
        for change in list(doc.iter(f"{W}rPrChange")):
            if not self._matches_author(change, author):
                continue
            rpr = change.getparent()
            old = change.find(f"{W}rPr")
            for child in list(rpr):
                rpr.remove(child)
            for child in list(old) if old is not None else []:
                rpr.append(child)
            count += 1
        for change in list(doc.iter(f"{W}pPrChange")):
            if not self._matches_author(change, author):
                continue
            ppr = change.getparent()
            old = change.find(f"{W}pPr")
            for child in list(ppr):
                if child.tag not in (f"{W}rPr", f"{W}sectPr"):
                    ppr.remove(child)
            for idx, child in enumerate(list(old) if old is not None else []):
                ppr.insert(idx, child)
            count += 1

        if count > 0:
            self._mark("word/document.xml")
        scope = "by_author" if author else "all"
        return {"rejected": count, "scope": scope}
