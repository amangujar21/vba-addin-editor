"""Project-wide search, replace, and procedure outline (IMP-08)."""

from __future__ import annotations

import re
from dataclasses import dataclass

from vba_addin_editor.adapters.ooxml_package_adapter import OoxmlPackageAdapter
from vba_addin_editor.domain.document import DocumentDraft, clone_draft
from vba_addin_editor.domain.session import DocumentSession
from vba_addin_editor.services.validation_service import validate_draft

PAGE_SIZE = 200
_WORD_EXTRA = "_"


def _is_word_char(ch: str) -> bool:
    return ch.isalnum() or ch == _WORD_EXTRA


@dataclass(frozen=True)
class SearchHit:
    target_id: str
    path: str
    kind: str  # vba | xml
    line: int
    column: int
    offset: int
    length: int
    snippet: str


@dataclass(frozen=True)
class SearchResults:
    query: str
    match_case: bool
    whole_word: bool
    include_xml: bool
    revision: int
    total: int
    hits: tuple[SearchHit, ...]


@dataclass
class ReplaceApplyResult:
    ok: bool
    changed: int = 0
    draft: DocumentDraft | None = None
    reason: str | None = None
    problems: tuple[str, ...] = ()
    before_modules: list | None = None
    after_modules: list | None = None
    before_xml: list | None = None
    after_xml: list | None = None


@dataclass(frozen=True)
class ProcedureInfo:
    module_id: str
    name: str
    kind: str
    visibility: str
    is_static: bool
    line: int
    signature: str


class SearchService:
    def search(
        self,
        draft: DocumentDraft,
        query: str,
        *,
        match_case: bool = False,
        whole_word: bool = False,
        include_xml: bool = False,
        revision: int = 0,
    ) -> SearchResults:
        if not query:
            return SearchResults(query, match_case, whole_word, include_xml, revision, 0, ())
        hits: list[SearchHit] = []
        for module in draft.final_module_state():
            hits.extend(
                self._scan(
                    module.body,
                    query,
                    match_case=match_case,
                    whole_word=whole_word,
                    target_id=module.id,
                    path=module.current_name,
                    kind="vba",
                )
            )
        if include_xml:
            for part in draft.xml_parts:
                if not part.editable or part.text is None:
                    continue
                hits.extend(
                    self._scan(
                        part.text,
                        query,
                        match_case=match_case,
                        whole_word=whole_word,
                        target_id=part.path,
                        path=part.path,
                        kind="xml",
                    )
                )
        return SearchResults(
            query=query,
            match_case=match_case,
            whole_word=whole_word,
            include_xml=include_xml,
            revision=revision,
            total=len(hits),
            hits=tuple(hits),
        )

    def page(self, results: SearchResults, page: int = 0) -> tuple[SearchHit, ...]:
        start = page * PAGE_SIZE
        return results.hits[start : start + PAGE_SIZE]

    def preview_replacements(
        self,
        draft: DocumentDraft,
        results: SearchResults,
        replacement: str,
        selected: set[tuple[str, int]],
    ) -> list[tuple[SearchHit, str, str]]:
        """Return (hit, before, after) for selected occurrences."""
        rows = []
        for hit in results.hits:
            if (hit.target_id, hit.offset) not in selected:
                continue
            text = self._text_for(draft, hit)
            if text is None:
                continue
            before = text[hit.offset : hit.offset + hit.length]
            after = replacement
            rows.append((hit, before, after))
        return rows

    def apply_replacements(
        self,
        session: DocumentSession,
        results: SearchResults,
        replacement: str,
        selected: set[tuple[str, int]],
        *,
        package_adapter: OoxmlPackageAdapter | None = None,
    ) -> ReplaceApplyResult:
        if session.revision != results.revision:
            return ReplaceApplyResult(ok=False, reason="stale_proposal", problems=("Search results are out of date.",))
        from vba_addin_editor.services.history_service import snapshot_modules, snapshot_xml

        clone = clone_draft(session.draft)
        before_modules = snapshot_modules(session.draft)
        before_xml = snapshot_xml(session.draft)
        changed = self.replace_all(clone, results, replacement, selected)
        if changed == 0:
            return ReplaceApplyResult(ok=True, changed=0, draft=clone)
        problems = list(validate_draft(clone))
        adapter = package_adapter or OoxmlPackageAdapter()
        for part in clone.changed_xml_parts():
            problems.extend(adapter.validate_draft_part(part))
        if problems:
            return ReplaceApplyResult(
                ok=False,
                reason="invalid_draft",
                problems=tuple(problems),
            )
        return ReplaceApplyResult(
            ok=True,
            changed=changed,
            draft=clone,
            before_modules=before_modules,
            after_modules=snapshot_modules(clone),
            before_xml=before_xml,
            after_xml=snapshot_xml(clone),
        )

    def _text_for(self, draft: DocumentDraft, hit: SearchHit) -> str | None:
        if hit.kind == "vba":
            module = draft.module_by_id(hit.target_id)
            return None if module is None else module.body
        part = draft.xml_part_by_path(hit.target_id)
        if part is None or part.text is None:
            return None
        return part.text

    def replace_all(
        self,
        draft: DocumentDraft,
        results: SearchResults,
        replacement: str,
        selected: set[tuple[str, int]],
    ) -> int:
        """Apply literal replacements descending per target. Returns replacements made."""
        by_target: dict[str, list[SearchHit]] = {}
        for hit in results.hits:
            if (hit.target_id, hit.offset) not in selected:
                continue
            by_target.setdefault(hit.target_id, []).append(hit)
        changed = 0
        for target_id, hits in by_target.items():
            module = draft.module_by_id(target_id)
            if module is not None:
                module.body = self._replace_text(module.body, hits, replacement)
                changed += len(hits)
                continue
            part = draft.xml_part_by_path(target_id)
            if part is not None and part.editable and part.text is not None:
                part.text = self._replace_text(part.text, hits, replacement)
                changed += len(hits)
        return changed

    def _replace_text(self, text: str, hits: list[SearchHit], replacement: str) -> str:
        ordered = sorted(hits, key=lambda item: item.offset, reverse=True)
        used: list[tuple[int, int]] = []
        for hit in ordered:
            span = (hit.offset, hit.offset + hit.length)
            if any(not (span[1] <= left or span[0] >= right) for left, right in used):
                continue
            text = text[: hit.offset] + replacement + text[hit.offset + hit.length :]
            used.append(span)
        return text

    def _scan(
        self,
        text: str,
        query: str,
        *,
        match_case: bool,
        whole_word: bool,
        target_id: str,
        path: str,
        kind: str,
    ) -> list[SearchHit]:
        haystack = text if match_case else text.lower()
        needle = query if match_case else query.lower()
        hits: list[SearchHit] = []
        start = 0
        while True:
            idx = haystack.find(needle, start)
            if idx < 0:
                break
            if whole_word and not self._whole_word(text, idx, len(query)):
                start = idx + 1
                continue
            line = text.count("\n", 0, idx) + 1
            line_start = text.rfind("\n", 0, idx) + 1
            column = idx - line_start + 1
            line_end = text.find("\n", idx)
            snippet = text[line_start : line_end if line_end >= 0 else line_start + 120]
            hits.append(
                SearchHit(
                    target_id=target_id,
                    path=path,
                    kind=kind,
                    line=line,
                    column=column,
                    offset=idx,
                    length=len(query),
                    snippet=snippet.strip()[:200],
                )
            )
            start = idx + max(len(query), 1)
        return hits

    def _whole_word(self, text: str, idx: int, length: int) -> bool:
        left = text[idx - 1] if idx > 0 else ""
        right = text[idx + length] if idx + length < len(text) else ""
        if left and _is_word_char(left):
            return False
        return not (right and _is_word_char(right))

    def outline(self, body: str, module_id: str) -> list[ProcedureInfo]:
        return parse_procedures(body, module_id)


_PROC_KINDS = ("Sub", "Function", "Property")
_VIS = ("Public", "Private", "Friend")


def parse_procedures(body: str, module_id: str) -> list[ProcedureInfo]:
    """Lexical procedure outline; skips comments and strings; joins continuations."""
    logical = _join_continuations(_strip_comments_and_strings(body))
    results: list[ProcedureInfo] = []
    pattern = re.compile(
        r"^(?P<vis>Public|Private|Friend)?\s*(?P<stat>Static\s+)?(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?P<name>[A-Za-z_][A-Za-z0-9_]*)",
        re.IGNORECASE,
    )
    for line_no, line in enumerate(logical.splitlines(), start=1):
        stripped = line.strip()
        if stripped.lower().startswith("end "):
            continue
        match = pattern.match(stripped)
        if not match:
            continue
        kind = re.sub(r"\s+", " ", match.group("kind")).title().replace("Property ", "Property ")
        results.append(
            ProcedureInfo(
                module_id=module_id,
                name=match.group("name"),
                kind=kind,
                visibility=(match.group("vis") or "Public"),
                is_static=bool(match.group("stat")),
                line=line_no,
                signature=stripped,
            )
        )
    return results


def _join_continuations(text: str) -> str:
    lines = text.split("\n")
    out: list[str] = []
    buf = ""
    for line in lines:
        stripped = line.rstrip()
        if stripped.endswith(" _"):
            buf += stripped[:-1]
            continue
        out.append(buf + stripped)
        buf = ""
    if buf:
        out.append(buf)
    return "\n".join(out)


def _strip_comments_and_strings(text: str) -> str:
    """Replace string and comment contents with spaces so they cannot match."""
    out: list[str] = []
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch == '"':
            out.append('"')
            i += 1
            while i < n and text[i] != '"':
                out.append(" ")
                i += 1
            if i < n:
                out.append('"')
                i += 1
            continue
        if ch == "'":
            while i < n and text[i] != "\n":
                out.append(" ")
                i += 1
            continue
        # REM comment at token start of line
        out.append(ch)
        i += 1
    return "".join(out)
