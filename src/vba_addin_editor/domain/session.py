"""Document session: captured baseline, draft, revision, and history."""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from pathlib import Path

from vba_addin_editor.domain.document import DocumentDraft, FileFingerprint
from vba_addin_editor.domain.history import UndoHistory


def new_session_id() -> str:
    return uuid.uuid4().hex


@dataclass(frozen=True)
class ProposalBinding:
    """Save/review/import/conflict proposals bind to this identity."""

    session_id: str
    revision: int
    baseline_sha256: str


@dataclass
class DocumentSession:
    """Application-owned editing session, separate from DocumentDraft.

    Dirty state remains a comparison against baseline, not revision inequality.
    Every accepted mutation, undo/redo, and baseline replacement increments
    revision. Stable module IDs are valid within one baseline generation.
    """

    session_id: str
    draft: DocumentDraft
    captured_path: Path
    captured_fingerprint: FileFingerprint
    session_dir: Path
    revision: int = 1
    last_checkpoint_revision: int = 0
    history: UndoHistory = field(default_factory=UndoHistory)
    original_exists: bool = True
    recovery_generation: str | None = None

    @property
    def original_path(self) -> Path:
        return self.draft.baseline.path

    @property
    def baseline_sha256(self) -> str:
        return self.captured_fingerprint.sha256

    def binding(self) -> ProposalBinding:
        return ProposalBinding(
            session_id=self.session_id,
            revision=self.revision,
            baseline_sha256=self.baseline_sha256,
        )

    def matches_proposal(self, binding: ProposalBinding) -> bool:
        return (
            binding.session_id == self.session_id
            and binding.revision == self.revision
            and binding.baseline_sha256 == self.baseline_sha256
        )

    def bump(self) -> int:
        self.revision += 1
        return self.revision
