from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable


@dataclass
class ValidationArtifact:
    kind: str
    path: Path
    status: str = "GERADO"


@dataclass
class ValidationJob:
    network: str
    status: str
    created_at: str = field(default_factory=lambda: datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    artifacts: list[ValidationArtifact] = field(default_factory=list)
    alerts: list[str] = field(default_factory=list)

    def add_artifact(self, kind: str, path: str | Path, status: str = "GERADO") -> None:
        self.artifacts.append(ValidationArtifact(kind=kind, path=Path(path), status=status))

    def add_alerts(self, alerts: Iterable[str]) -> None:
        for alert in alerts:
            text = str(alert).strip()
            if text:
                self.alerts.append(text)

