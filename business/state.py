from dataclasses import dataclass, field
from typing import List, Optional, Any


@dataclass
class AppState:
    single_file: Optional[str] = None
    single_config: Optional[Any] = None
    batch_files: List[str] = field(default_factory=list)
    batch_mappings: List[Any] = field(default_factory=list)
