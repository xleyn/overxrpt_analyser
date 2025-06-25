from pathlib import Path
import json
import sys
import time


class FileManager:
    """Class for managing I/O operations."""

    if getattr(sys, "frozen", False):
        project_dir = Path(sys.executable).parent
    else:
        project_dir = Path(__file__).parent.parent

    with open(project_dir.joinpath("config/file_structure.json")) as f:
        config = json.load(f)
        paths_from_proj_dir = dict(
            zip(
                config.keys(),
                map(
                    lambda v, project_dir=project_dir: (
                        Path(v) if Path(v).is_absolute() else project_dir / v
                    ),
                    config.values(),
                ),
            )
        )

    @classmethod
    def creation_control(cls):
        """Checks if I/O paths exist. Exits programme if not, rectifying the issues."""
        need_to_quit = False
        for path in cls.paths_from_proj_dir:
            if not path.exists:
                print(f"Does not exist: {path}")
                need_to_quit = True
        if need_to_quit:
            print("Please rectify issues. Quitting application.")
            time.sleep(3)
            sys.exit()
        else:
            print("All I/O paths exist, as expected!")
