from multiprocessing import freeze_support
from pathlib import Path
import sys

from markitdown_gui.converter import convert_files
from markitdown_gui.app import main


if __name__ == "__main__":
    freeze_support()
    if len(sys.argv) > 3 and sys.argv[1] == "--convert-smoke":
        output_dir = Path(sys.argv[2])
        sources = [Path(item) for item in sys.argv[3:]]
        results = convert_files(sources, output_dir)
        for result in results:
            status = "ok" if result.succeeded else f"failed: {result.error}"
            print(f"{result.source.name}: {status}")
        raise SystemExit(0 if all(result.succeeded for result in results) else 1)
    main()
