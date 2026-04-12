# Contributing

Contributions are welcome. Please follow the guidelines below.

## Reporting Issues

If you encounter a CDX file that fails to parse, please open an issue and include:
- The error message and full traceback
- The version of Python, RDKit, and pycdxml you are using
- If possible, a minimal anonymised CDX file that reproduces the issue

## Submitting Pull Requests

1. Fork the repository and create a feature branch from `main`.
2. Make your changes with clear, descriptive commit messages.
3. Ensure your code follows the existing style (PEP 8, docstrings on all public functions).
4. Test your changes against the example CDX files in `examples/input/cdx/`.
5. Open a pull request with a clear description of what was changed and why.

## Known Limitations

- Very old CDX files (ChemDraw 7 era and earlier) may fail due to non-standard binary encoding.
- Structures containing reaction arrows, biological sequences, or polymer notation may not yield valid SMILES.
- OLE objects that are not CDX files (e.g., embedded Excel charts) are automatically skipped.
