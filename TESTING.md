# Running the Test Suite

These tests use `pytest`.

## 1. Create and activate a virtual environment (recommended)

```pwsh
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

## 2. Install dependencies

Install runtime dependencies:

```pwsh
pip install -r requirements.txt
```

Install dev/test dependencies:

```pwsh
pip install -r requirements-dev.txt
```

## 3. Run tests

From the repository root:

```pwsh
python -m pytest
```

If you prefer the `pytest` entrypoint and it resolves correctly in your environment:

```pwsh
pytest
```
