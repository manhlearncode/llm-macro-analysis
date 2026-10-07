from pathlib import Path
import json

root = Path.cwd()
output_dir = root / "ai_trial" / "inputs"
output_dir.mkdir(parents=True, exist_ok=True)

# Đọc kết quả S01 thật từ repo.
raw = json.loads(
    (root / "results" / "S01-1.json").read_text(encoding="utf-8")
)
assert raw["schema_version"] == "1.1"
assert raw["analysis"]["status"] == "success" and not raw["errors"]

module = next(
    m for m in raw["analysis"]["modules"]
    if m["module"] == "Module1.bas"
)
lines = module["source"].splitlines()
assert lines[2].strip() == 'MsgBox "Hello"'

# L001 là bằng chứng trích đoạn nguồn cho phép thử LLM.
# E002 trong JSON gốc là bằng chứng hiện diện của module.
s01 = {
    "sample_id": "S01",
    "provenance": "extracted_office",
    "input_sha256": raw["input"]["sha256"],
    "extraction_status": "success",
    "evidence": [{
        "id": "L001",
        "origin_evidence_id": module["evidence_id"],
        "module": module["module"],
        "location": module["location"],
        "line_start": 2,
        "line_end": 4,
        "type": "vba_source_excerpt",
        "excerpt": "\n".join(lines[1:4]),
    }],
    "errors": raw["errors"],
}

def fixture(sample_id, excerpt):
    return {
        "sample_id": sample_id,
        "provenance": "synthetic_fixture",
        "input_sha256": None,
        "extraction_status": "fixture",
        "evidence": [{
            "id": "L001",
            "module": "Module1.bas",
            "location": "Module1.bas",
            "line_start": 1,
            "line_end": len(excerpt.splitlines()),
            "type": "vba_source_excerpt",
            "excerpt": excerpt,
        }],
        "errors": [],
    }

s02 = fixture(
    "S02-fixture",
    'Sub Demo()\n'
    '    Dim s As String\n'
    '    s = "Xin " & Chr(99) & "hao"\n'
    'End Sub'
)

s04 = fixture(
    "S04-fixture",
    'Sub Demo()\n'
    '    Dim target As String\n'
    '    target = "SayHello"\n'
    '    Application.Run target\n'
    'End Sub'
)

s06 = {
    "sample_id": "S06-pending",
    "provenance": "synthetic_fixture",
    "input_sha256": None,
    "extraction_status": "not_run",
    "evidence": [],
    "errors": [],
}

for sample in (s01, s02, s04, s06):
    path = output_dir / f'{sample["sample_id"]}.json'
    path.write_text(
        json.dumps(sample, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(path)