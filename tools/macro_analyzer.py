#!/usr/bin/env python3

import argparse
import hashlib
import json
import re
from datetime import datetime
from pathlib import Path

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    VBA_Parser = None


SCHEMA_VERSION = "1.1"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()

    with path.open("rb") as f:
        while True:
            chunk = f.read(1024 * 1024)

            if not chunk:
                break

            digest.update(chunk)

    return digest.hexdigest()


def make_error(error_type: str, message: str) -> dict:
    return {
        "type": error_type,
        "message": message,
    }


def make_evidence(
    evidence_id: str,
    module: str,
    location: str,
    evidence_type: str,
    description: str,
    source: str | None = None,
) -> dict:
    evidence = {
        "evidence_id": evidence_id,
        "module": module,
        "location": location,
        "type": evidence_type,
        "description": description,
    }

    if source is not None:
        evidence["source"] = source

    return evidence


def detect_module_type(module_name: str, stream_path: str) -> str:
    name = module_name.lower()

    if name == "thisdocument":
        return "document"

    if name.startswith("sheet"):
        return "worksheet"

    if name.startswith("thisworkbook"):
        return "workbook"

    if "class" in name:
        return "class"

    if stream_path.lower().endswith(".bas"):
        return "standard"

    return "standard"


def normalize_source(source) -> str:
    if source is None:
        return ""

    if isinstance(source, bytes):
        return source.decode("utf-8", errors="replace")

    return str(source)


def add_pattern_evidence(
    source: str,
    module_name: str,
    location: str,
    evidence: list,
    counter: list,
) -> None:
    patterns = [
        {
            "type": "auto_execution",
            "description": "Phát hiện AutoExec/AutoOpen/Document_Open hoặc sự kiện tự động thực thi",
            "patterns": [
                r"\bAutoOpen\b",
                r"\bAutoClose\b",
                r"\bAutoExec\b",
                r"\bDocument_Open\b",
                r"\bDocument_Close\b",
                r"\bWorkbook_Open\b",
                r"\bWorkbook_BeforeClose\b",
                r"\bAutoNew\b",
            ],
        },
        {
            "type": "command_execution",
            "description": "Phát hiện API hoặc hàm có khả năng thực thi command",
            "patterns": [
                r"\bShell\s*\(",
                r"\bWScript\.Shell\b",
                r"\bShellExecute\b",
                r"\bCreateProcess\b",
                r"\bExec\s*\(",
            ],
        },
        {
            "type": "powershell",
            "description": "Phát hiện PowerShell",
            "patterns": [
                r"\bpowershell(?:\.exe)?\b",
                r"\bPowerShell\b",
                r"-enc(?:odedcommand)?\b",
            ],
        },
        {
            "type": "process_creation",
            "description": "Phát hiện CreateObject/GetObject có thể được sử dụng để tạo COM object",
            "patterns": [
                r"\bCreateObject\s*\(",
                r"\bGetObject\s*\(",
            ],
        },
        {
            "type": "network",
            "description": "Phát hiện thành phần liên quan đến network/HTTP",
            "patterns": [
                r"\bMSXML2\.XMLHTTP\b",
                r"\bMSXML2\.ServerXMLHTTP\b",
                r"\bWinHttp\.WinHttpRequest\b",
                r"\bXMLHTTP\b",
                r"\bWinHttpRequest\b",
                r"\bADODB\.Stream\b",
            ],
        },
        {
            "type": "file_operation",
            "description": "Phát hiện thao tác file",
            "patterns": [
                r"\bOpen\s+.+\s+For\s+(?:Binary|Input|Output|Append)",
                r"\bKill\s+",
                r"\bFileCopy\s*\(",
                r"\bName\s+.+\s+As\s+",
                r"\bSaveToFile\s*\(",
            ],
        },
        {
            "type": "registry",
            "description": "Phát hiện thao tác Windows Registry",
            "patterns": [
                r"\bRegWrite\b",
                r"\bRegRead\b",
                r"\bRegDelete\b",
                r"\bHKEY_CURRENT_USER\b",
                r"\bHKEY_LOCAL_MACHINE\b",
            ],
        },
        {
            "type": "url",
            "description": "Phát hiện URL trong VBA source",
            "patterns": [
                r"https?://[^\s\"']+",
                r"ftp://[^\s\"']+",
            ],
        },
        {
            "type": "obfuscation",
            "description": "Phát hiện một số kỹ thuật thường được dùng để làm rối VBA",
            "patterns": [
                r"\bChr\s*\(",
                r"\bChrW\s*\(",
                r"\bAsc\s*\(",
                r"\bStrReverse\s*\(",
                r"\bReplace\s*\(",
                r"\bXor\b",
            ],
        },
    ]

    for item in patterns:
        matched = False
        matched_text = []

        for pattern in item["patterns"]:
            matches = re.findall(
                pattern,
                source,
                flags=re.IGNORECASE,
            )

            if matches:
                matched = True

                for match in matches[:5]:
                    if isinstance(match, tuple):
                        match = match[0]

                    matched_text.append(str(match))

        if not matched:
            continue

        evidence_id = f"E{counter[0]:03d}"
        counter[0] += 1

        source_text = ", ".join(
            dict.fromkeys(matched_text)
        )

        add_pattern = item["description"]

        if source_text:
            add_pattern += f"; pattern: {source_text}"

        evidence.append(
            make_evidence(
                evidence_id=evidence_id,
                module=module_name,
                location=location,
                evidence_type=item["type"],
                description=add_pattern,
            )
        )


def extract_vba(path: Path):
    if VBA_Parser is None:
        raise RuntimeError(
            "Missing dependency 'oletools'. "
            "Install it with: pip install oletools"
        )

    parser = VBA_Parser(str(path))

    modules = []
    evidence = []

    evidence_counter = [1]

    try:
        has_vba = parser.detect_vba_macros()

        if not has_vba:
            return modules, evidence, False

        for (
            filename,
            stream_path,
            vba_filename,
            source_code,
        ) in parser.extract_macros():
            source = normalize_source(source_code)

            module_name = vba_filename or Path(
                stream_path
            ).name

            module_type = detect_module_type(
                module_name,
                stream_path,
            )

            location = stream_path

            source_status = (
                "extracted"
                if source.strip()
                else "empty"
            )

            module_evidence_id = f"E{evidence_counter[0]:03d}"
            evidence_counter[0] += 1

            modules.append(
                {
                    "module": module_name,
                    "type": module_type,
                    "location": location,
                    "evidence_id": module_evidence_id,
                    "source_extracted": bool(source.strip()),
                    "source_status": source_status,
                    "source": source,
                }
            )

            evidence.append(
                make_evidence(
                    evidence_id=module_evidence_id,
                    module=module_name,
                    location=location,
                    evidence_type="module_presence",
                    description=(
                        "Phát hiện và trích xuất module VBA"
                        if source.strip()
                        else
                        "Phát hiện module VBA nhưng source rỗng"
                    ),
                )
            )

            if source.strip():
                add_pattern_evidence(
                    source=source,
                    module_name=module_name,
                    location=location,
                    evidence=evidence,
                    counter=evidence_counter,
                )

        return modules, evidence, True

    finally:
        parser.close()


def analyze(path: Path) -> dict:
    errors = []

    if not path.exists():
        return {
            "schema_version": SCHEMA_VERSION,
            "input": {
                "filename": path.name,
                "sha256": None,
            },
            "analysis": {
                "status": "error",
                "scope": "vba_static_analysis",
                "has_macro": None,
                "modules": [],
            },
            "evidence": [],
            "errors": [
                make_error(
                    "FileNotFoundError",
                    f"Input file does not exist: {path}",
                )
            ],
        }

    if not path.is_file():
        return {
            "schema_version": SCHEMA_VERSION,
            "input": {
                "filename": path.name,
                "sha256": None,
            },
            "analysis": {
                "status": "error",
                "scope": "vba_static_analysis",
                "has_macro": None,
                "modules": [],
            },
            "evidence": [],
            "errors": [
                make_error(
                    "NotAFile",
                    f"Input path is not a file: {path}",
                )
            ],
        }

    try:
        file_hash = sha256_file(path)

    except Exception as exc:
        return {
            "schema_version": SCHEMA_VERSION,
            "input": {
                "filename": path.name,
                "sha256": None,
            },
            "analysis": {
                "status": "error",
                "scope": "vba_static_analysis",
                "has_macro": None,
                "modules": [],
            },
            "evidence": [],
            "errors": [
                make_error(
                    type(exc).__name__,
                    str(exc),
                )
            ],
        }

    try:
        modules, evidence, has_macro = extract_vba(path)

    except Exception as exc:
        return {
            "schema_version": SCHEMA_VERSION,
            "input": {
                "filename": path.name,
                "sha256": file_hash,
            },
            "analysis": {
                "status": "error",
                "scope": "vba_static_analysis",
                "has_macro": None,
                "modules": [],
            },
            "evidence": [],
            "errors": [
                make_error(
                    type(exc).__name__,
                    str(exc),
                )
            ],
        }

    return {
        "schema_version": SCHEMA_VERSION,
        "input": {
            "filename": path.name,
            "sha256": file_hash,
        },
        "analysis": {
            "status": "success",
            "scope": "vba_static_analysis",
            "has_macro": has_macro,
            "module_count": len(modules),
            "modules": modules,
        },
        "evidence": evidence,
        "errors": errors,
    }


def build_auto_output_path(input_path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")

    return Path("results") / (
        f"{input_path.stem}_{timestamp}.json"
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Static analyzer for VBA Macro Office documents"
        )
    )

    parser.add_argument(
        "input",
        type=Path,
        help=(
            "Input Office document "
            "(.docm/.dotm/.xlsm/.xlam/.doc/.xls)"
        ),
    )

    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help=(
            "Output JSON file. If omitted, "
            "results/<tenfile>_<YYYYMMDD-HHMMSS>.json"
        ),
    )

    parser.add_argument(
        "--ai",
        action="store_true",
        help="Enable LLM-assisted analysis",
    )

    parser.add_argument(
        "--no-llm",
        action="store_true",
        help="Disable LLM-assisted analysis",
    )

    args = parser.parse_args()

    if args.ai and args.no_llm:
        parser.error(
            "--ai and --no-llm cannot be used together"
        )

    output = args.output or build_auto_output_path(args.input)

    result = analyze(args.input)

    output.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    output.write_text(
        json.dumps(
            result,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print(f"[+] Input      : {args.input}")
    print(
        f"[+] SHA256     : "
        f"{result['input']['sha256']}"
    )

    print(
        f"[+] Has VBA    : "
        f"{result['analysis']['has_macro']}"
    )

    print(
        f"[+] Module count: "
        f"{result['analysis'].get('module_count', 0)}"
    )

    print(f"[+] Output     : {output}")

    if result["errors"]:
        print("[!] Errors:")

        for error in result["errors"]:
            print(
                f"    {error['type']}: "
                f"{error['message']}"
            )

    return (
        0
        if result["analysis"]["status"] == "success"
        else 1
    )


if __name__ == "__main__":
    raise SystemExit(main())
