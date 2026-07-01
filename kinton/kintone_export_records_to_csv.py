"""
kintone → CSVバックアップ用 Pythonデモ

目的:
- kintoneアプリのレコードをPythonで取得
- フィールドコードではなく、画面表示名に近い列名でCSV保存
- データ移行・バックアップ・集計を想定したデータ連携デモにする

対応:
- APIトークン認証
- ログイン名 + パスワード認証
- 500件ずつ取得
- UTF-8 BOM付きCSVで保存（Excelで文字化けしにくい）
- JSONも同時保存

使い方 1: ログイン名 + パスワード認証（おすすめ）
  $env:KINTONE_BASE_URL="https://あなたのサブドメイン.cybozu.com"
  $env:KINTONE_APP_ID="6"
  $env:KINTONE_LOGIN_NAME="あなたのkintoneログイン名"
  $env:KINTONE_EXPORT_NAME="工事予定管理"
  python .\kintone_export_records_to_csv.py

使い方 2: APIトークン認証
  $env:KINTONE_BASE_URL="https://あなたのサブドメイン.cybozu.com"
  $env:KINTONE_APP_ID="6"
  $env:KINTONE_API_TOKEN="APIトークン"
  $env:KINTONE_EXPORT_NAME="工事予定管理"
  python .\kintone_export_records_to_csv.py

任意設定:
  $env:KINTONE_QUERY='通知 in ("要確認") order by レコード番号 desc'

注意:
- APIトークン、ログイン名、パスワードはGitHubに上げないでください。
- PowerShellの環境変数にパスワードを入れた場合は、作業後に削除してください。
"""

from __future__ import annotations

import base64
import csv
import getpass
import json
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import requests


@dataclass
class KintoneAuth:
    mode: str
    api_token: str | None = None
    login_name: str | None = None
    password: str | None = None

    def headers(self, has_body: bool = False) -> dict[str, str]:
        headers: dict[str, str] = {}

        if self.mode == "password":
            if not self.login_name or self.password is None:
                raise RuntimeError("パスワード認証に必要なログイン名またはパスワードがありません。")
            raw = f"{self.login_name}:{self.password}".encode("utf-8")
            headers["X-Cybozu-Authorization"] = base64.b64encode(raw).decode("ascii")
        elif self.mode == "api_token":
            if not self.api_token:
                raise RuntimeError("APIトークン認証に必要なAPIトークンがありません。")
            headers["X-Cybozu-API-Token"] = self.api_token
        else:
            raise RuntimeError(f"不明な認証モードです: {self.mode}")

        if has_body:
            headers["Content-Type"] = "application/json"

        return headers


def env_required(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"環境変数 {name} が未設定です。")
    return value


def build_auth() -> KintoneAuth:
    login_name = os.environ.get("KINTONE_LOGIN_NAME", "").strip()
    password = os.environ.get("KINTONE_PASSWORD", "")
    api_token = os.environ.get("KINTONE_API_TOKEN", "").strip()

    if login_name:
        if not password:
            password = getpass.getpass("kintone password: ")
        return KintoneAuth(mode="password", login_name=login_name, password=password)

    if api_token:
        return KintoneAuth(mode="api_token", api_token=api_token)

    raise RuntimeError(
        "認証情報がありません。KINTONE_LOGIN_NAME または KINTONE_API_TOKEN を設定してください。"
    )


def request_json(method: str, url: str, auth: KintoneAuth, **kwargs: Any) -> dict[str, Any]:
    has_body = "json" in kwargs or "data" in kwargs
    response = requests.request(
        method,
        url,
        headers=auth.headers(has_body=has_body),
        timeout=60,
        **kwargs,
    )

    if not response.ok:
        print(f"API実行に失敗しました: {method} {url}", file=sys.stderr)
        print(f"認証モード: {auth.mode}", file=sys.stderr)
        print(f"HTTP Status: {response.status_code}", file=sys.stderr)
        print(response.text, file=sys.stderr)
        response.raise_for_status()

    if response.text.strip():
        return response.json()
    return {}


def safe_filename(name: str) -> str:
    name = name.strip() or "kintone_export"
    return re.sub(r'[\\/:*?"<>|]+', "_", name)


def get_field_map(base_url: str, app_id: str, auth: KintoneAuth) -> dict[str, dict[str, Any]]:
    """戻り値: field_code -> {label, type}"""
    url = f"{base_url}/k/v1/app/form/fields.json"
    data = request_json("GET", url, auth, params={"app": app_id})
    properties = data.get("properties", {})

    field_map: dict[str, dict[str, Any]] = {}
    print(f"app={app_id} フィールド情報を取得しました。")

    for code, info in properties.items():
        label = str(info.get("label", "")).strip() or code
        field_type = str(info.get("type", ""))
        field_map[code] = {"label": label, "type": field_type}

    # システムフィールドの表示名
    field_map.setdefault("$id", {"label": "レコード番号", "type": "RECORD_NUMBER"})
    field_map.setdefault("$revision", {"label": "リビジョン", "type": "__REVISION__"})

    return field_map


def flatten_value(value: Any) -> str:
    """kintoneの値をCSV向けの文字列へ変換する。"""
    if value is None:
        return ""

    if isinstance(value, str):
        return value

    if isinstance(value, (int, float, bool)):
        return str(value)

    if isinstance(value, list):
        parts: list[str] = []
        for item in value:
            if isinstance(item, dict):
                # ユーザー選択・組織選択・添付ファイルなどをなるべく見やすくする
                if "name" in item:
                    parts.append(str(item.get("name", "")))
                elif "code" in item:
                    parts.append(str(item.get("code", "")))
                else:
                    parts.append(json.dumps(item, ensure_ascii=False))
            else:
                parts.append(str(item))
        return "; ".join(parts)

    if isinstance(value, dict):
        if "name" in value:
            return str(value.get("name", ""))
        if "code" in value:
            return str(value.get("code", ""))
        return json.dumps(value, ensure_ascii=False)

    return str(value)


def fetch_records(base_url: str, app_id: str, auth: KintoneAuth, query: str = "") -> list[dict[str, Any]]:
    url = f"{base_url}/k/v1/records.json"

    all_records: list[dict[str, Any]] = []
    limit = 500
    offset = 0

    while True:
        query_parts: list[str] = []
        if query.strip():
            query_parts.append(query.strip())
        query_parts.append(f"limit {limit}")
        query_parts.append(f"offset {offset}")
        full_query = " ".join(query_parts)

        data = request_json("GET", url, auth, params={"app": app_id, "query": full_query})
        records = data.get("records", [])
        all_records.extend(records)

        print(f"取得中: {len(all_records)} 件")

        if len(records) < limit:
            break

        offset += limit

        if offset >= 10000:
            print("注意: offset 10000件に到達しました。大量データではカーソルAPI化を検討してください。")
            break

    return all_records


def record_to_row(record: dict[str, Any], field_codes: list[str]) -> dict[str, str]:
    row: dict[str, str] = {}

    for code in field_codes:
        cell = record.get(code, {})
        if isinstance(cell, dict) and "value" in cell:
            row[code] = flatten_value(cell.get("value"))
        else:
            row[code] = flatten_value(cell)

    return row


def build_csv_headers(field_codes: list[str], field_map: dict[str, dict[str, Any]]) -> tuple[list[str], dict[str, str]]:
    """CSV列名を作る。同じ表示名がある場合はフィールドコードを付けて重複回避。"""
    label_counts: dict[str, int] = {}
    for code in field_codes:
        label = field_map.get(code, {}).get("label", code)
        label_counts[label] = label_counts.get(label, 0) + 1

    headers: list[str] = []
    code_to_header: dict[str, str] = {}

    for code in field_codes:
        label = field_map.get(code, {}).get("label", code)
        if label_counts.get(label, 0) > 1:
            header = f"{label}__{code}"
        else:
            header = label
        headers.append(header)
        code_to_header[code] = header

    return headers, code_to_header


def save_backup(
    records: list[dict[str, Any]],
    field_map: dict[str, dict[str, Any]],
    export_name: str,
    app_id: str,
    query: str,
) -> tuple[Path, Path]:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = safe_filename(f"{export_name}_app{app_id}_{timestamp}")

    backup_dir = Path("backup")
    backup_dir.mkdir(exist_ok=True)

    csv_path = backup_dir / f"{base_name}.csv"
    json_path = backup_dir / f"{base_name}.json"

    # レコード内に実際に存在するフィールドコードを順序付きで集める。
    preferred_system = ["$id", "$revision", "作成日時", "更新日時", "作成者", "更新者"]
    seen: set[str] = set()
    field_codes: list[str] = []

    for code in preferred_system:
        if any(code in record for record in records):
            field_codes.append(code)
            seen.add(code)

    for code in field_map.keys():
        if code not in seen and any(code in record for record in records):
            field_codes.append(code)
            seen.add(code)

    for record in records:
        for code in record.keys():
            if code not in seen:
                field_codes.append(code)
                seen.add(code)

    headers, code_to_header = build_csv_headers(field_codes, field_map)

    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()

        for record in records:
            raw_row = record_to_row(record, field_codes)
            row = {code_to_header[code]: raw_row.get(code, "") for code in field_codes}
            writer.writerow(row)

    payload = {
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "app_id": app_id,
        "export_name": export_name,
        "query": query,
        "count": len(records),
        "field_map": field_map,
        "records": records,
    }

    with json_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return csv_path, json_path


def main() -> None:
    try:
        base_url = env_required("KINTONE_BASE_URL").rstrip("/")
        app_id = env_required("KINTONE_APP_ID")
        export_name = os.environ.get("KINTONE_EXPORT_NAME", "").strip() or f"kintone_app_{app_id}"
        query = os.environ.get("KINTONE_QUERY", "").strip()

        auth = build_auth()

        print(f"認証モード: {auth.mode}")
        print(f"バックアップ対象 app={app_id} name={export_name}")

        field_map = get_field_map(base_url, app_id, auth)
        records = fetch_records(base_url, app_id, auth, query=query)

        csv_path, json_path = save_backup(records, field_map, export_name, app_id, query=query)

        print("バックアップ完了")
        print(f"CSV : {csv_path}")
        print(f"JSON: {json_path}")
        print(f"件数: {len(records)}")

    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
