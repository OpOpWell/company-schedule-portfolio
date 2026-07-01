"""
kintone 工事予定管理: フィールド名からフィールドコードを自動取得して1件登録するデモ

目的:
- kintoneのフィールドコードを手入力で直さずに、
  フォーム情報から「フィールド名 -> フィールドコード」を取得して登録する。
- CSVから作ったアプリでフィールドコードが自動名になっていても動かしやすくする。

使い方:
PowerShellで以下を設定してから実行してください。

$env:KINTONE_BASE_URL="https://あなたのサブドメイン.cybozu.com"
$env:KINTONE_APP_ID="アプリID"
$env:KINTONE_API_TOKEN="APIトークン"

python kintone_add_by_label.py

APIトークン権限:
- レコード追加
- レコード閲覧
- アプリ管理
があると安全です。
保存後は必ず「アプリを更新」してください。

注意:
- APIトークンはGitHubに上げないでください。
- このスクリプトはフィールドコードを変更しません。
  現在のフィールド名に対応するフィールドコードを読んで、そのコードで登録します。
"""

from __future__ import annotations

import os
import sys
from typing import Any

import requests


def env_required(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"環境変数 {name} が未設定です。")
    return value


def get_field_map(base_url: str, app_id: str, api_token: str) -> dict[str, str]:
    url = f"{base_url}/k/v1/app/form/fields.json"
    response = requests.get(
        url,
        headers={"X-Cybozu-API-Token": api_token},
        params={"app": app_id},
        timeout=30,
    )

    if not response.ok:
        print("フィールド情報の取得に失敗しました。", file=sys.stderr)
        print(f"HTTP Status: {response.status_code}", file=sys.stderr)
        print(response.text, file=sys.stderr)
        response.raise_for_status()

    properties = response.json().get("properties", {})
    field_map: dict[str, str] = {}

    print("フィールド名 -> フィールドコード:")
    for code, info in properties.items():
        label = str(info.get("label", "")).strip()
        field_type = info.get("type", "")
        if label:
            field_map[label] = code
            print(f"{label} -> {code} ({field_type})")

    return field_map


def build_record_by_label(field_map: dict[str, str]) -> dict[str, dict[str, Any]]:
    values_by_label = {
        "工事ID": "PY-001",
        "工事名": "Python API連携テスト工事",
        "現場": "APIテスト現場",
        "依頼主": "サンプル依頼主",
        "連絡先": "018-999-9999",
        "契約金額": "API登録テスト",
        "開始日": "2026-07-21",
        "終了日": "2026-07-22",
        "状態": "着工前",
        "担当": "",
        "通知": "要確認",
        "備考": "Pythonからフィールドコードを自動取得して登録したテストレコード",
    }

    record: dict[str, dict[str, Any]] = {}
    missing_labels: list[str] = []

    for label, value in values_by_label.items():
        code = field_map.get(label)
        if not code:
            missing_labels.append(label)
            continue

        if value == "":
            continue

        record[code] = {"value": value}

    if missing_labels:
        print("見つからなかったフィールド名:")
        for label in missing_labels:
            print(f"- {label}")

    return record


def add_record(base_url: str, app_id: str, api_token: str, record: dict[str, dict[str, Any]]) -> dict[str, Any]:
    url = f"{base_url}/k/v1/record.json"
    payload = {"app": app_id, "record": record}

    response = requests.post(
        url,
        headers={
            "X-Cybozu-API-Token": api_token,
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )

    if not response.ok:
        print("レコード登録に失敗しました。", file=sys.stderr)
        print(f"HTTP Status: {response.status_code}", file=sys.stderr)
        print(response.text, file=sys.stderr)
        response.raise_for_status()

    return response.json()


def main() -> None:
    try:
        base_url = env_required("KINTONE_BASE_URL").rstrip("/")
        app_id = env_required("KINTONE_APP_ID")
        api_token = env_required("KINTONE_API_TOKEN")

        field_map = get_field_map(base_url, app_id, api_token)
        record = build_record_by_label(field_map)

        if not record:
            raise RuntimeError("登録するフィールドがありません。フィールド名とフォーム設定を確認してください。")

        print("送信するフィールドコード:")
        for code in record.keys():
            print(f"- {code}")

        result = add_record(base_url, app_id, api_token, record)

    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    print("kintoneへ登録しました。")
    print(f"record id: {result.get('id')}")
    print(f"revision : {result.get('revision')}")


if __name__ == "__main__":
    main()
