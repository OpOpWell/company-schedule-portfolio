"""
kintone 工事予定管理アプリへ Python から1件登録する最小デモ

使い方:
1. kintone側で「工事予定管理」アプリのAPIトークンを作成
   - 権限: レコード追加
2. PowerShellで環境変数を設定
   $env:KINTONE_BASE_URL="https://あなたのサブドメイン.cybozu.com"
   $env:KINTONE_APP_ID="アプリID"
   $env:KINTONE_API_TOKEN="APIトークン"
3. 実行
   python kintone_add_construction_record.py

注意:
- APIトークンはGitHubへ上げないでください。
- kintoneの「フィールドコード」が下の record のキーと違う場合は、record側をフィールドコードに合わせて直してください。
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


def add_record() -> dict[str, Any]:
    base_url = env_required("KINTONE_BASE_URL").rstrip("/")
    app_id = env_required("KINTONE_APP_ID")
    api_token = env_required("KINTONE_API_TOKEN")

    url = f"{base_url}/k/v1/record.json"

    headers = {
        "X-Cybozu-API-Token": api_token,
        "Content-Type": "application/json",
    }

    payload = {
        "app": app_id,
        "record": {
            "工事ID": {"value": "PY-001"},
            "工事名": {"value": "Python API連携テスト工事"},
            "現場": {"value": "APIテスト現場"},
            "依頼主": {"value": "サンプル依頼主"},
            "連絡先": {"value": "018-999-9999"},
            "契約金額": {"value": "API登録テスト"},
            "開始日": {"value": "2026-07-21"},
            "終了日": {"value": "2026-07-22"},
            "状態": {"value": "着工前"},
            "担当": {"value": ""},
            "通知": {"value": "要確認"},
            "備考": {"value": "Pythonからkintone REST APIで登録したテストレコード"},
        },
    }

    response = requests.post(url, headers=headers, json=payload, timeout=30)

    if not response.ok:
        print("登録に失敗しました。", file=sys.stderr)
        print(f"HTTP Status: {response.status_code}", file=sys.stderr)
        print(response.text, file=sys.stderr)
        response.raise_for_status()

    return response.json()


def main() -> None:
    try:
        result = add_record()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    print("kintoneへ登録しました。")
    print(f"record id: {result.get('id')}")
    print(f"revision : {result.get('revision')}")


if __name__ == "__main__":
    main()
