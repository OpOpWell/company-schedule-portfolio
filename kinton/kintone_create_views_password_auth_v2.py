"""
kintone 一覧ビュー作成 + 反映までPythonで行う版 v2

v2修正点:
- 反映状況確認APIのパラメータを apps[0].app から apps[0] に修正。

原因対策:
- APIトークンで PUT /k/v1/preview/app/views.json が 403 GAIA_NO01 になる場合、
  APIトークン権限またはカスタマイズ形式一覧の制限で止まっている可能性があります。
- この版は、kintoneのログイン名 + パスワード認証にも対応します。
- パスワードは環境変数に入れなくても、実行時に非表示入力できます。

作成する一覧:
- 工事予定管理アプリ:
  1. 未完了工事一覧
  2. 施工中一覧
  3. 要確認一覧

資格管理アプリIDを指定した場合:
  4. 資格更新期限あり一覧

使い方 1: パスワードを非表示入力する方法（おすすめ）
  $env:KINTONE_BASE_URL="https://あなたのサブドメイン.cybozu.com"
  $env:KINTONE_APP_ID="工事予定管理のアプリID"
  $env:KINTONE_LOGIN_NAME="あなたのkintoneログイン名"
  python .\kintone_create_views_password_auth_v2.py

使い方 2: パスワードも環境変数に入れる方法
  $env:KINTONE_BASE_URL="https://あなたのサブドメイン.cybozu.com"
  $env:KINTONE_APP_ID="工事予定管理のアプリID"
  $env:KINTONE_LOGIN_NAME="あなたのkintoneログイン名"
  $env:KINTONE_PASSWORD="あなたのkintoneパスワード"
  python .\kintone_create_views_password_auth_v2.py

資格管理も一緒に作る場合:
  $env:KINTONE_QUAL_APP_ID="資格管理アプリID"

注意:
- ログイン名・パスワード・APIトークンはGitHubに上げないでください。
- このスクリプトは一覧設定を更新したあと、アプリ設定の運用環境反映APIも実行します。
- 画面で「アプリを更新」を押す操作までPythonで行います。
"""

from __future__ import annotations

import base64
import getpass
import os
import sys
import time
from dataclasses import dataclass
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
        "認証情報がありません。KINTONE_LOGIN_NAME を設定するか、KINTONE_API_TOKEN を設定してください。"
    )


def request_json(method: str, url: str, auth: KintoneAuth, **kwargs: Any) -> dict[str, Any]:
    has_body = "json" in kwargs or "data" in kwargs
    response = requests.request(
        method,
        url,
        headers=auth.headers(has_body=has_body),
        timeout=30,
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


def get_field_map(base_url: str, app_id: str, auth: KintoneAuth) -> dict[str, str]:
    url = f"{base_url}/k/v1/app/form/fields.json"
    data = request_json("GET", url, auth, params={"app": app_id})
    properties = data.get("properties", {})

    field_map: dict[str, str] = {}
    print(f"app={app_id} フィールド名 -> フィールドコード")
    for code, info in properties.items():
        label = str(info.get("label", "")).strip()
        field_type = info.get("type", "")
        if label:
            field_map[label] = code
            print(f"  {label} -> {code} ({field_type})")
    return field_map


def get_existing_views(base_url: str, app_id: str, auth: KintoneAuth) -> dict[str, Any]:
    # 動作テスト環境の一覧を取得します。
    # これにより、未反映の一覧変更があってもなるべく維持できます。
    url = f"{base_url}/k/v1/preview/app/views.json"
    data = request_json("GET", url, auth, params={"app": app_id, "lang": "ja"})
    return data.get("views", {})


def clean_view_for_update(view: dict[str, Any]) -> dict[str, Any]:
    # GET結果には id や builtinType など、PUTに不要/不可の項目が混ざるため落とします。
    allowed_keys = {
        "index",
        "type",
        "name",
        "fields",
        "filterCond",
        "sort",
        "date",
        "title",
        "html",
        "pager",
        "device",
    }
    cleaned = {k: v for k, v in view.items() if k in allowed_keys}

    # LIST一覧に name がない場合はkintone側でエラーになることがあるため補完します。
    # ただしキー名が正式な一覧名なので、呼び出し側で補完します。
    return cleaned


def normalize_existing_views(existing: dict[str, Any]) -> dict[str, Any]:
    views: dict[str, Any] = {}
    for name, view in existing.items():
        cleaned = clean_view_for_update(view)
        if cleaned.get("type") in ("LIST", "CALENDAR", "CUSTOM") and "name" not in cleaned:
            cleaned["name"] = name
        views[name] = cleaned
    return views


def require_codes(field_map: dict[str, str], labels: list[str]) -> list[str]:
    codes: list[str] = []
    missing: list[str] = []
    for label in labels:
        code = field_map.get(label)
        if code:
            codes.append(code)
        else:
            missing.append(label)

    if missing:
        print("見つからなかったフィールド名は一覧から除外しました:")
        for label in missing:
            print(f"  - {label}")

    return codes


def put_views(base_url: str, app_id: str, auth: KintoneAuth, views: dict[str, Any]) -> str | None:
    url = f"{base_url}/k/v1/preview/app/views.json"
    payload = {"app": app_id, "views": views}
    result = request_json("PUT", url, auth, json=payload)
    revision = result.get("revision")
    print(f"app={app_id} 一覧設定を更新しました。revision={revision}")
    return revision


def deploy_app(base_url: str, app_id: str, auth: KintoneAuth, revision: str | None = None) -> None:
    url = f"{base_url}/k/v1/preview/app/deploy.json"
    app_payload: dict[str, Any] = {"app": app_id}
    if revision:
        app_payload["revision"] = revision

    payload = {"apps": [app_payload], "revert": False}

    try:
        request_json("POST", url, auth, json=payload)
        print(f"app={app_id} アプリ設定の反映を開始しました。")
    except requests.HTTPError:
        print("アプリ設定の反映APIで止まりました。画面の「アプリを更新」で反映してください。", file=sys.stderr)
        raise


def check_deploy_status(base_url: str, app_id: str, auth: KintoneAuth) -> None:
    # 反映は非同期なので、数回だけ確認します。
    url = f"{base_url}/k/v1/preview/app/deploy.json"

    for _ in range(10):
        time.sleep(1)
        data = request_json("GET", url, auth, params={"apps[0]": app_id})
        apps = data.get("apps", [])
        if not apps:
            continue

        status = apps[0].get("status")
        print(f"app={app_id} 反映状況: {status}")

        if status in ("SUCCESS", "FAIL", "CANCEL"):
            return

    print("反映状況確認はタイムアウトしました。kintone画面で一覧を確認してください。")


def add_construction_views(base_url: str, app_id: str, auth: KintoneAuth) -> str | None:
    field_map = get_field_map(base_url, app_id, auth)

    required = ["状態", "通知", "開始日"]
    missing_required = [label for label in required if label not in field_map]
    if missing_required:
        raise RuntimeError(f"工事予定管理アプリに必要なフィールドがありません: {', '.join(missing_required)}")

    state = field_map["状態"]
    notice = field_map["通知"]
    start = field_map["開始日"]
    end = field_map.get("終了日", start)

    existing = get_existing_views(base_url, app_id, auth)
    views = normalize_existing_views(existing)

    base_fields = require_codes(
        field_map,
        ["工事ID", "工事名", "現場", "依頼主", "開始日", "終了日", "状態", "通知", "担当"],
    )

    confirm_fields = require_codes(
        field_map,
        ["工事ID", "工事名", "現場", "依頼主", "状態", "通知", "開始日", "終了日", "備考"],
    )

    views["未完了工事一覧"] = {
        "index": "100",
        "type": "LIST",
        "name": "未完了工事一覧",
        "fields": base_fields,
        "filterCond": f'{state} not in ("完了", "請求済み", "中止")',
        "sort": f"{start} asc",
    }

    views["施工中一覧"] = {
        "index": "101",
        "type": "LIST",
        "name": "施工中一覧",
        "fields": base_fields,
        "filterCond": f'{state} in ("施工中")',
        "sort": f"{end} asc",
    }

    views["要確認一覧"] = {
        "index": "102",
        "type": "LIST",
        "name": "要確認一覧",
        "fields": confirm_fields,
        "filterCond": f'{notice} in ("要確認") or {state} in ("延期")',
        "sort": f"{start} asc",
    }

    return put_views(base_url, app_id, auth, views)


def add_qualification_views(base_url: str, app_id: str, auth: KintoneAuth) -> str | None:
    field_map = get_field_map(base_url, app_id, auth)

    deadline = field_map.get("更新期限")
    if not deadline:
        raise RuntimeError("資格管理アプリに「更新期限」フィールドが見つかりません。")

    existing = get_existing_views(base_url, app_id, auth)
    views = normalize_existing_views(existing)

    fields = require_codes(
        field_map,
        ["資格ID", "所有者", "区分", "資格名", "取得区分", "保有状況", "更新期限", "コピー有無", "状態", "通知", "備考"],
    )

    views["資格更新期限あり一覧"] = {
        "index": "100",
        "type": "LIST",
        "name": "資格更新期限あり一覧",
        "fields": fields,
        "filterCond": f'{deadline} != ""',
        "sort": f"{deadline} asc",
    }

    return put_views(base_url, app_id, auth, views)


def main() -> None:
    try:
        base_url = env_required("KINTONE_BASE_URL").rstrip("/")
        construction_app_id = env_required("KINTONE_APP_ID")
        auth = build_auth()

        print(f"認証モード: {auth.mode}")
        print("工事予定管理アプリの一覧を作成します。")
        construction_revision = add_construction_views(base_url, construction_app_id, auth)
        deploy_app(base_url, construction_app_id, auth, construction_revision)
        check_deploy_status(base_url, construction_app_id, auth)

        qual_app_id = os.environ.get("KINTONE_QUAL_APP_ID", "").strip()
        if qual_app_id:
            print("資格管理アプリの一覧を作成します。")
            qual_revision = add_qualification_views(base_url, qual_app_id, auth)
            deploy_app(base_url, qual_app_id, auth, qual_revision)
            check_deploy_status(base_url, qual_app_id, auth)
        else:
            print("KINTONE_QUAL_APP_ID が未設定のため、資格管理一覧はスキップしました。")

    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    print("完了しました。kintoneの一覧ドロップダウンを確認してください。")


if __name__ == "__main__":
    main()
