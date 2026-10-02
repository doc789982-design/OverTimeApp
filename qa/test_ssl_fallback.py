#!/usr/bin/env python3
"""Тест доверия к корпоративным сертификатам при проверке обновлений.

Ситуация с практики: на части рабочих компьютеров сетевой трафик
проверяется подменой сертификатов (в цепочке сайта — самоподписанный
корпоративный центр). Обычный SSL-контекст Python про него не знает,
браузер — знает. Программа теперь делает как браузер: при ошибке
CERTIFICATE_VERIFY_FAILED повторяет запрос с контекстом, которому
доверены все сертификаты из хранилищ Windows.

Проверяется:
  A. _der_to_pem: DER из хранилища превращается в корректный PEM;
  B. windows_ssl_context: создаётся и кэшируется, не падает без Windows;
  C. _urlopen_https: при ошибке сертификата — повтор с контекстом
     (и без повторов на обычных сетевых ошибках);
  D. _is_cert_error отличает ошибку сертификата от обычной сети;
  E. _fetch_version_json: понятное сообщение с именем сайта вместо
     «Не удалось прочитать version.json (SSL…)»;
  F. обычная сетевая ошибка — прежний текст.

Запуск:
    OVERTIMETAB_TEST_SSL_FALLBACK=1 python3 qa/test_ssl_fallback.py
"""
import os
import sys
import urllib.error

os.environ.setdefault("OVERTIMETAB_TEST_SSL_FALLBACK", "1")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import app_update  # noqa: E402


class _Sentinel:
    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def main() -> int:
    import ssl
    import base64

    # ── A. DER → PEM ──
    der = b"\x30\x03\x02\x01\x05" * 8          # произвольные байты
    pem = app_update._der_to_pem(der)
    assert pem.startswith("-----BEGIN CERTIFICATE-----\n")
    assert pem.endswith("-----END CERTIFICATE-----\n")
    body = pem.split("-----")[2].strip()
    assert base64.b64decode(body) == der
    print("A: DER из хранилища превращается в PEM ✓")

    # ── B. контекст ──
    app_update._WINDOWS_SSL_CONTEXT = None      # сброс кэша
    ctx1 = app_update.windows_ssl_context()
    ctx2 = app_update.windows_ssl_context()
    assert ctx1 is ctx2, "контекст должен кэшироваться"
    assert ctx1.check_hostname
    app_update._WINDOWS_SSL_CONTEXT = None
    print("B: контекст хранилищ Windows создаётся и кэшируется ✓")

    # ── C. повтор при ошибке сертификата ──
    real_urlopen = app_update.urllib.request.urlopen
    calls = []

    def fake_urlopen(url, timeout=None, data=None, context=None):
        calls.append(context is not None)
        if len(calls) == 1:
            raise urllib.error.URLError(
                ssl.SSLCertVerificationError(
                    "SSL: CERTIFICATE_VERIFY_FAILED "
                    "self-signed certificate in certificate chain"))
        return _Sentinel()

    try:
        app_update.urllib.request.urlopen = fake_urlopen
        app_update._WINDOWS_SSL_CONTEXT = None
        got = app_update._urlopen_https("https://post.mvd.ru/version.json", 8)
        assert isinstance(got, _Sentinel)
        assert calls == [False, True], calls  # второй заход — с контекстом
    finally:
        app_update.urllib.request.urlopen = real_urlopen

    # обычная сетевая ошибка — без повторов
    calls2 = []

    def fake_network(url, timeout=None, data=None, context=None):
        calls2.append(1)
        raise urllib.error.URLError(OSError(101, "Network is unreachable"))

    try:
        app_update.urllib.request.urlopen = fake_network
        try:
            app_update._urlopen_https("https://post.mvd.ru/version.json", 8)
            raise AssertionError("должна была быть ошибка сети")
        except urllib.error.URLError:
            pass
        assert len(calls2) == 1, "повтор допустим только для сертификатов"
    finally:
        app_update.urllib.request.urlopen = real_urlopen
    print("C: при ошибке сертификата — повтор с контекстом; сеть — сразу ошибка ✓")

    # ── D. классификация ──
    assert app_update._is_cert_error(
        urllib.error.URLError(ssl.SSLCertVerificationError("verify failed")))
    assert app_update._is_cert_error(ssl.SSLError("bad"))
    assert not app_update._is_cert_error(
        urllib.error.URLError(OSError(101, "unreachable")))
    print("D: ошибка сертификата отличается от обычной сети ✓")

    # ── E. понятное сообщение ──
    def fake_fail(url, timeout, data=None):
        raise urllib.error.URLError(
            ssl.SSLCertVerificationError(
                "self-signed certificate in certificate chain"))

    real = app_update._urlopen_https
    try:
        app_update._urlopen_https = fake_fail
        info, err = app_update._fetch_version_json(
            "https://post.mvd.ru/~u@mvd.ru/")
        assert info is None
        assert "не доверяет сертификату" in err, err
        assert "post.mvd.ru" in err, err
        assert "администратору" in err, err
    finally:
        app_update._urlopen_https = real
    print("E: «Компьютер не доверяет сертификату сайта post.mvd.ru…» ✓")

    # ── F. обычная ошибка — прежний текст ──
    def fake_net_fail(url, timeout, data=None):
        raise urllib.error.URLError(OSError(101, "Network is unreachable"))

    try:
        app_update._urlopen_https = fake_net_fail
        info, err = app_update._fetch_version_json("https://post.mvd.ru/")
        assert info is None
        assert "Не удалось прочитать version.json" in err, err
    finally:
        app_update._urlopen_https = real
    print("F: обычная сетевая ошибка — прежний текст ✓")

    print("═══ КОРПОРАТИВНЫЕ СЕРТИФИКАТЫ: ПОВТОР КАК БРАУЗЕР + ПОНЯТНАЯ ОШИБКА ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
