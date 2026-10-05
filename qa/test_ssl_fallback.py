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

    # ── G. ответ в памяти для пути WinHTTP ──
    resp = app_update._BytesResponse(b"0123456789", 200, {"Content-Length": "10"})
    with resp as r:
        assert r.read(4) == b"0123"
        assert r.read() == b"456789"
        assert r.headers.get("Content-Length") == "10"
        assert r.status == 200
    print("G: ответ WinHTTP читается как ответ urlopen ✓")

    # ── H. третья попытка — проверка силами Windows ──
    import tempfile
    from pathlib import Path

    def cert_err(msg="self-signed certificate in certificate chain"):
        return urllib.error.URLError(
            ssl.SSLCertVerificationError("SSL: CERTIFICATE_VERIFY_FAILED " + msg))

    real_urlopen2 = app_update.urllib.request.urlopen
    real_winhttp = app_update._winhttp_get
    report_dir = Path(tempfile.mkdtemp(prefix="ssl_report_"))
    app_update.set_ssl_report_dir(report_dir)
    app_update._LAST_SSL_REPORT = ""
    order = []

    def always_cert(url, timeout=None, data=None, context=None):
        order.append("python:" + ("store" if context is not None else "default"))
        raise cert_err()

    def fake_winhttp_ok(url, timeout):
        order.append("winhttp")
        return app_update._BytesResponse(b'{"version": "2.0.0"}', 200)

    try:
        app_update.urllib.request.urlopen = always_cert
        app_update._winhttp_get = fake_winhttp_ok
        app_update._WINDOWS_SSL_CONTEXT = None
        with app_update._urlopen_https("https://post.mvd.ru/version.json", 8) as r:
            assert r.read() == b'{"version": "2.0.0"}'
        assert order == ["python:default", "python:store", "winhttp"], order
        assert app_update._LAST_SSL_REPORT == "", "успех — отчёт не нужен"
    finally:
        app_update.urllib.request.urlopen = real_urlopen2
        app_update._winhttp_get = real_winhttp
    print("H: после двух отказов Python соединяет сам Windows ✓")

    # ── I. все пути отказали — файл-отчёт и подсказка в сообщении ──
    order2 = []

    def fake_winhttp_cert(url, timeout):
        order2.append("winhttp")
        raise ssl.SSLError(
            "Windows сам проверил сертификат сайта и не принял его: "
            "корневой центр сертификации не входит в доверенные Windows")

    try:
        app_update.urllib.request.urlopen = always_cert
        app_update._winhttp_get = fake_winhttp_cert
        app_update._WINDOWS_SSL_CONTEXT = None
        try:
            app_update._urlopen_https("https://post.mvd.ru/version.json", 8)
            raise AssertionError("должна была быть ошибка сертификата")
        except ssl.SSLError:
            pass
        assert order2 == ["winhttp"], order2
        assert app_update._LAST_SSL_REPORT.endswith("update_ssl_report.txt"), \
            app_update._LAST_SSL_REPORT
        report = Path(app_update._LAST_SSL_REPORT).read_text(encoding="utf-8")
        assert "post.mvd.ru" in report, report
        assert "обычный SSL Python" in report, report
        assert "корни из хранилищ Windows" in report, report
        assert "проверка Windows (WinHTTP)" in report, report
        assert "не входит в доверенные" in report, report
        # сообщение об ошибке ведёт к файлу отчёта
        msg = app_update._cert_error_message(
            "https://post.mvd.ru/version.json", cert_err())
        assert "Подробности — в файле" in msg, msg
        assert "update_ssl_report.txt" in msg, msg
    finally:
        app_update.urllib.request.urlopen = real_urlopen2
        app_update._winhttp_get = real_winhttp
    print("I: полный отказ — отчёт записан, сообщение ведёт к файлу ✓")

    # ── J. WinHTTP не запустился — исходная ошибка сертификата ──
    def fake_winhttp_dead(url, timeout):
        raise RuntimeError("WinHTTP доступен только на Windows")

    try:
        app_update.urllib.request.urlopen = always_cert
        app_update._winhttp_get = fake_winhttp_dead
        app_update._WINDOWS_SSL_CONTEXT = None
        try:
            app_update._urlopen_https("https://post.mvd.ru/version.json", 8)
            raise AssertionError("должна была быть ошибка сертификата")
        except urllib.error.URLError as e:
            assert app_update._is_cert_error(e), "исходная ошибка сертификата"
        report = Path(app_update._LAST_SSL_REPORT).read_text(encoding="utf-8")
        assert "не запустилась" in report, report
    finally:
        app_update.urllib.request.urlopen = real_urlopen2
        app_update._winhttp_get = real_winhttp
    print("J: без WinHTTP — исходная ошибка, отчёт с причиной ✓")

    # ── K. расшифровка кодов отказа Windows ──
    t = app_update._winhttp_flags_text(0x00004000)
    assert "корневой центр" in t, t
    t2 = app_update._winhttp_flags_text(0x00002000 | 0x00010000)
    assert "срок действия" in t2 and "отозван" in t2, t2
    assert "0x" in app_update._winhttp_flags_text(0x80000000)
    print("K: коды отказа Windows читаются по-русски ✓")

    # ── L. ограничения пути WinHTTP ──
    try:
        app_update._winhttp_get("http://post.mvd.ru/", 8)
        raise AssertionError("только https")
    except ValueError:
        pass
    if sys.platform != "win32":
        try:
            app_update._winhttp_get("https://post.mvd.ru/", 8)
            raise AssertionError("не Windows — пути нет")
        except RuntimeError:
            pass
    print("L: WinHTTP — только https и только Windows ✓")

    # ── M. сигнатуры WinHTTP — как в настоящем API Windows ──
    # в 275 WinHttpSendRequest был описан с лишним аргументом — путь падал
    # «this function takes 8 arguments (7 given)»; теперь сверяем всё
    protos = app_update._winhttp_prototypes()
    real_params = {
        "WinHttpOpen": 5, "WinHttpConnect": 4, "WinHttpOpenRequest": 7,
        "WinHttpSetTimeouts": 5, "WinHttpSendRequest": 7,
        "WinHttpReceiveResponse": 2, "WinHttpQueryHeaders": 6,
        "WinHttpQueryDataAvailable": 2, "WinHttpReadData": 4,
        "WinHttpQueryOption": 4, "WinHttpCloseHandle": 1,
    }
    assert set(protos) == set(real_params), set(protos) ^ set(real_params)
    for name, n in real_params.items():
        restype, argtypes = protos[name]
        assert len(argtypes) == n, (name, len(argtypes), n)
        assert restype is not None, name
    print("M: все " + str(len(real_params))
          + " сигнатур WinHTTP совпадают с настоящим API ✓")

    # ── N. фиктивный winhttp.dll: каждый вызов — с тем числом аргументов ──
    import ctypes

    class _FakeWinHttp:
        def __init__(self, script):
            self.calls = []          # (имя функции, число аргументов)
            self.script = script

        def __getattr__(self, name):
            def call(*args):
                self.calls.append((name, len(args)))
                return self.script[name](args)
            return call

    def _set(arg, value):
        arg._obj.value = value       # запись в byref-переменную

    def _run(script):
        fake = _FakeWinHttp(script)
        real_windll = getattr(ctypes, "WinDLL", None)   # на Linux его нет
        ctypes.WinDLL = lambda name, use_last_error=False: fake
        os.environ["OVERTIMETAB_TEST_WINHTTP"] = "1"
        try:
            try:
                return fake, app_update._winhttp_get(
                    "https://post.mvd.ru/version.json", 8), None
            except Exception as e:
                return fake, None, e
        finally:
            if real_windll is None:
                del ctypes.WinDLL
            else:
                ctypes.WinDLL = real_windll
            del os.environ["OVERTIMETAB_TEST_WINHTTP"]

    def _check_args(fake):
        for name, argc in fake.calls:
            assert argc == len(protos[name][1]), (name, argc)

    ok_script = {
        "WinHttpOpen": lambda a: 0x1111,
        "WinHttpSetTimeouts": lambda a: 1,
        "WinHttpConnect": lambda a: 0x2222,
        "WinHttpOpenRequest": lambda a: 0x3333,
        "WinHttpSendRequest": lambda a: 1,
        "WinHttpReceiveResponse": lambda a: 1,
        "WinHttpQueryHeaders": lambda a: (_set(a[3], 200), _set(a[4], 4), 1)[2],
        "WinHttpQueryDataAvailable": lambda a: (_set(a[1], 0), 1)[1],
        "WinHttpReadData": lambda a: 1,
        "WinHttpCloseHandle": lambda a: 1,
        "WinHttpQueryOption": lambda a: 0,
    }

    # успех: 200, пустое тело
    fake, result, exc = _run(ok_script)
    assert exc is None, exc
    with result as r:
        assert r.status == 200 and r.read() == b""
    _check_args(fake)
    assert any(n == "WinHttpSendRequest" for n, _ in fake.calls)

    # 404 → HTTPError с телом
    state = {"avail": 0}

    def _avail(a):
        state["avail"] += 1
        _set(a[1], 5 if state["avail"] == 1 else 0)
        return 1

    def _read(a):
        a[1][0:5] = b"error"
        _set(a[3], 5)
        return 1

    fake, result, exc = _run(dict(
        ok_script,
        WinHttpQueryHeaders=lambda a: (_set(a[3], 404), _set(a[4], 4), 1)[2],
        WinHttpQueryDataAvailable=_avail, WinHttpReadData=_read))
    assert isinstance(exc, urllib.error.HTTPError) and exc.code == 404, exc
    assert exc.read() == b"error"
    _check_args(fake)

    # отказ сертификата: код 12175 + флаг «корень не в доверенных»
    def _send_fail(a):
        app_update._WINHTTP_TEST_LAST_ERROR = 12175
        return 0

    fake, result, exc = _run(dict(
        ok_script, WinHttpSendRequest=_send_fail,
        WinHttpQueryOption=lambda a: (_set(a[2], 0x00004000), _set(a[3], 4), 1)[2]))
    app_update._WINHTTP_TEST_LAST_ERROR = None
    assert isinstance(exc, ssl.SSLError), exc
    assert "корневой центр" in str(exc), exc
    _check_args(fake)

    # переадресация: 302 на другой адрес → там 200
    hops = {"n": 0}

    def _headers(a):
        if a[1] == 33:                # Location
            if a[3] is not None:
                a[3].value = "https://post.mvd.ru/next/version.json"
            _set(a[4], 80)
            return 1
        hops["n"] += 1
        _set(a[3], 302 if hops["n"] == 1 else 200)
        _set(a[4], 4)
        return 1

    fake, result, exc = _run(dict(ok_script, WinHttpQueryHeaders=_headers))
    assert exc is None, exc
    assert result.status == 200 and hops["n"] == 2
    _check_args(fake)
    print("N: фиктивный WinHTTP — вызовы сходятся с сигнатурами; 200, 404, сертификат, переадресация ✓")

    # ── P. подсчёт хранилищ сертификатов безопасен без Windows ──
    counts = app_update._windows_store_counts()
    if sys.platform == "win32":
        assert set(counts) == {"ROOT", "CA"}
    else:
        assert counts == {}, counts
    assert app_update._machine_store_count("ROOT") is None \
        or sys.platform == "win32"
    print("P: подсчёт хранилищ (пользователь + компьютер) безопасен ✓")

    app_update.set_ssl_report_dir(None)
    app_update._LAST_SSL_REPORT = ""
    print("═══ КОРПОРАТИВНЫЕ СЕРТИФИКАТЫ: ТРИ ПУТИ + ОТЧЁТ ДЛЯ ДИАГНОСТИКИ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
