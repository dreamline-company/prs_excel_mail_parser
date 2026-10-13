"""Почтовая часть бота на поддельном IMAP-сервере.

Сервер ведёт себя как настоящий: у писем постоянные UID, порядковые номера
сдвигаются после EXPUNGE, FETCH RFC822 ставит \\Seen, BODY.PEEK — нет.
"""

import importlib
import sys
from email.message import EmailMessage
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def _message(subject: str, attachment: str | None) -> bytes:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg.set_content("сводка")
    if attachment:
        msg.add_attachment(
            b"PK fake xlsx",
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=attachment,
        )
    return msg.as_bytes()


class FakeMailbox:
    def __init__(self, messages: list[tuple[int, bytes]]) -> None:
        # [uid, raw, flags]
        self.items = [[uid, raw, set()] for uid, raw in messages]
        self.logins: list[tuple[str, str]] = []


class FakeIMAP:
    mailbox: FakeMailbox

    def __init__(self, host, port, ssl_context=None, timeout=None) -> None:
        self.host, self.port = host, port

    def login(self, user, password):
        self.mailbox.logins.append((user, password))
        return "OK", [b"logged in"]

    def select(self, folder):
        return "OK", [str(len(self.mailbox.items)).encode()]

    def _by_uid(self, uid: bytes):
        return next(item for item in self.mailbox.items if item[0] == int(uid))

    def uid(self, command, *args):
        if command == "SEARCH":
            uids = [str(i[0]) for i in self.mailbox.items if "\\Seen" not in i[2]]
            return "OK", [" ".join(uids).encode()]
        if command == "FETCH":
            uid, spec = args
            item = self._by_uid(uid)
            if "PEEK" not in spec:
                item[2].add("\\Seen")
            return "OK", [(b"1 (UID %s BODY[] {n}" % uid, item[1]), b")"]
        if command == "STORE":
            uid, _op, flags = args
            self._by_uid(uid)[2].add(flags.strip("()"))
            return "OK", [b""]
        raise AssertionError(command)

    def expunge(self):
        self.mailbox.items = [i for i in self.mailbox.items if "\\Deleted" not in i[2]]
        return "OK", [b""]

    def close(self):
        return "OK", [b""]

    def logout(self):
        return "BYE", [b""]


@pytest.fixture
def bot(monkeypatch, tmp_path):
    monkeypatch.setenv("PARSER_EMAIL", "emg_prs-svodki@emg.kmg.kz")
    monkeypatch.setenv("PARSER_PASSWORD", "secret")
    monkeypatch.setenv("IMAP_HOST", "mail.emg.kmg.kz")
    monkeypatch.delenv("PARSER_LOGIN", raising=False)
    sys.modules.pop("main", None)
    module = importlib.import_module("main")
    monkeypatch.setattr(module, "inbox_dir", tmp_path)
    monkeypatch.setattr(module.imaplib, "IMAP4_SSL", FakeIMAP)
    return module


def test_deletes_exactly_processed_messages_by_uid(bot, tmp_path) -> None:
    # 10 — без вложения, 11 и 13 — сводки, 12 — уже прочитанное письмо.
    FakeIMAP.mailbox = FakeMailbox(
        [
            (10, _message("привет", None)),
            (11, _message("сводка 1", "svodka_1.xlsx")),
            (12, _message("старое", None)),
            (13, _message("сводка 2", "svodka_2.xlsx")),
        ],
    )
    FakeIMAP.mailbox.items[2][2].add("\\Seen")

    bot.fetch_mail_to_inbox()

    left = {item[0]: item[2] for item in FakeIMAP.mailbox.items}
    assert set(left) == {10, 12}  # удалены ровно сводки, остальные на месте
    assert "\\Seen" in left[10]  # письмо без xlsx помечено прочитанным
    assert sorted(p.name for p in tmp_path.iterdir()) == [
        "svodka_1_11.xlsx",
        "svodka_2_13.xlsx",
    ]


def test_failed_save_keeps_message_unread(bot, monkeypatch, tmp_path) -> None:
    FakeIMAP.mailbox = FakeMailbox([(21, _message("сводка", "svodka.xlsx"))])
    monkeypatch.setattr(bot, "inbox_dir", tmp_path / "нет такой папки")

    bot.fetch_mail_to_inbox()

    [item] = FakeIMAP.mailbox.items
    assert item[0] == 21
    assert "\\Seen" not in item[2]  # следующая проверка найдёт письмо снова


def test_login_override_for_exchange(bot, monkeypatch) -> None:
    FakeIMAP.mailbox = FakeMailbox([])
    monkeypatch.setattr(bot, "LOGIN", "EMG\\emg_prs-svodki")

    assert bot.connect_to_mail() is not None
    assert FakeIMAP.mailbox.logins == [("EMG\\emg_prs-svodki", "secret")]
