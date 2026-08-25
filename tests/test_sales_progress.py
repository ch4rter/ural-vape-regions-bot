from datetime import date, datetime
from decimal import Decimal

from sales_progress import (
    ManagerProgress,
    fetch_team_revenue,
    month_timing,
    progress_bar,
    render_progress,
)


class FakeClient:
    def __init__(self, rows):
        self.rows = rows
        self.calls = []

    def _rows(self, endpoint, params):
        self.calls.append((endpoint, params))
        return self.rows


def test_fetch_team_revenue_uses_one_lightweight_request_and_sorts():
    valera = "https://api.moysklad.ru/entity/saleschannel/valera"
    andrey = "https://api.moysklad.ru/entity/saleschannel/andrey"
    client = FakeClient([
        {"sum": 125_000, "salesChannel": {"meta": {"href": valera}}},
        {"sum": 300_000, "salesChannel": {"meta": {"href": andrey}}},
        {"sum": 25_000, "salesChannel": {"meta": {"href": valera + "?expand=x"}}},
    ])

    rows = fetch_team_revenue(
        client, "2026-08", [("Валера", valera), ("Андрей", andrey)]
    )

    assert [(row.name, row.revenue) for row in rows] == [
        ("Андрей", Decimal("3000")),
        ("Валера", Decimal("1500")),
    ]
    assert len(client.calls) == 1
    endpoint, params = client.calls[0]
    assert endpoint == "entity/demand"
    assert params["expand"] == "salesChannel"
    assert "applicable=true" in params["filter"]
    assert "positions" not in params["expand"]


def test_progress_bar_and_rendered_ranking_are_revenue_sorted():
    rows = [
        ManagerProgress("Валера", "v", Decimal("730")),
        ManagerProgress("Андрей", "a", Decimal("200")),
    ]
    text = render_progress(
        "август 2026", "2026-08", Decimal("1000"), rows,
        now=datetime(2026, 8, 25, 14, 35),
    )
    assert progress_bar(Decimal("730"), Decimal("1000")) == "🟩" * 7 + "⬜" * 3
    assert "<b>93,0%</b>" in text
    assert text.startswith("🟩" * 9 + "⬜  <b>93,0%</b>\n\n")
    assert text.index("<b>Валера</b>") < text.index("<b>Андрей</b>")
    assert "До конца месяца: <b>7</b>" in text
    assert "Обновлено: 25.08.2026 в 14:35:00" in text


def test_month_timing_handles_past_and_future_months():
    assert month_timing("2026-07", date(2026, 8, 1))[0] == "📅 Месяц завершён"
    assert month_timing("2026-09", date(2026, 8, 1))[0] == "📅 Месяц ещё не начался"
