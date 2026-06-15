import httpx
import respx

from m365_migrate.graph_client import GraphClient, GraphError


@respx.mock
def test_get_all_follows_pagination(static_token):
    base = "https://graph.microsoft.com/v1.0"
    # Both the initial request and the nextLink request hit the /users path, so
    # drive them with a single route returning the two pages in sequence.
    route = respx.get(f"{base}/users")
    route.side_effect = [
        httpx.Response(
            200,
            json={
                "value": [{"id": "1"}, {"id": "2"}],
                "@odata.nextLink": f"{base}/users?$skiptoken=abc",
            },
        ),
        httpx.Response(200, json={"value": [{"id": "3"}]}),
    ]

    client = GraphClient(static_token)
    ids = [u["id"] for u in client.get_all("/users")]
    assert ids == ["1", "2", "3"]


@respx.mock
def test_retries_on_429_then_succeeds(static_token):
    base = "https://graph.microsoft.com/v1.0"
    route = respx.get(f"{base}/organization")
    route.side_effect = [
        httpx.Response(429, headers={"Retry-After": "0"}, text="throttled"),
        httpx.Response(200, json={"value": [{"displayName": "Fabrikam"}]}),
    ]

    slept = []
    client = GraphClient(static_token, sleep=slept.append)
    data = client.get("/organization")

    assert data["value"][0]["displayName"] == "Fabrikam"
    assert slept == [0.0]  # honored Retry-After: 0


@respx.mock
def test_raises_graph_error_on_4xx(static_token):
    base = "https://graph.microsoft.com/v1.0"
    respx.get(f"{base}/users").mock(return_value=httpx.Response(403, text="forbidden"))

    client = GraphClient(static_token)
    try:
        client.get("/users")
    except GraphError as exc:
        assert exc.status_code == 403
    else:  # pragma: no cover
        raise AssertionError("expected GraphError")
