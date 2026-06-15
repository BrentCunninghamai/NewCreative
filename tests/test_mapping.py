from m365_migrate.mapping import read_mapping, rewrite_upn, write_mapping
from m365_migrate.models import PlannedUser


def test_rewrite_upn_matches_source_domain():
    assert (
        rewrite_upn("jane@contoso.onmicrosoft.com", "contoso.onmicrosoft.com", "fabrikam.onmicrosoft.com")
        == "jane@fabrikam.onmicrosoft.com"
    )


def test_rewrite_upn_is_case_insensitive_on_domain():
    assert (
        rewrite_upn("jane@CONTOSO.onmicrosoft.com", "contoso.onmicrosoft.com", "fabrikam.onmicrosoft.com")
        == "jane@fabrikam.onmicrosoft.com"
    )


def test_rewrite_upn_leaves_other_domains_intact():
    assert (
        rewrite_upn("jane@vanity.com", "contoso.onmicrosoft.com", "fabrikam.onmicrosoft.com")
        == "jane@vanity.com"
    )


def test_mapping_round_trip(tmp_path):
    planned = [
        PlannedUser(
            source_id="1",
            source_upn="jane@contoso.onmicrosoft.com",
            target_upn="jane@fabrikam.onmicrosoft.com",
            display_name="Jane Doe",
            action="create",
        ),
        PlannedUser(
            source_id="2",
            source_upn="guest@contoso.onmicrosoft.com",
            target_upn="guest@fabrikam.onmicrosoft.com",
            display_name=None,
            action="skip",
            reason="guest user",
        ),
    ]
    path = write_mapping(tmp_path / "out" / "user_mapping.csv", planned)
    assert path.exists()

    loaded = read_mapping(path)
    assert len(loaded) == 2
    assert loaded[0].target_upn == "jane@fabrikam.onmicrosoft.com"
    assert loaded[1].action == "skip"
    assert loaded[1].reason == "guest user"
    assert loaded[1].display_name is None
