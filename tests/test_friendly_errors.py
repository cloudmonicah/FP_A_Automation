from shared.friendly_errors import friendly, GENERIC


def test_technical_error_names_never_reach_the_user():
    for error in (
        AttributeError("'float' object has no attribute 'quantize'"),
        KeyError("GL Account Mapping"),
        ValueError("Expected one Revenue subtotal for {'Company': 'ALIV'}; found 0."),
        RuntimeError('Tableau "Copy" menu could not be found.'),
    ):
        text = friendly(error)
        assert "Error" not in text and "quantize" not in text and "{" not in text
    assert friendly(AttributeError("x")) == GENERIC


def test_common_problems_explain_what_to_do():
    assert "Close it" in friendly(PermissionError("denied"))
    assert "different month" in friendly(
        ValueError("CBL statement period does not match the selected closed month.")
    )
    assert "227,563" in friendly(
        ValueError("Media rows add to 227,562.00 but Grand Total is 227,563.00; check")
    )
