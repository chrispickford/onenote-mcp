"""Small helpers shared by the e2e tests."""


def without_ids(value):
    """Return *value* (get_page output) with every "id" key removed, recursively.

    get_page reports each item's OneNote objectID; tests that compare content
    with what they wrote ignore those IDs.
    """
    if isinstance(value, dict):
        return {key: without_ids(item) for key, item in value.items() if key != "id"}
    if isinstance(value, list):
        return [without_ids(item) for item in value]
    return value
