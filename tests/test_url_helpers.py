import unittest

from listings_project import Config, ListingsProjectAnalyzer, slugify_title, split_uuid_suffix


class UrlHelperTests(unittest.TestCase):
    def setUp(self):
        self.analyzer = ListingsProjectAnalyzer(Config())

    def test_clean_url_normalizes_host_and_removes_query(self):
        value = self.analyzer.clean_url(
            "https://listingsproject.com/listings/example-title?source=email#details"
        )
        self.assertEqual(
            value,
            "https://www.listingsproject.com/listings/example-title",
        )

    def test_clean_url_rejects_external_host(self):
        self.assertIsNone(
            self.analyzer.clean_url("https://example.com/listings/example-title")
        )

    def test_clean_url_rejects_non_listing_path(self):
        self.assertIsNone(
            self.analyzer.clean_url(
                "https://www.listingsproject.com/real-estate/new-york-city/sublets"
            )
        )

    def test_uuid_suffix_is_split(self):
        slug = "sunny-room-123e4567-e89b-12d3-a456-426614174000"
        base, suffix = split_uuid_suffix(slug)
        self.assertEqual(base, "sunny-room")
        self.assertEqual(suffix, "123e4567-e89b-12d3-a456-426614174000")

    def test_non_uuid_slug_is_unchanged(self):
        base, suffix = split_uuid_suffix("sunny-room")
        self.assertEqual(base, "sunny-room")
        self.assertIsNone(suffix)

    def test_slugify_title_is_explicitly_approximate(self):
        self.assertEqual(slugify_title("Bright Room — SoHo!"), "bright-room-soho")


if __name__ == "__main__":
    unittest.main()
