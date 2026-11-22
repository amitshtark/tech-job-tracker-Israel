import json
import unittest
from unittest.mock import patch

from bs4 import BeautifulSoup

from job_crawler import (
    batch_verify_jobs_with_ai,
    extract_candidates_from_soup,
    extract_clean_text_from_soup,
    has_explicit_target_location,
    is_explicit_non_target_location,
    is_complete_ai_approval,
)


class JobLocationTests(unittest.TestCase):
    def test_reported_palantir_jobs_do_not_pass_israel_location_gate(self):
        observed_jobs = {
            "https://jobs.lever.co/palantir/1345438c-ebfc-4fa5-b545-30c1414f317c": "Washington, D.C.",
            "https://jobs.lever.co/palantir/a8174f9c-6f46-46b4-8e15-d1ff9e37c9eb": "Palo Alto, CA",
            "https://jobs.lever.co/palantir/96a0ce26-cf84-4fa8-934b-acc4363620b2": "Washington, D.C.",
        }

        for url, page_location in observed_jobs.items():
            with self.subTest(url=url):
                self.assertFalse(has_explicit_target_location(page_location))
                self.assertTrue(is_explicit_non_target_location(page_location))
        self.assertFalse(has_explicit_target_location("Unspecified"))
        self.assertFalse(has_explicit_target_location("Remote"))
        self.assertTrue(is_explicit_non_target_location("Remote"))
        self.assertFalse(is_explicit_non_target_location(""))
        self.assertFalse(is_explicit_non_target_location("Unspecified"))
        self.assertFalse(is_explicit_non_target_location(" unspecified "))

    def test_explicit_israel_location_still_passes(self):
        valid_israel_job = {
            "url": (
                "https://www.upwind.io/careers/co/tel-aviv/5C.D67/"
                "research-software-engineer/all"
            ),
            "page_location": "Ramat Gan, Israel",
        }

        self.assertTrue(has_explicit_target_location(valid_israel_job["page_location"]))
        self.assertTrue(is_complete_ai_approval({
            "approved": True,
            "is_single_specific_job": True,
            "is_target_role": True,
            "is_matching_yoe": True,
            "is_matching_location": True,
        }))

    def test_ai_prompt_does_not_infer_location_from_company(self):
        response = {
            "candidates": [
                {"content": {"parts": [{"text": json.dumps({"evaluations": []})}]}}
            ]
        }
        with patch("job_crawler.call_gemini_with_retries", return_value=response) as call:
            batch_verify_jobs_with_ai(
                [{
                    "company": "Palantir Israel",
                    "title": "Backend Software Engineer",
                    "url": "https://jobs.lever.co/palantir/example",
                    "page_location": "Washington, D.C.",
                    "requirements_text": "Role requirements",
                }]
            )

        prompt = call.call_args.args[0]["contents"][0]["parts"][0]["text"]
        self.assertIn("Never infer location from the company name", prompt)
        self.assertIn("If it is Unspecified", prompt)
        self.assertIn(".co.il", prompt)
        self.assertIn("Hebrew text", prompt)
        self.assertIn("title, URL, and description", prompt)
        self.assertNotIn("Israeli company with no foreign office mentioned", prompt)

    def test_ai_approval_requires_all_four_criteria(self):
        approval = {
            "approved": True,
            "is_single_specific_job": True,
            "is_target_role": True,
            "is_matching_yoe": True,
            "is_matching_location": True,
        }

        self.assertTrue(is_complete_ai_approval(approval))
        self.assertFalse(is_complete_ai_approval({**approval, "is_matching_location": False}))
        self.assertFalse(is_complete_ai_approval({**approval, "is_matching_location": None}))

    def test_card_location_is_preserved_on_html_candidates(self):
        soup = BeautifulSoup(
            """
            <article>
              <h2>Backend Software Engineer</h2>
              <p>Tel Aviv, Israel</p>
              <a href="/jobs/backend-engineer">Backend Software Engineer</a>
            </article>
            """,
            "html.parser",
        )

        candidates = extract_candidates_from_soup(
            soup,
            "https://example.com/careers",
            {"name": "Example"},
        )

        self.assertEqual(len(candidates), 1)
        self.assertEqual(candidates[0]["page_location"], "Tel Aviv")

    def test_jsonld_country_object_is_normalized(self):
        description = "A software engineering role. " * 15
        job_posting = {
            "@context": "https://schema.org",
            "@type": "JobPosting",
            "description": description,
            "jobLocation": {
                "address": {
                    "addressLocality": "Ramat Gan",
                    "addressCountry": {"@type": "Country", "name": "IL"},
                }
            },
        }
        soup = BeautifulSoup(
            f'<script type="application/ld+json">{json.dumps(job_posting)}</script>',
            "html.parser",
        )

        _, page_location = extract_clean_text_from_soup(soup)

        self.assertEqual(page_location, "Ramat Gan, Israel")
        self.assertTrue(has_explicit_target_location(page_location))


if __name__ == "__main__":
    unittest.main()
