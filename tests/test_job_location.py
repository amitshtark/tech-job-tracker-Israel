import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from bs4 import BeautifulSoup

import job_crawler
from job_crawler import (
    apply_second_judge_results,
    add_confirmed_jobs_to_seen,
    audit_run_with_ai,
    batch_verify_jobs_with_ai,
    extract_candidates_from_soup,
    extract_clean_text_from_soup,
    has_excessive_min_yoe,
    has_explicit_target_location,
    is_explicit_non_target_location,
    is_title_candidate,
    is_complete_ai_approval,
    should_reject_pre_ai_location,
    new_run_diagnostics,
    record_company_scrape,
    record_filter_rejection_sample,
    record_playwright_detail_cap,
    record_rejected_candidate,
    call_gemini_with_retries,
    is_excluded_role_title,
    is_seniority_title,
)


class JobLocationTests(unittest.TestCase):
    def test_run_diagnostics_track_scrape_methods_failures_and_rejections(self):
        diagnostics = new_run_diagnostics(2)
        record_company_scrape(
            diagnostics,
            "Example",
            "https://example.com/careers",
            True,
            method="ats_api",
        )
        record_company_scrape(
            diagnostics,
            "Broken",
            "https://broken.example/jobs",
            False,
            reason="HTTP error 404",
        )
        record_rejected_candidate(
            diagnostics,
            {"company": "Example", "title": "Engineer", "url": "https://example.com/job"},
            "description_gate",
            "Description shorter than 150 characters",
        )
        self.assertEqual(record_playwright_detail_cap(diagnostics, 42, 40), 2)

        self.assertEqual(diagnostics["companies"]["successful_scrapes"], 1)
        self.assertEqual(diagnostics["companies"]["failed_scrapes"], 1)
        self.assertEqual(diagnostics["companies"]["scrape_methods"]["ats_api"], 1)
        self.assertEqual(diagnostics["companies"]["failure_breakdown"], {"404": 1})
        self.assertEqual(
            diagnostics["companies"]["broken_404"],
            [{"name": "Broken", "url": "https://broken.example/jobs"}],
        )
        self.assertEqual(
            diagnostics["ai"]["rejected_sample"][0]["rejection_stage"],
            "description_gate",
        )
        self.assertEqual(diagnostics["filtering"]["playwright_detail_fetch_cap_hits"], 2)

    def test_diagnostics_record_location_yoe_and_spa_samples_up_to_five(self):
        diagnostics = new_run_diagnostics(1)
        for index in range(7):
            record_filter_rejection_sample(
                diagnostics,
                "location_rejected_sample",
                {"company": f"Company {index}", "title": "Engineer", "page_location": "US", "url": f"https://x/{index}"},
            )
            record_filter_rejection_sample(
                diagnostics,
                "yoe_rejected_sample",
                {"company": f"Company {index}", "title": "Engineer", "matched_phrase": "3+ years", "url": f"https://x/{index}"},
            )
            record_filter_rejection_sample(
                diagnostics,
                "spa_failures_sample",
                {"company": f"Company {index}", "url": f"https://x/{index}", "reason": "Timeout"},
            )

        self.assertEqual(len(diagnostics["location_rejected_sample"]), 5)
        self.assertEqual(len(diagnostics["yoe_rejected_sample"]), 5)
        self.assertEqual(len(diagnostics["spa_failures_sample"]), 5)
        self.assertEqual(diagnostics["location_rejected_sample"][0]["page_location"], "US")
        self.assertEqual(diagnostics["yoe_rejected_sample"][0]["matched_phrase"], "3+ years")

    def test_stage_one_retries_http_429_once_with_rotated_user_agent(self):
        rate_limited = type("Response", (), {"status_code": 429})()
        ok_response = type(
            "Response",
            (),
            {"status_code": 200, "text": f"<html><body>{'Career information. ' * 40}</body></html>"},
        )()
        with (
            patch("job_crawler.try_ats_api_scrape", return_value=(False, [])),
            patch("job_crawler.requests.get", side_effect=[rate_limited, ok_response]) as get,
            patch("job_crawler.time.sleep") as sleep,
        ):
            result = job_crawler.scrape_single_company({
                "name": "Example",
                "career_url": "https://example.com/careers",
            })

        self.assertTrue(result[1])
        self.assertEqual(get.call_count, 2)
        sleep.assert_called_once_with(3)
        self.assertNotEqual(
            get.call_args_list[0].kwargs["headers"]["User-Agent"],
            get.call_args_list[1].kwargs["headers"]["User-Agent"],
        )

    def test_second_judge_removes_false_positive_before_seen_jobs_save(self):
        false_positive = {
            "company": "Example",
            "title": "Senior Backend Engineer",
            "url": "https://example.com/jobs/senior-backend",
            "location": "Tel Aviv",
            "experience": "5+ years",
            "text_snippet": "Requires at least five years.",
        }
        valid_job = {
            "company": "Example",
            "title": "Backend Engineer",
            "url": "https://example.com/jobs/backend",
            "location": "Tel Aviv",
            "experience": "0-2 years",
        }
        false_positive_hash = job_crawler.compute_job_hash(
            false_positive["company"],
            false_positive["title"],
            false_positive["url"],
        )
        confirmed_new_jobs = [false_positive.copy(), valid_job.copy()]
        seen_jobs = [{"company": "Existing", "title": "Engineer", "url": "https://example.com/old"}]
        audit_result = {
            "second_judge": [{
                "job_index": 1,
                "is_single_specific_job": True,
                "is_target_role": True,
                "is_matching_yoe": False,
                "is_matching_location": True,
                "approved": False,
                "reason": "Requires more experience than configured",
            }, {
                "job_index": 2,
                "is_single_specific_job": True,
                "is_target_role": True,
                "is_matching_yoe": True,
                "is_matching_location": True,
                "approved": True,
                "reason": "",
            }],
        }

        demoted = apply_second_judge_results(
            confirmed_new_jobs,
            seen_jobs,
            [false_positive, valid_job],
            audit_result,
        )
        add_confirmed_jobs_to_seen(confirmed_new_jobs, seen_jobs)
        with tempfile.TemporaryDirectory() as temp_dir:
            seen_jobs_path = Path(temp_dir, "seen_jobs.json")
            job_crawler.save_json(str(seen_jobs_path), seen_jobs)
            saved_jobs = json.loads(seen_jobs_path.read_text(encoding="utf-8"))

        self.assertEqual(confirmed_new_jobs, [valid_job])
        self.assertEqual(len(demoted), 1)
        self.assertEqual(demoted[0]["reason"], "Requires more experience than configured")
        saved_hashes = {
            job_crawler.compute_job_hash(item["company"], item["title"], item["url"])
            for item in saved_jobs
        }
        self.assertNotIn(false_positive_hash, saved_hashes)
        self.assertIn(
            job_crawler.compute_job_hash(valid_job["company"], valid_job["title"], valid_job["url"]),
            saved_hashes,
        )

    def test_gemini_rate_limit_error_tracks_exact_429_responses(self):
        responses = [
            type("Response", (), {"status_code": 429, "text": "rate limited"})(),
            type("Response", (), {"status_code": 503, "text": "unavailable"})(),
            type("Response", (), {"status_code": 429, "text": "rate limited"})(),
        ]
        with (
            patch.object(job_crawler, "GEMINI_API_KEY", "test-key"),
            patch.object(job_crawler.requests, "post", side_effect=responses),
            patch.object(job_crawler.time, "sleep"),
            self.assertRaises(job_crawler.GeminiRateLimitError) as raised,
        ):
            call_gemini_with_retries({"contents": []})

        self.assertEqual(raised.exception.count, 2)

    def test_ai_audit_makes_one_call_with_config_and_funnel_evidence(self):
        response_body = {
            "second_judge": [],
            "conclusions": "נמצאו צווארי בקבוק.",
            "recommendations": ["בדקו את מגבלת Playwright.", "תקנו קישורי 404."],
        }
        response = {
            "candidates": [{
                "content": {"parts": [{"text": json.dumps(response_body, ensure_ascii=False)}]}
            }]
        }
        diagnostics = new_run_diagnostics(1)
        diagnostics["ai"]["approved_jobs"] = [{
            "company": "Example",
            "title": "Backend Engineer",
            "url": "https://example.com/job",
            "location": "Tel Aviv",
            "experience": "0-2 years",
            "text_snippet": "Role text for second-judge verification",
        }]
        diagnostics["ai"]["rejected_sample"] = [{
            "company": "Example",
            "title": "Engineer",
            "url": "https://example.com/other",
            "rejection_stage": "description_gate",
            "reason": "Description shorter than 150 characters",
        }]

        with patch("job_crawler.call_gemini_with_retries", return_value=response) as gemini_call:
            result = audit_run_with_ai(
                diagnostics,
                {"min_yoe": 0, "max_yoe": 2, "location": "Israel", "target_roles": ["Backend"]},
            )

        self.assertEqual(result, response_body)
        gemini_call.assert_called_once()
        prompt = gemini_call.call_args.args[0]["contents"][0]["parts"][0]["text"]
        self.assertIn("Role text for second-judge verification", prompt)
        self.assertIn("Description shorter than 150 characters", prompt)
        self.assertIn('"target_roles": ["Backend"]', prompt)

    def test_config_local_json_overrides_config_json(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir, "config.json")
            local_config_path = Path(temp_dir, "config.local.json")
            config_path.write_text(json.dumps({"location": "Germany", "max_yoe": 5}), encoding="utf-8")
            local_config_path.write_text(json.dumps({"location": "Israel"}), encoding="utf-8")

            with (
                patch.object(job_crawler, "BASE_DIR", temp_dir),
                patch.object(job_crawler, "CONFIG_FILE", str(config_path)),
            ):
                config = job_crawler.load_user_config()

        self.assertEqual(config["location"], "Israel")
        self.assertEqual(config["max_yoe"], job_crawler.DEFAULT_CONFIG["max_yoe"])

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

    def test_unspecified_page_location_passes_pre_ai_gate_but_foreign_location_does_not(self):
        for location in ("", "Unspecified", " unspecified "):
            with self.subTest(location=location):
                self.assertFalse(is_explicit_non_target_location(location))

        self.assertTrue(is_explicit_non_target_location("Washington, D.C."))
        self.assertFalse(is_explicit_non_target_location("Tel Aviv, Israel"))

    def test_dialog_hebrew_job_on_coil_with_noisy_location_passes_pre_ai_gate(self):
        job = {
            "company": "Dialog",
            "title": "Dialog — מפתח Full Stack",
            "url": "https://www.dialog.co.il/high-tech/jobs/software/full-stack-developer",
            "page_location": "מספר משרה: 95400 / לינק ללכאן",
        }

        self.assertFalse(should_reject_pre_ai_location(job))

    def test_pre_ai_location_gate_still_rejects_explicit_foreign_location_on_coil(self):
        job = {
            "title": "מפתח Full Stack",
            "url": "https://www.dialog.co.il/jobs/full-stack",
            "page_location": "Washington, D.C.",
        }

        self.assertTrue(should_reject_pre_ai_location(job))

    def test_yoe_filter_uses_range_minimum_and_keeps_uncertain_ranges_for_ai(self):
        self.assertFalse(has_excessive_min_yoe("Requires 1-3 years of experience", 2))
        self.assertFalse(has_excessive_min_yoe("Requires 2-3 years of experience", 2))
        self.assertTrue(has_excessive_min_yoe("Requires 3-5 years of experience", 2))
        self.assertTrue(has_excessive_min_yoe("Requires 3+ years of experience", 2))

    def test_title_filter_rejects_seniority_and_configured_excluded_roles(self):
        for title in ("Senior Backend Engineer", "Lead Software Engineer"):
            with self.subTest(title=title):
                self.assertTrue(is_seniority_title(title))
                self.assertFalse(is_title_candidate(
                    title,
                    "https://jobs.example.com/careers/backend-engineer",
                ))

        self.assertFalse(is_title_candidate(
            "Junior QA Tester",
            "https://jobs.example.com/careers/qa-tester",
        ))
        self.assertTrue(is_excluded_role_title("Junior QA Tester"))
        self.assertFalse(is_title_candidate(
            "Technical Support Engineer",
            "https://jobs.example.com/jobs/support-engineer",
        ))

    def test_title_filter_keeps_non_senior_tech_roles_for_ai_verification(self):
        self.assertTrue(is_title_candidate(
            "Backend Engineer",
            "https://jobs.example.com/careers/backend-engineer",
        ))
        self.assertFalse(is_seniority_title("Backend Engineer"))
        self.assertTrue(is_title_candidate(
            "Junior Backend Engineer",
            "https://jobs.example.com/careers/backend-engineer",
        ))
        self.assertTrue(is_title_candidate(
            "Software Engineer",
            "https://jobs.example.com/jobs/software-engineer",
        ))
        self.assertTrue(is_title_candidate(
            "Backend Engineer - Tel Aviv",
            "https://jobs.example.com/jobs/backend-engineer",
        ))
        self.assertFalse(is_title_candidate(
            "Senior QA Engineer",
            "https://jobs.example.com/jobs/qa-engineer",
        ))
        self.assertFalse(is_title_candidate(
            "DevOps Engineer",
            "https://jobs.example.com/jobs/devops-engineer",
        ))
        self.assertFalse(is_title_candidate(
            "Customer Success Engineer",
            "https://jobs.example.com/jobs/customer-success-engineer",
        ))
        self.assertFalse(is_title_candidate(
            "Sales Engineer",
            "https://jobs.example.com/jobs/sales-engineer",
        ))

    def test_title_filter_rejects_non_technical_analyst_roles(self):
        for title in ("Performance Analyst", "Decision Analyst"):
            with self.subTest(title=title):
                self.assertFalse(is_title_candidate(
                    title,
                    f"https://jobs.example.com/jobs/{title.lower().replace(' ', '-')}",
                ))

        self.assertTrue(is_title_candidate(
            "Data Analyst",
            "https://jobs.example.com/jobs/data-analyst",
        ))

    def test_seniority_yoe_checks_are_enforced_after_batch_ai_response(self):
        response = {
            "candidates": [{
                "content": {
                    "parts": [{
                        "text": json.dumps({
                            "evaluations": [{
                                "job_index": 1,
                                "approved": True,
                                "is_single_specific_job": True,
                                "is_target_role": True,
                                "is_matching_yoe": True,
                                "is_matching_location": True,
                            }]
                        })
                    }]
                }
            }]
        }
        with (
            patch.object(job_crawler, "USER_CONFIG", {
                **job_crawler.USER_CONFIG,
                "max_yoe": 2,
            }),
            patch("job_crawler.call_gemini_with_retries", return_value=response) as gemini_call,
        ):
            evaluations = batch_verify_jobs_with_ai([{
                "company": "Example",
                "title": "Senior Backend Engineer",
                "url": "https://example.com/jobs/senior-backend",
                "requirements_text": "B.Sc. and Python required.",
            }])

        prompt = gemini_call.call_args.args[0]["contents"][0]["parts"][0]["text"]
        self.assertIn("is_matching_yoe to true", prompt)
        self.assertIn("STRICTLY REJECT titles", prompt)
        self.assertFalse(evaluations[0]["approved"])
        self.assertFalse(evaluations[0]["is_matching_yoe"])

    def test_ai_prompts_treat_unspecified_minimum_as_matching_and_seniority_as_rejected(self):
        audit_response = {
            "candidates": [{
                "content": {"parts": [{"text": json.dumps({
                    "second_judge": [{
                        "job_index": 1,
                        "is_single_specific_job": True,
                        "is_target_role": True,
                        "is_matching_yoe": True,
                        "is_matching_location": True,
                        "approved": True,
                        "reason": "",
                    }],
                    "recommendations": [],
                })}]}
            }]
        }
        with patch("job_crawler.call_gemini_with_retries", return_value=audit_response) as gemini_call:
            result = audit_run_with_ai(
                {"ai": {"approved_jobs": [{
                    "company": "Example",
                    "title": "Lead Backend Engineer",
                    "url": "https://example.com/job",
                }]}},
                {"min_yoe": 0, "max_yoe": 2, "location": "Israel", "target_roles": ["Backend"]},
            )
        prompt = gemini_call.call_args.args[0]["contents"][0]["parts"][0]["text"]
        self.assertIn("strictly reject any title indicating", prompt)
        self.assertIn("treat it as matching 0-2 YOE", prompt)
        self.assertFalse(result["second_judge"][0]["approved"])
        self.assertFalse(result["second_judge"][0]["is_matching_yoe"])

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

    def test_css_location_extraction_ignores_hebrew_job_metadata(self):
        soup = BeautifulSoup(
            """
            <div class="job-location">מספר משרה: 95400 / לינק ללכאן</div>
            <div class="location">### 12345</div>
            """,
            "html.parser",
        )

        _, page_location = extract_clean_text_from_soup(soup)

        self.assertEqual(page_location, "Unspecified")


if __name__ == "__main__":
    unittest.main()
