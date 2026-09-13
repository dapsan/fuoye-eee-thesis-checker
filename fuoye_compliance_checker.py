#!/usr/bin/env python3
"""
FUOYE Postgraduate Seminar Report & Thesis Compliance Checker
Department of Electrical and Electronics Engineering (EEE) & SPGS
Federal University Oye-Ekiti (FUOYE)

Automated script for parsing .docx manuscripts and verifying compliance against:
1. Proposal Report Guidelines (FUOYE EEE Template)
2. Progress Report I & II Guidelines (FUOYE EEE Template)
3. Post-Data & Final Thesis/Dissertation Guidelines (FUOYE SPGS & EEE Templates)
4. Reference Authenticity & Integrity Verification
"""

import sys
import os
import re
import argparse
import docx
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

class FUOYEComplianceAgent:
    def __init__(self, doc_path, seminar_type="proposal", degree_type="phd"):
        self.doc_path = doc_path
        self.seminar_type = seminar_type.lower()
        self.degree_type = degree_type.lower()
        
        if not os.path.exists(doc_path):
            raise FileNotFoundError(f"Document file not found at '{doc_path}'.")
            
        self.doc = docx.Document(doc_path)
        self.full_text = "\n".join([p.text for p in self.doc.paragraphs])
        
        self.report = {
            "seminar_type": self.seminar_type.upper(),
            "degree_type": self.degree_type.upper(),
            "passed": [],
            "warnings": [],
            "critical_violations": [],
            "stats": {},
            "compliance_score": 100,
            "reference_authenticity_report": {
                "total_references_found": 0,
                "total_in_text_citations": 0,
                "matched_citations": 0,
                "unmatched_in_text": [],
                "uncited_references": [],
                "incomplete_metadata": [],
                "valid_doi_url_count": 0,
                "invalid_doi_url_count": 0,
                "authenticity_index": 100.0,
                "status": "PASS"
            }
        }

    def run_full_audit(self):
        self._audit_margins()
        self._audit_typography_and_layout()
        self._audit_title()
        self._audit_summary_or_abstract()
        self._audit_structure()
        self._audit_captions_and_equations()
        self._audit_referencing()
        self._audit_reference_authenticity()
        self._calculate_compliance_score()
        return self.report

    def _calculate_compliance_score(self):
        """Calculates compliance score focusing ONLY on Passed Checks vs Critical Violations.
        Warnings/Advisories are indicated separately and do NOT penalize the compliance score."""
        num_pass = len(self.report["passed"])
        num_critical = len(self.report["critical_violations"])
        total_score_checks = num_pass + num_critical
        if total_score_checks > 0:
            self.report["compliance_score"] = int((num_pass / total_score_checks) * 100)
        else:
            self.report["compliance_score"] = 100

    def _audit_margins(self):
        """Audits page margins against 3.125cm (1.25 in) left and 2.5cm (1.0 in) top/right/bottom."""
        for idx, sec in enumerate(self.doc.sections):
            left_in = round(sec.left_margin.inches, 2) if sec.left_margin else 0.0
            right_in = round(sec.right_margin.inches, 2) if sec.right_margin else 0.0
            top_in = round(sec.top_margin.inches, 2) if sec.top_margin else 0.0
            bottom_in = round(sec.bottom_margin.inches, 2) if sec.bottom_margin else 0.0

            margin_issues = []
            if not (1.20 <= left_in <= 1.30):
                margin_issues.append(f"Left margin is {left_in}\" (Required: 1.25\" / 3.125 cm)")
            if not (0.95 <= right_in <= 1.05):
                margin_issues.append(f"Right margin is {right_in}\" (Required: 1.0\" / 2.5 cm)")
            if not (0.95 <= top_in <= 1.05):
                margin_issues.append(f"Top margin is {top_in}\" (Required: 1.0\" / 2.5 cm)")
            if not (0.95 <= bottom_in <= 1.05):
                margin_issues.append(f"Bottom margin is {bottom_in}\" (Required: 1.0\" / 2.5 cm)")

            if not margin_issues:
                self.report["passed"].append(f"Page Margins (Section {idx+1}): Compliant (Left 1.25\", Right/Top/Bottom 1.0\").")
            else:
                self.report["critical_violations"].append(
                    f"Margin Violation in Section {idx+1}: " + ", ".join(margin_issues)
                )

    def _audit_typography_and_layout(self):
        """Audits font family (Times New Roman), font sizes, and block paragraphing."""
        non_tnr_runs = 0
        indented_paras = 0

        for p in self.doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue

            if p.paragraph_format.first_line_indent and p.paragraph_format.first_line_indent.pt > 0:
                indented_paras += 1

            for run in p.runs:
                if run.font.name and run.font.name != "Times New Roman":
                    non_tnr_runs += 1

        if non_tnr_runs == 0:
            self.report["passed"].append("Typography: All scanned text runs strictly use 'Times New Roman'.")
        else:
            self.report["warnings"].append(f"Typography Advisory: Detected {non_tnr_runs} text run(s) using non-standard fonts.")

        if indented_paras == 0:
            self.report["passed"].append("Paragraphing: 100% compliant Block Paragraphing (no indentation).")
        else:
            self.report["warnings"].append(
                f"Paragraphing Advisory: Found {indented_paras} paragraph(s) with first-line indent. FUOYE mandates Block Paragraphing."
            )

    def _audit_title(self):
        """Audits research title length (Max 17 words)."""
        title_candidates = []
        for p in self.doc.paragraphs[:15]:
            text = p.text.strip()
            if text and not any(kw in text.upper() for kw in ["FEDERAL UNIVERSITY", "DEPARTMENT OF", "REPORT TEMPLATE", "COVER PAGE", "TITLE PAGE", "SUBMITTED TO"]):
                if len(text.split()) > 3:
                    title_candidates.append(text)
                    break

        if title_candidates:
            title = title_candidates[0]
            word_count = len(title.split())
            self.report["stats"]["title_word_count"] = word_count
            if word_count <= 17:
                self.report["passed"].append(f"Title Word Count: '{title[:50]}...' contains {word_count} words (Compliant: Max 17 allowed).")
            else:
                self.report["critical_violations"].append(
                    f"Title Word Count Exceeded: Research title has {word_count} words (Maximum allowed: 17 words). Title: '{title}'"
                )
        else:
            self.report["warnings"].append("Title Check: Could not automatically extract title block on page 1–2.")

    def _audit_summary_or_abstract(self):
        """Audits word limits and narrative formatting of Summary / Abstract."""
        header_name = "SUMMARY" if self.seminar_type in ["proposal", "progress"] else "ABSTRACT"
        
        summary_paras = []
        capture = False

        for p in self.doc.paragraphs:
            text = p.text.strip()
            if text.upper() == header_name:
                capture = True
                continue
            if capture:
                if text.startswith("CHAPTER") or text.upper().startswith("TABLE OF CONTENTS") or text.upper().startswith("DECLARATION"):
                    break
                if text:
                    summary_paras.append(text)

        full_summary_text = " ".join(summary_paras)
        word_count = len(full_summary_text.split())
        self.report["stats"]["summary_word_count"] = word_count

        min_w, max_w = (400, 500) if self.degree_type == "phd" else (350, 400)

        if word_count == 0:
            self.report["critical_violations"].append(f"{header_name} Missing: Could not locate section headed '{header_name}'.")
        elif min_w <= word_count <= max_w:
            self.report["passed"].append(f"{header_name} Word Count: {word_count} words (Compliant with {self.degree_type.upper()} range of {min_w}–{max_w} words).")
        else:
            self.report["warnings"].append(
                f"{header_name} Word Count Advisory: {word_count} words found (Target for {self.degree_type.upper()} is {min_w}–{max_w} words)."
            )

        if len(summary_paras) > 1 and self.seminar_type in ["proposal", "progress"]:
            self.report["warnings"].append(
                f"{header_name} Structure Advisory: Written across {len(summary_paras)} paragraphs. Guidelines mandate a single coherent narrative block."
            )

    def _audit_structure(self):
        """Audits chapter structure depending on seminar type."""
        if self.seminar_type == "proposal":
            required = {
                "1.1": "Background", "1.2": "Statement of the Problem", "1.3": "Aim and Objectives",
                "1.4": "Research Questions", "1.5": "Justification", "1.6": "Significance", "1.7": "Scope",
                "2.2": "Theoretical Framework", "2.3": "Conceptual Framework", "2.4": "Review of Related Literature", "2.5": "Research Gap",
                "5.3": "Research Schedule", "5.4": "Budget Estimate"
            }
            missing = [f"{k} {v}" for k, v in required.items() if k not in self.full_text]
            if not missing:
                self.report["passed"].append("Proposal Structure: All required sections (including Research Schedule and Budget Estimate) were detected.")
            else:
                self.report["critical_violations"].append(f"Proposal Structure Violation: Missing mandatory section(s): {', '.join(missing)}")

        elif self.seminar_type == "progress":
            has_completed = "Work Completed" in self.full_text or "WORK COMPLETED" in self.full_text.upper()
            has_outstanding = "Outstanding Work" in self.full_text or "OUTSTANDING WORK" in self.full_text.upper() or "5.4 Outstanding Work" in self.full_text
            
            if has_completed and has_outstanding:
                self.report["passed"].append("Progress Report Structure: Explicit distinction between 'Work Completed' and 'Outstanding Work' verified.")
            else:
                self.report["critical_violations"].append(
                    "Progress Report Violation: Missing mandatory 'Work Completed' or 'Outstanding Work' subheadings in Chapter 3/4/5."
                )

        elif self.seminar_type == "thesis":
            total_words = len(self.full_text.split())
            self.report["stats"]["total_word_count"] = total_words
            limit = 90000 if self.degree_type == "phd" else 45000
            if total_words <= limit:
                self.report["passed"].append(f"Thesis Word Limit: {total_words} words (Compliant: Max limit is {limit} words for {self.degree_type.upper()}).")
            else:
                self.report["critical_violations"].append(
                    f"Thesis Word Limit Violation: Manuscript word count ({total_words}) exceeds maximum permitted limit ({limit} words)."
                )

    def _audit_captions_and_equations(self):
        """Audits table/figure captions and equation numbering."""
        tbl_captions = re.findall(r'(Table\s+\d+\.\d+:?[^\n]*)', self.full_text, re.IGNORECASE)
        fig_captions = re.findall(r'(Figure\s+\d+\.\d+:?[^\n]*)', self.full_text, re.IGNORECASE)
        equations = re.findall(r'\(\d+\.\d+\)', self.full_text)

        if tbl_captions:
            self.report["passed"].append(f"Table Captions: Found {len(tbl_captions)} properly formatted 'Table X.Y' caption(s).")
        if fig_captions:
            self.report["passed"].append(f"Figure Captions: Found {len(fig_captions)} properly formatted 'Figure X.Y' caption(s).")
        if equations:
            self.report["passed"].append(f"Equation Formatting: Found {len(equations)} chapter-numbered equation reference(s) e.g., (3.1).")

    def _audit_referencing(self):
        """Audits APA referencing and presence of References section."""
        if "REFERENCES" in self.full_text.upper():
            self.report["passed"].append("Reference List: 'REFERENCES' chapter heading detected.")
        else:
            self.report["critical_violations"].append("Reference List Violation: Missing 'REFERENCES' section.")

        apa_citations = re.findall(r'\([A-Z][a-z]+(?:\s+(?:and|&)\s+[A-Z][a-z]+|\s+et\s+al\.)?,\s+\d{4}\)', self.full_text)
        if apa_citations:
            self.report["passed"].append(f"In-Text Citations: Detected {len(apa_citations)} APA 7th Edition formatted citation(s).")
        else:
            self.report["warnings"].append("In-Text Citations Advisory: Could not detect standard APA parenthetical citations e.g., (Author, Year).")

    def _audit_reference_authenticity(self):
        """Performs a comprehensive verification report on the authenticity and integrity of references."""
        ref_report = self.report["reference_authenticity_report"]
        
        # 1. Extract Reference List Items
        ref_section_text = ""
        ref_match = re.search(r'\bREFERENCES\b\n+(.*)', self.full_text, re.DOTALL | re.IGNORECASE)
        if ref_match:
            ref_section_text = ref_match.group(1).split("APPENDIX")[0]
        
        ref_entries = [line.strip() for line in ref_section_text.split("\n") if len(line.strip()) > 15]
        ref_report["total_references_found"] = len(ref_entries)

        # 2. Extract In-Text Parenthetical & Narrative Citations
        in_text_authors = []
        # Pattern for (Author, Year) or (Author & Author, Year) or (Author et al., Year)
        citations = re.findall(r'([A-Z][a-zA-B\-]+)(?:\s+(?:and|&)\s+[A-Z][a-zA-B\-]+|\s+et\s+al\.)?,\s+(\d{4})', self.full_text)
        ref_report["total_in_text_citations"] = len(citations)

        # Cross-Matching In-Text vs Reference Entries
        unmatched_citations = []
        for author, year in citations:
            matched = False
            for entry in ref_entries:
                if author.lower() in entry.lower() and year in entry:
                    matched = True
                    break
            if not matched:
                unmatched_citations.append(f"{author} ({year})")
        
        ref_report["unmatched_in_text"] = list(set(unmatched_citations))

        # Check for Uncited Entries in References List
        uncited_refs = []
        for entry in ref_entries:
            year_match = re.search(r'\((\d{4})\)|\b(\d{4})\b', entry)
            author_match = re.search(r'^([A-Z][a-zA-B\-]+)', entry)
            if year_match and author_match:
                yr = year_match.group(1) or year_match.group(2)
                au = author_match.group(1)
                # search in full_text outside REFERENCES section
                body_without_refs = self.full_text.split("REFERENCES")[0]
                if au.lower() not in body_without_refs.lower():
                    uncited_refs.append(f"{au} ({yr})")

        ref_report["uncited_references"] = list(set(uncited_refs))

        # 3. Check Bibliographic Metadata Completeness & Fabrication Risk
        incomplete_entries = []
        valid_doi_url = 0
        invalid_doi_url = 0

        for idx, entry in enumerate(ref_entries, 1):
            has_year = bool(re.search(r'\b(19|20)\d{2}\b', entry))
            has_title = len(entry.split()) >= 6
            has_doi_or_url = "doi" in entry.lower() or "http" in entry.lower() or "https" in entry.lower()

            if not has_year or not has_title:
                incomplete_entries.append(f"Entry #{idx}: Incomplete bibliographic details (Missing year or title).")

            # Check DOI / URL validity
            doi_matches = re.findall(r'10\.\d{4,9}/[-._;()/:A-Z0-9]+', entry, re.IGNORECASE)
            url_matches = re.findall(r'https?://[^\s]+', entry)

            if doi_matches or url_matches:
                valid_doi_url += 1
            else:
                # Missing digital identifier where applicable
                pass

        ref_report["incomplete_metadata"] = incomplete_entries
        ref_report["valid_doi_url_count"] = valid_doi_url

        # Compute Authenticity Index (%)
        total_refs = len(ref_entries) if len(ref_entries) > 0 else 1
        matched_citations_count = len(citations) - len(ref_report["unmatched_in_text"])
        alignment_rate = max(0, (matched_citations_count / len(citations)) * 100) if len(citations) > 0 else 100.0
        completeness_rate = max(0, ((total_refs - len(incomplete_entries)) / total_refs) * 100)
        
        auth_index = round((alignment_rate * 0.6) + (completeness_rate * 0.4), 1)
        ref_report["authenticity_index"] = auth_index

        if auth_index >= 85.0 and len(ref_report["unmatched_in_text"]) == 0:
            ref_report["status"] = "HIGH AUTHENTICITY (Verified)"
            self.report["passed"].append(f"Reference Authenticity Check: Authenticity Index is {auth_index}% (Verified & Grounded).")
        elif auth_index >= 60.0:
            ref_report["status"] = "MODERATE RISK (Minor Discrepancies Detected)"
            self.report["warnings"].append(
                f"Reference Authenticity Advisory: Index is {auth_index}%. Found {len(ref_report['unmatched_in_text'])} in-text citation(s) missing from Reference List."
            )
        else:
            ref_report["status"] = "HIGH FABRICATION RISK (Requires Review)"
            self.report["critical_violations"].append(
                f"Reference Integrity Violation: Authenticity Index dropped to {auth_index}%. Multiple uncited or phantom citations detected."
            )

    def print_summary(self):
        print("\n" + "="*65)
        print("  FUOYE POSTGRADUATE MANUSCRIPT COMPLIANCE AUDIT REPORT")
        print(f"  Target: {self.seminar_type.upper()} | Degree: {self.degree_type.upper()}")
        print(f"  OFFICIAL COMPLIANCE SCORE: {self.report['compliance_score']}%")
        print("="*65 + "\n")

        print("✔ PASSED CHECKS (Included in Score):")
        for item in self.report["passed"]:
            print(f"  [PASS] {item}")

        if self.report["critical_violations"]:
            print("\n❌ CRITICAL VIOLATIONS (Included in Score - Requires Correction):")
            for item in self.report["critical_violations"]:
                print(f"  [FAIL] {item}")
        else:
            print("\n🎉 CONGRATULATIONS: No Critical Violations Detected!")

        if self.report["warnings"]:
            print("\n⚠️ ADVISORIES & WARNINGS (Indicated for Guidance - Excluded from Score):")
            for item in self.report["warnings"]:
                print(f"  [WARN] {item}")

        # Separate Reference Authenticity Report
        ref_rep = self.report["reference_authenticity_report"]
        print("\n" + "-"*65)
        print("  🔍 SEPARATE AUDIT: REFERENCE AUTHENTICITY & INTEGRITY REPORT")
        print(f"  Authenticity Status: {ref_rep['status']}")
        print(f"  Authenticity Index:  {ref_rep['authenticity_index']}%")
        print(f"  Total In-Text Citations Scanned: {ref_rep['total_in_text_citations']}")
        print(f"  Total Reference Entries Found:  {ref_rep['total_references_found']}")
        if ref_rep["unmatched_in_text"]:
            print(f"  ⚠️ Unmatched In-Text Citations (Missing from References): {', '.join(ref_rep['unmatched_in_text'])}")
        if ref_rep["uncited_references"]:
            print(f"  ⚠️ Uncited References in List: {', '.join(ref_rep['uncited_references'])}")
        if ref_rep["incomplete_metadata"]:
            print(f"  ⚠️ Entries with Incomplete Bibliographic Details: {len(ref_rep['incomplete_metadata'])}")
        print("-"*65 + "\n")

def main():
    parser = argparse.ArgumentParser(description="FUOYE Postgraduate Seminar Compliance Checker")
    parser.add_argument("docx_path", help="Path to the .docx manuscript file")
    parser.add_argument("--type", choices=["proposal", "progress", "thesis"], default="proposal", help="Seminar type")
    parser.add_argument("--degree", choices=["phd", "master"], default="phd", help="Degree program")
    args = parser.parse_args()

    agent = FUOYEComplianceAgent(args.docx_path, seminar_type=args.type, degree_type=args.degree)
    agent.run_full_audit()
    agent.print_summary()

if __name__ == "__main__":
    main()
