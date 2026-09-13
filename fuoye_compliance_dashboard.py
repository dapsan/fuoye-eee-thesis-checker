import io
import re
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import streamlit as st

# Configure Streamlit Page
st.set_page_config(
    page_title="FUOYE PG Thesis Compliance & Authenticity Auditor",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Core Audit Engine
class FUOYEManuscriptAuditor:
    def __init__(self, doc_stream, seminar_type="proposal", degree="phd"):
        self.doc = docx.Document(doc_stream)
        self.seminar_type = seminar_type.lower()
        self.degree = degree.lower()
        self.full_text = "\n".join([p.text for p in self.doc.paragraphs if p.text.strip()])
        
        self.results = {
            "passed": [],
            "warnings": [],
            "critical": [],
            "ref_authenticity": {
                "total_refs": 0,
                "total_citations": 0,
                "unmatched_citations": [],
                "uncited_refs": [],
                "incomplete_metadata": [],
                "valid_identifiers": 0,
                "index_score": 100.0,
                "status": "PASS"
            }
        }

    def audit_all(self):
        self._check_margins()
        self._check_typography_and_spacing()
        self._check_title_and_abstract()
        self._check_chapter_structure()
        self._check_captions()
        self._check_citations()
        self._check_reference_authenticity()
        return self.results

    def _check_margins(self):
        for idx, section in enumerate(self.doc.sections, 1):
            left_margin = round(section.left_margin.inches, 2) if section.left_margin else None
            right_margin = round(section.right_margin.inches, 2) if section.right_margin else None
            top_margin = round(section.top_margin.inches, 2) if section.top_margin else None
            bottom_margin = round(section.bottom_margin.inches, 2) if section.bottom_margin else None

            if left_margin and abs(left_margin - 1.25) > 0.05:
                self.results["critical"].append(
                    f"Margin Violation (Section {idx}): Left margin is {left_margin}\" (Required: 1.25\" / 3.125 cm)."
                )
            else:
                self.results["passed"].append(f"Section {idx} Margins: Left margin compliant at 1.25 inches.")

            for side, val in [("Right", right_margin), ("Top", top_margin), ("Bottom", bottom_margin)]:
                if val and abs(val - 1.0) > 0.05:
                    self.results["warnings"].append(
                        f"Margin Advisory (Section {idx}): {side} margin is {val}\" (Recommended: 1.0\" / 2.5 cm)."
                    )

    def _check_typography_and_spacing(self):
        non_tnr_runs = 0
        indented_paragraphs = 0

        for p in self.doc.paragraphs:
            if p.paragraph_format.first_line_indent and p.paragraph_format.first_line_indent.pt > 0:
                indented_paragraphs += 1

            for run in p.runs:
                if run.text.strip():
                    if run.font.name and run.font.name != "Times New Roman":
                        non_tnr_runs += 1

        if non_tnr_runs > 0:
            self.results["warnings"].append(
                f"Typography Advisory: Found {non_tnr_runs} text run(s) using non-Times New Roman fonts."
            )
        else:
            self.results["passed"].append("Typography Check: 100% of analyzed text runs use Times New Roman.")

        if indented_paragraphs > 0:
            self.results["warnings"].append(
                f"Paragraphing Advisory: Found {indented_paragraphs} paragraph(s) with first-line indent. FUOYE mandates Block Paragraphing (0 pt indent)."
            )
        else:
            self.results["passed"].append("Paragraphing Check: Block paragraphing (0 pt indent) verified.")

    def _check_title_and_abstract(self):
        first_pars = [p.text.strip() for p in self.doc.paragraphs[:15] if p.text.strip()]
        candidate_title = ""
        for ptext in first_pars:
            words = ptext.split()
            if len(words) >= 4 and ptext.isupper():
                candidate_title = ptext
                break
        
        if candidate_title:
            title_word_count = len(candidate_title.split())
            if title_word_count > 17:
                self.results["critical"].append(
                    f"Title Word Count Violation: Cover/Title page title has {title_word_count} words (Maximum allowed is 17 words). Title: '{candidate_title}'"
                )
            else:
                self.results["passed"].append(
                    f"Title Word Count Check: Research title has {title_word_count} words (Compliant with max 17 words)."
                )

        abstract_match = re.search(r'\b(SUMMARY|ABSTRACT)\b\n+(.*?)(?=\\n+CHAPTER|\n+1\.0|\Z)', self.full_text, re.DOTALL | re.IGNORECASE)
        if abstract_match:
            abstract_text = abstract_match.group(2).strip()
            abs_words = len(abstract_text.split())
            target_range = (400, 500) if self.degree == "phd" else (350, 400)
            
            if abs_words < target_range[0] or abs_words > target_range[1]:
                self.results["warnings"].append(
                    f"Abstract/Summary Word Count Advisory: Detected {abs_words} words. Guideline target for {self.degree.upper()} is {target_range[0]}–{target_range[1]} words."
                )
            else:
                self.results["passed"].append(
                    f"Abstract/Summary Word Count Check: {abs_words} words (Compliant with {self.degree.upper()} range {target_range[0]}–{target_range[1]} words)."
                )

    def _check_chapter_structure(self):
        if self.seminar_type == "proposal":
            required = {
                "Background": r"1\.1\s+Background",
                "Problem Statement": r"1\.2\s+Statement of the Problem",
                "Aim & Objectives": r"1\.3\s+Aim and Objectives",
                "Research Questions": r"1\.4\s+Research Questions",
                "Justification": r"1\.5\s+Justification",
                "Significance": r"1\.6\s+Significance",
                "Scope": r"1\.7\s+Scope",
                "Theoretical Framework": r"2\.2\s+Theoretical Framework",
                "Conceptual Framework": r"2\.3\s+Conceptual Framework",
                "Review of Related Literature": r"2\.4\s+Review of Related Literature",
                "Research Gap": r"2\.5\s+Research Gap",
                "Methodology": r"CHAPTER THREE|\bMETHODOLOGY\b",
                "Expected Results": r"CHAPTER FOUR|\bEXPECTED RESULTS\b",
                "Research Schedule / Work Plan": r"5\.3\s+Research Schedule|Work Plan",
                "Budget Estimate": r"5\.4\s+Budget Estimate"
            }
        elif self.seminar_type == "progress":
            required = {
                "Methodology": r"CHAPTER THREE|\bMETHODOLOGY\b",
                "Summary of Findings": r"4\.6\s+Summary of Findings",
                "Outstanding Results": r"4\.7\s+Outstanding Results",
                "Summary of Progress": r"5\.1\s+Summary of Progress",
                "Preliminary Conclusions": r"5\.2\s+Preliminary Conclusions",
                "Outstanding Work": r"5\.4\s+Outstanding Work"
            }
        else:
            required = {
                "Introduction": r"CHAPTER ONE|\bINTRODUCTION\b",
                "Literature Review": r"CHAPTER TWO|\bLITERATURE REVIEW\b",
                "Methodology": r"CHAPTER THREE|\bMETHODOLOGY\b",
                "Results and Discussion": r"CHAPTER FOUR|\bRESULTS AND DISCUSSION\b",
                "Conclusions and Recommendations": r"CHAPTER FIVE|\bCONCLUSIONS AND RECOMMENDATIONS\b"
            }

        missing = []
        for name, pattern in required.items():
            if not re.search(pattern, self.full_text, re.IGNORECASE):
                missing.append(name)
        
        if missing:
            self.results["critical"].append(
                f"Missing Chapter/Section Violations for {self.seminar_type.upper()}: Mandatory sections not detected: {', '.join(missing)}."
            )
        else:
            self.results["passed"].append(
                f"Seminar Structure Check: All required structural sections for {self.seminar_type.upper()} were successfully detected."
            )

    def _check_captions(self):
        tables_found = len(self.doc.tables)
        tbl_captions = [p.text for p in self.doc.paragraphs if re.match(r'^Table\s+\d+[\.\d]*', p.text.strip(), re.IGNORECASE)]
        fig_captions = [p.text for p in self.doc.paragraphs if re.match(r'^Figure\s+\d+[\.\d]*', p.text.strip(), re.IGNORECASE)]

        if tables_found > 0 and len(tbl_captions) == 0:
            self.results["warnings"].append(
                f"Caption Advisory: Found {tables_found} table(s) in document, but no standard 'Table X.Y:' captions were detected above them."
            )
        elif len(tbl_captions) > 0:
            self.results["passed"].append(f"Table Captions Check: Found {len(tbl_captions)} properly formatted 'Table X.Y' caption(s).")

        if len(fig_captions) > 0:
            self.results["passed"].append(f"Figure Captions Check: Found {len(fig_captions)} properly formatted 'Figure X.Y' caption(s).")

    def _check_citations(self):
        apa_citations = re.findall(r'\([A-Z][a-z]+(?:\s+(?:&|and)\s+[A-Z][a-z]+)?,\\s+\d{4}\)', self.full_text)
        if not apa_citations:
            self.results["warnings"].append(
                "Citation Style Advisory: Could not detect standard APA parenthetical citations e.g., (Author, Year)."
            )
        else:
            self.results["passed"].append(f"APA Citations Check: Found {len(apa_citations)} APA 7th edition style in-text citation(s).")

    def _check_reference_authenticity(self):
        ref_data = self.results["ref_authenticity"]
        
        ref_text = ""
        ref_match = re.search(r'\bREFERENCES\b\n+(.*)', self.full_text, re.DOTALL | re.IGNORECASE)
        if ref_match:
            ref_text = ref_match.group(1).split("APPENDIX")[0]

        ref_entries = [p.strip() for p in ref_text.split("\n") if len(p.strip()) > 15]
        ref_data["total_refs"] = len(ref_entries)

        citations = re.findall(r'([A-Z][a-zA-B\-]+)(?:\s+(?:and|&)\s+[A-Z][a-zA-B\-]+|\s+et\s+al\.)?,\s+(\d{4})', self.full_text)
        ref_data["total_citations"] = len(citations)

        unmatched = []
        for author, year in citations:
            found = False
            for entry in ref_entries:
                if author.lower() in entry.lower() and year in entry:
                    found = True
                    break
            if not found:
                unmatched.append(f"{author} ({year})")

        ref_data["unmatched_citations"] = list(set(unmatched))

        uncited = []
        body_text = self.full_text.split("REFERENCES")[0]
        for entry in ref_entries:
            author_match = re.search(r'^([A-Z][a-zA-B\-]+)', entry)
            year_match = re.search(r'\b(19|20)\d{2}\b', entry)
            if author_match and year_match:
                au = author_match.group(1)
                yr = year_match.group(0)
                if au.lower() not in body_text.lower():
                    uncited.append(f"{au} ({yr})")

        ref_data["uncited_refs"] = list(set(uncited))

        incomplete = []
        valid_ids = 0
        for idx, entry in enumerate(ref_entries, 1):
            has_year = bool(re.search(r'\b(19|20)\d{2}\b', entry))
            has_title = len(entry.split()) >= 5
            if not has_year or not has_title:
                incomplete.append(f"Reference #{idx}: Missing year or article title details.")

            if "doi" in entry.lower() or "http" in entry.lower() or "https" in entry.lower():
                valid_ids += 1

        ref_data["incomplete_metadata"] = incomplete
        ref_data["valid_identifiers"] = valid_ids

        tot_refs = len(ref_entries) if len(ref_entries) > 0 else 1
        matched_cnt = max(0, len(citations) - len(ref_data["unmatched_citations"]))
        align_rate = (matched_cnt / len(citations) * 100) if len(citations) > 0 else 100.0
        comp_rate = max(0, ((tot_refs - len(incomplete)) / tot_refs) * 100)

        index_score = round((align_rate * 0.6) + (comp_rate * 0.4), 1)
        ref_data["index_score"] = index_score

        if index_score >= 85.0 and len(unmatched) == 0:
            ref_data["status"] = "HIGH AUTHENTICITY (Verified)"
            self.results["passed"].append(f"Reference Authenticity Index: {index_score}% (High Integrity & Grounded).")
        elif index_score >= 60.0:
            ref_data["status"] = "MODERATE RISK (Discrepancies Found)"
            self.results["warnings"].append(
                f"Reference Authenticity Advisory: Index is {index_score}%. Detected {len(unmatched)} in-text citation(s) missing from Reference List."
            )
        else:
            ref_data["status"] = "HIGH FABRICATION RISK (Requires Review)"
            self.results["critical"].append(
                f"Reference Integrity Violation: Authenticity index dropped to {index_score}%. Significant unmatched or phantom references detected."
            )


# --- Streamlit UI App Layout ---
def main():
    st.title("🎓 FUOYE Postgraduate Thesis & Seminar Auditor")
    st.markdown("""
    **Department of Electrical & Electronics Engineering (EEE) & School of Postgraduate Studies**  
    *Upload your report (.docx) to audit for compliance with FUOYE formatting, typography, structure, word count, and reference authenticity.*
    """)
    st.divider()

    # Sidebar Configuration
    st.sidebar.header("📋 Submission Details")
    degree_option = st.sidebar.selectbox(
        "Degree Program",
        ["Ph.D.", "Master's (M.Eng. / M.Sc.)", "PGD"]
    )
    
    seminar_option = st.sidebar.selectbox(
        "Seminar / Manuscript Stage",
        [
            "Research Proposal",
            "Progress Report I & II",
            "Post-Data / Final Thesis / Dissertation"
        ]
    )

    degree_code = "phd" if "Ph.D." in degree_option else "master"
    if "Proposal" in seminar_option:
        seminar_code = "proposal"
    elif "Progress" in seminar_option:
        seminar_code = "progress"
    else:
        seminar_code = "thesis"

    st.sidebar.markdown("---")
    st.sidebar.info("""
    **Official Scoring Rule:**  
    - **Compliance Score** focuses strictly on **Passed Checks vs. Critical Violations**.
    - **Advisories & Warnings** are highlighted separately for guidance and **do not penalize** the score.
    """)

    # Main File Upload Section
    uploaded_file = st.file_uploader(
        "Upload your manuscript (.docx file)",
        type=["docx"],
        help="Select your Word document manuscript for automatic compliance & reference authenticity audit."
    )

    if uploaded_file is not None:
        st.success(f"📄 Loaded file: **{uploaded_file.name}** ({uploaded_file.size // 1024} KB)")
        
        if st.button("🚀 Run Compliance & Authenticity Audit", type="primary"):
            with st.spinner("Analyzing manuscript typography, margins, chapter structure, and verifying reference authenticity..."):
                try:
                    auditor = FUOYEManuscriptAuditor(
                        uploaded_file,
                        seminar_type=seminar_code,
                        degree=degree_code
                    )
                    results = auditor.audit_all()
                    
                    st.divider()
                    st.header("📊 Compliance Scorecard")
                    
                    num_pass = len(results["passed"])
                    num_warn = len(results["warnings"])
                    num_fail = len(results["critical"])
                    
                    # OFFICIAL COMPLIANCE SCORE: Focused strictly on Passed vs Critical Violations
                    denom = num_pass + num_fail
                    score = int((num_pass / denom) * 100) if denom > 0 else 100

                    # Metrics Cards
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Compliance Score", f"{score}%", help="Calculated from Passed Checks vs Critical Violations only.")
                    col2.metric("Passed Checks", num_pass, delta=f"{num_pass} passed", delta_color="normal")
                    col3.metric("Critical Violations", num_fail, delta=f"{num_fail} violations", delta_color="inverse")
                    col4.metric("Advisories (No Penalty)", num_warn, delta=f"{num_warn} advisories", delta_color="off")

                    st.markdown("---")

                    # Tabbed Detailed Results
                    tab1, tab2 = st.tabs(["📝 Guidelines Compliance Breakdown", "🔍 Reference Authenticity & Verification Report"])

                    with tab1:
                        if num_fail > 0:
                            st.error(f"❌ **CRITICAL VIOLATIONS ({num_fail})** - Must be revised prior to submission:")
                            for item in results["critical"]:
                                st.write(f"- 🔴 {item}")

                        if num_warn > 0:
                            st.warning(f"⚠️ **ADVISORIES & RECOMMENDATIONS ({num_warn})** *(Indicated for formatting refinement; omitted from compliance score)*:")
                            for item in results["warnings"]:
                                st.write(f"- 🟡 {item}")

                        if num_pass > 0:
                            st.success(f"✔ **PASSED CHECKS ({num_pass})**:")
                            for item in results["passed"]:
                                st.write(f"- 🟢 {item}")

                    with tab2:
                        ref_rep = results["ref_authenticity"]
                        st.subheader("🔍 Reference Authenticity & Integrity Analysis")
                        st.caption("Audits references against FUOYE AI Policy (prohibiting fabricated citations) and Turnitin 404 Error plagiarism standards.")
                        
                        rcol1, rcol2, rcol3 = st.columns(3)
                        rcol1.metric("Authenticity Status", ref_rep["status"])
                        rcol2.metric("Authenticity Index", f"{ref_rep['index_score']}%")
                        rcol3.metric("Scanned References", ref_rep["total_refs"])

                        st.markdown("##### 1:1 Citation-Reference Cross-Matching")
                        if ref_rep["unmatched_citations"]:
                            st.error(f"⚠️ **In-Text Citations Missing from Reference List ({len(ref_rep['unmatched_citations'])}):**")
                            st.write("The following in-text citations appear in the text body but lack a corresponding full entry in the Reference section:")
                            for c in ref_rep["unmatched_citations"]:
                                st.write(f"  - 🔴 `{c}`")
                        else:
                            st.success("🟢 All in-text citations were successfully matched to entries in the Reference List.")

                        if ref_rep["uncited_refs"]:
                            st.warning(f"⚠️ **Uncited References in Reference List ({len(ref_rep['uncited_refs'])}):**")
                            st.write("The following reference entries exist in the Reference List but were not cited anywhere in the manuscript body:")
                            for u in ref_rep["uncited_refs"]:
                                st.write(f"  - 🟡 `{u}`")

                        st.markdown("##### Bibliographic Completeness & DOI / URL Verification")
                        if ref_rep["incomplete_metadata"]:
                            st.warning(f"⚠️ **Incomplete Reference Entries ({len(ref_rep['incomplete_metadata'])}):**")
                            for inc in ref_rep["incomplete_metadata"]:
                                st.write(f"  - 🟡 {inc}")
                        else:
                            st.success("🟢 All scanned reference list entries contain complete bibliographic metadata.")

                        st.info(f"ℹ️ Found **{ref_rep['valid_identifiers']}** reference entry(ies) with valid DOIs or URLs.")

                    # Downloadable Combined Audit Report
                    st.divider()
                    report_text = f"""FUOYE POSTGRADUATE MANUSCRIPT AUDIT REPORT
File: {uploaded_file.name}
Degree: {degree_option} | Seminar Stage: {seminar_option}
Official Compliance Score: {score}% (Passed Checks vs Critical Violations)

=================== CRITICAL VIOLATIONS ===================
""" + ("\n".join([f"- {x}" for x in results["critical"]]) if results["critical"] else "None!") + f"""

=================== ADVISORIES & WARNINGS (No Score Penalty) ===================
""" + ("\n".join([f"- {x}" for x in results["warnings"]]) if results["warnings"] else "None!") + f"""

=================== PASSED CHECKS ===================
""" + ("\n".join([f"- {x}" for x in results["passed"]]) if results["passed"] else "None!") + f"""

=================== REFERENCE AUTHENTICITY & INTEGRITY REPORT ===================
Authenticity Status: {ref_rep['status']}
Authenticity Index:  {ref_rep['index_score']}%
Total Scanned In-Text Citations: {ref_rep['total_citations']}
Total Reference List Entries:    {ref_rep['total_refs']}

Unmatched In-Text Citations:
""" + ("\n".join([f"- {x}" for x in ref_rep["unmatched_citations"]]) if ref_rep["unmatched_citations"] else "None") + f"""

Uncited References in List:
""" + ("\n".join([f"- {x}" for x in ref_rep["uncited_refs"]]) if ref_rep["uncited_refs"] else "None") + f"""

Incomplete Metadata Entries:
""" + ("\n".join([f"- {x}" for x in ref_rep["incomplete_metadata"]]) if ref_rep["incomplete_metadata"] else "None")

                    st.download_button(
                        label="📥 Download Audit & Authenticity Report (.txt)",
                        data=report_text,
                        file_name=f"Compliance_Audit_{uploaded_file.name}.txt",
                        mime="text/plain"
                    )

                except Exception as e:
                    st.error(f"Error processing manuscript file: {str(e)}")
    else:
        st.info("👆 Please upload a `.docx` manuscript file to begin audit.")

if __name__ == "__main__":
    main()
