import io
import re
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import streamlit as st

# Configure Streamlit Page
st.set_page_config(
    page_title="FUOYE PG Thesis Compliance Checker",
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
        self.full_text = "\n".join([p.text for p in self.doc.paragraphs])
        
        self.results = {
            "passed": [],
            "warnings": [],  # Advisories and Critical Warnings (e.g. margins) - excluded from score
            "critical": [],  # Critical structural violations - included in score
            "reference_authenticity": {}
        }

    def audit_all(self):
        self._check_margins()
        self._check_typography_and_spacing()
        self._check_title_and_abstract()
        self._check_chapter_structure()
        self._check_captions()
        self._check_citations_and_authenticity()
        return self.results

    def _check_margins(self):
        """
        Audits page margins against 3.125cm (1.25 in) left and 2.5cm (1.0 in) top/right/bottom.
        Margin violations are classified as Critical Warnings under warnings so they do NOT reduce score.
        """
        for idx, section in enumerate(self.doc.sections, 1):
            left_margin = round(section.left_margin.inches, 2) if section.left_margin else None
            right_margin = round(section.right_margin.inches, 2) if section.right_margin else None
            top_margin = round(section.top_margin.inches, 2) if section.top_margin else None
            bottom_margin = round(section.bottom_margin.inches, 2) if section.bottom_margin else None

            if left_margin and abs(left_margin - 1.25) > 0.05:
                self.results["warnings"].append(
                    f"Critical Warning - Margin Violation (Section {idx}): Left margin is {left_margin}\" (Required: 1.25\" / 3.125 cm)."
                )
            else:
                self.results["passed"].append(f"Section {idx} Margins: Left margin compliant at 1.25 inches.")

            for side, val in [("Right", right_margin), ("Top", top_margin), ("Bottom", bottom_margin)]:
                if val and abs(val - 1.0) > 0.05:
                    self.results["warnings"].append(
                        f"Critical Warning - Margin Violation (Section {idx}): {side} margin is {val}\" (Required: 1.0\" / 2.5 cm)."
                    )

    def _check_typography_and_spacing(self):
        non_tnr_runs = 0
        indented_paragraphs = 0
        total_runs = 0

        for p in self.doc.paragraphs:
            if p.paragraph_format.first_line_indent and p.paragraph_format.first_line_indent.pt > 0:
                indented_paragraphs += 1

            for run in p.runs:
                if run.text.strip():
                    total_runs += 1
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

        # Check Summary / Abstract
        abstract_match = re.search(r'\b(SUMMARY|ABSTRACT)\b\n+(.*?)(?=\n+CHAPTER|\n+1\.0|\Z)', self.full_text, re.DOTALL | re.IGNORECASE)
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
        else: # thesis / postdata
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

    def _check_citations_and_authenticity(self):
        parenthetical_citations = re.findall(r'\(([A-Z][a-zA-Z\s\-&]+),\s*(\d{4})\)', self.full_text)
        narrative_citations = re.findall(r'\b([A-Z][a-zA-Z\s\-]+)\s*\((\d{4})\)', self.full_text)
        
        all_cited_authors = set()
        for authors, yr in parenthetical_citations + narrative_citations:
            cleaned_author = authors.replace("&", "").replace("and", "").replace("et al.", "").strip().split()[0]
            if len(cleaned_author) > 2:
                all_cited_authors.add((cleaned_author, yr))

        ref_section = ""
        if "REFERENCES" in self.full_text.upper():
            parts = re.split(r'\bREFERENCES\b', self.full_text, flags=re.IGNORECASE)
            if len(parts) > 1:
                ref_section = parts[-1]

        ref_entries = [p.strip() for p in ref_section.split("\n") if len(p.strip()) > 15]

        matched_citations = 0
        unmatched_citations = []
        for author, yr in all_cited_authors:
            found = False
            for ref in ref_entries:
                if author.lower() in ref.lower() and yr in ref:
                    found = True
                    break
            if found:
                matched_citations += 1
            else:
                unmatched_citations.append(f"{author} ({yr})")

        total_citations = len(all_cited_authors)
        authenticity_rate = (matched_citations / total_citations * 100) if total_citations > 0 else 100.0

        if total_citations > 0:
            self.results["passed"].append(f"APA Citations Check: Detected {total_citations} unique APA formatted citation(s).")
        else:
            self.results["warnings"].append("Citation Style Advisory: Could not detect standard APA parenthetical citations e.g., (Author, Year).")

        status_text = "HIGH AUTHENTICITY (Verified)" if authenticity_rate >= 85 else ("MODERATE RISK" if authenticity_rate >= 60 else "HIGH FABRICATION RISK")

        self.results["reference_authenticity"] = {
            "total_in_text_citations": total_citations,
            "total_ref_entries": len(ref_entries),
            "matched_citations": matched_citations,
            "unmatched_citations": unmatched_citations,
            "authenticity_index": round(authenticity_rate, 1),
            "status": status_text
        }


# --- Streamlit UI App Layout ---
def main():
    st.title("🎓 FUOYE Postgraduate Thesis & Seminar Auditor")
    st.markdown("""
    **Department of Electrical & Electronics Engineering (EEE) & School of Postgraduate Studies**  
    *Upload your report (.docx) to audit for compliance with FUOYE formatting, typography, structure, and word count guidelines.*
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
    **Required Manuscript Specs:**
    - **Font**: Times New Roman 12 pt (1.5 line spacing)
    - **Margins**: Left 1.25", Top/Right/Bottom 1.0"
    - **Title**: Max 17 words
    - **Summary**: 400–500 words (Ph.D.) / 350–400 words (Master's)
    - **Paragraphing**: Block style (0 pt indent)
    """)

    # Main File Upload Section
    uploaded_file = st.file_uploader(
        "Upload your manuscript (.docx file)",
        type=["docx"],
        help="Select your Word document manuscript for automatic compliance checking."
    )

    if uploaded_file is not None:
        st.success(f"📄 Loaded file: **{uploaded_file.name}** ({uploaded_file.size // 1024} KB)")
        
        if st.button("🚀 Run Compliance Audit", type="primary"):
            with st.spinner("Analyzing manuscript typography, margins, structure, and citations..."):
                try:
                    auditor = FUOYEManuscriptAuditor(
                        uploaded_file,
                        seminar_type=seminar_code,
                        degree=degree_code
                    )
                    results = auditor.audit_all()
                    
                    st.divider()
                    st.header("📊 Compliance Scorecard & Audit Results")
                    
                    num_pass = len(results["passed"])
                    num_warn = len(results["warnings"])
                    num_fail = len(results["critical"])
                    
                    # Official Compliance Score: Passed Checks vs Critical Violations only
                    total_eval = num_pass + num_fail
                    score = int((num_pass / total_eval) * 100) if total_eval > 0 else 100

                    # Metrics Cards
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Compliance Score", f"{score}%")
                    col2.metric("Passed Checks", num_pass, delta=f"{num_pass} passed", delta_color="normal")
                    col3.metric("Warnings & Advisories", num_warn, delta=f"{num_warn} warnings", delta_color="off")
                    col4.metric("Critical Violations", num_fail, delta=f"{num_fail} violations", delta_color="inverse")

                    st.markdown("---")

                    # Detailed Breakdowns
                    if num_fail > 0:
                        st.error(f"❌ **CRITICAL STRUCTURAL VIOLATIONS ({num_fail})** - Must be fixed before submission:")
                        for item in results["critical"]:
                            st.write(f"- 🔴 {item}")

                    if num_warn > 0:
                        st.warning(f"⚠️ **CRITICAL WARNINGS & ADVISORIES ({num_warn})** - Excluded from score calculation:")
                        for item in results["warnings"]:
                            if "Critical Warning" in item:
                                st.write(f"- 🟧 **{item}**")
                            else:
                                st.write(f"- 🟡 {item}")

                    if num_pass > 0:
                        st.success(f"✔ **PASSED CHECKS ({num_pass})**:")
                        for item in results["passed"]:
                            st.write(f"- 🟢 {item}")

                    # Reference Authenticity & Integrity Report Panel
                    st.divider()
                    st.header("🔍 Reference Authenticity & Integrity Audit")
                    ref_auth = results.get("reference_authenticity", {})
                    
                    auth_col1, auth_col2, auth_col3 = st.columns(3)
                    auth_col1.metric("Authenticity Index", f"{ref_auth.get('authenticity_index', 0)}%")
                    auth_col2.metric("In-Text Citations Scanned", ref_auth.get("total_in_text_citations", 0))
                    auth_col3.metric("Reference List Entries", ref_auth.get("total_ref_entries", 0))

                    if ref_auth.get("authenticity_index", 0) >= 85:
                        st.success(f"🟢 **Status: {ref_auth.get('status')}** — All or most citations correspond to entries in the Reference List.")
                    elif ref_auth.get("authenticity_index", 0) >= 60:
                        st.warning(f"🟡 **Status: {ref_auth.get('status')}** — Some citations were not found in the Reference List.")
                    else:
                        st.error(f"🔴 **Status: {ref_auth.get('status')}** — Significant discrepancy between citations and Reference List.")

                    if ref_auth.get("unmatched_citations"):
                        st.write("⚠️ **Unmatched In-Text Citations (Missing from References):**")
                        for unc in ref_auth.get("unmatched_citations"):
                            st.write(f"  - `{unc}`")

                    # Generate Downloadable Summary Report
                    st.divider()
                    report_content = f"""FUOYE POSTGRADUATE MANUSCRIPT COMPLIANCE AUDIT REPORT
File: {uploaded_file.name}
Degree: {degree_option} | Seminar Stage: {seminar_option}
Official Compliance Score: {score}% (Passed Checks vs Critical Violations)

=================== CRITICAL STRUCTURAL VIOLATIONS ===================
""" + ("\n".join([f"- {x}" for x in results["critical"]]) if results["critical"] else "None!") + f"""

=================== CRITICAL WARNINGS & ADVISORIES ===================
""" + ("\n".join([f"- {x}" for x in results["warnings"]]) if results["warnings"] else "None!") + f"""

=================== REFERENCE AUTHENTICITY AUDIT ===================
Status: {ref_auth.get('status')}
Authenticity Index: {ref_auth.get('authenticity_index')}%
In-Text Citations: {ref_auth.get('total_in_text_citations')} | Reference Entries: {ref_auth.get('total_ref_entries')}
Unmatched Citations: {", ".join(ref_auth.get('unmatched_citations', [])) if ref_auth.get('unmatched_citations') else "None"}

=================== PASSED CHECKS ===================
""" + ("\n".join([f"- {x}" for x in results["passed"]]) if results["passed"] else "None!")

                    st.download_button(
                        label="📥 Download Complete Audit Report (.txt)",
                        data=report_content,
                        file_name=f"Compliance_Audit_{uploaded_file.name}.txt",
                        mime="text/plain"
                    )

                except Exception as e:
                    st.error(f"Error processing manuscript file: {str(e)}")
    else:
        st.info("👆 Please upload a `.docx` manuscript file to begin audit.")

if __name__ == "__main__":
    main()
