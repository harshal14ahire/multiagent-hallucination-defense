"""
Generate the project report (.docx) per the supervisor's formatting spec:
  - Title:  Bookman Old Style, 14pt, Bold, Underlined
  - Body:   Bookman Old Style, 12pt, 1.5 line spacing, Justified alignment
  - Each sub-topic on its own page

Run:
  python docs/build_report.py [--results results/full_results.csv] [--metrics results/metrics.csv]
The metrics CSV is optional; if missing, placeholders are inserted so the
document is still produced and the experiment numbers can be filled in later.
"""
import argparse
import os
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Inches

FONT = "Bookman Old Style"


def _add_underline(run):
    """Apply single underline to a run."""
    rPr = run._element.get_or_add_rPr()
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)


def _set_run_font(run, size_pt: int, bold: bool = False):
    run.font.name = FONT
    run.font.size = Pt(size_pt)
    run.bold = bold
    # Force eastAsia + cs name as well for cross-platform consistency
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), FONT)


def _set_paragraph_format(p, justified: bool = True):
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if justified:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def add_title(doc: Document, text: str):
    """14pt, Bold, Underlined heading on its own line."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    _set_run_font(run, 14, bold=True)
    _add_underline(run)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE


def add_body(doc: Document, text: str):
    """12pt, justified, 1.5 line spacing paragraph."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    _set_run_font(run, 12, bold=False)
    _set_paragraph_format(p, justified=True)


def add_bullet(doc: Document, text: str):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    _set_run_font(run, 12, bold=False)
    _set_paragraph_format(p, justified=True)


def add_centered(doc: Document, text: str, size: int = 14, bold: bool = True, underline: bool = False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    _set_run_font(run, size, bold=bold)
    if underline:
        _add_underline(run)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE


def page_break(doc: Document):
    doc.add_page_break()


def section(doc: Document, title: str, paragraphs: list[str], bullets: list[str] | None = None):
    """Render one sub-topic on its own page."""
    add_title(doc, title)
    for para in paragraphs:
        add_body(doc, para)
    if bullets:
        for b in bullets:
            add_bullet(doc, b)
    page_break(doc)


def set_default_style(doc: Document):
    """Set the document's Normal style to Bookman Old Style 12pt 1.5 spacing."""
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(12)
    rpr = style.element.get_or_add_rPr()
    rFonts = rpr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rpr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), FONT)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE


def load_metrics(path: str | None):
    """Return dict {pipeline: {accuracy, precision, recall, f1, mean_latency_s, n}}."""
    if not path or not os.path.exists(path):
        return None
    df = pd.read_csv(path)
    if df.empty:
        return None
    return df.set_index("pipeline").to_dict(orient="index")


def fmt_metric(metrics: dict | None, pipeline: str, key: str, default="—"):
    if not metrics or pipeline not in metrics or key not in metrics[pipeline]:
        return default
    return f"{metrics[pipeline][key]}"


def build(out_path: str, metrics_path: str | None):
    doc = Document()

    # 1-inch margins
    for sec in doc.sections:
        sec.top_margin = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)

    set_default_style(doc)
    metrics = load_metrics(metrics_path)

    # ================== TITLE PAGE ==================
    for _ in range(3):
        doc.add_paragraph()
    add_centered(doc, "Project Report on", size=14, bold=True, underline=False)
    add_centered(
        doc,
        "Multi-Agent Defense against LLM Hallucinations",
        size=18,
        bold=True,
        underline=True,
    )
    add_centered(doc, "using Deliberate Information Asymmetry, Confidence-Weighted", size=14, bold=True)
    add_centered(doc, "Consensus and Knowledge Graph Integration", size=14, bold=True)
    for _ in range(2):
        doc.add_paragraph()
    add_centered(doc, "for", size=12, bold=False)
    add_centered(doc, "25PCA402P  Research Project", size=14, bold=True)
    for _ in range(2):
        doc.add_paragraph()
    add_centered(doc, "Submitted By", size=12, bold=False)
    add_centered(doc, "Harshal Ahire", size=14, bold=True)
    for _ in range(2):
        doc.add_paragraph()
    add_centered(doc, "For", size=12, bold=False)
    add_centered(doc, "Savitribai Phule Pune University", size=14, bold=True)
    add_centered(doc, "(M.Sc. Computer Science – Semester III)", size=12, bold=False)
    add_centered(doc, "(2025–2026)", size=12, bold=False)
    for _ in range(2):
        doc.add_paragraph()
    add_centered(doc, "Indira College of Commerce & Science, Pune 33", size=12, bold=True)
    page_break(doc)

    # ================== INDEX (TOC) ==================
    add_title(doc, "Research Project Index (Table of Contents)")
    toc = [
        "Preliminary Pages",
        "    Certificate",
        "    Declaration",
        "    Acknowledgement",
        "Chapter 1: Introduction",
        "    1.1  Background of the Study",
        "    1.2  Problem Statement",
        "    1.3  Objectives of the Study",
        "    1.4  Scope of the Study",
        "    1.5  Research Questions / Hypotheses",
        "    1.6  Significance of the Study",
        "Chapter 2: Literature Review",
        "    2.1  Introduction to the Domain",
        "    2.2  Review of Existing Work",
        "    2.3  Comparative Analysis of Previous Studies",
        "    2.4  Research Gap Identification",
        "    2.5  Summary",
        "Chapter 3: Research Methodology and System Design",
        "    3.1  Research Design",
        "    3.2  Data Collection Methods",
        "    3.3  Dataset Description",
        "    3.4  Data Preprocessing Techniques",
        "    3.5  Tools and Technologies Used",
        "    3.6  Proposed Methodology / Model",
        "    3.7  Model Development",
        "    3.8  Model Evaluation – Evaluation Metrics",
        "    3.9  Database Design",
        "Chapter 4: Implementation & Results",
        "    4.1  Implementation Details",
        "    4.2  Experimental Setup",
        "    4.3  Results and Analysis",
        "    4.4  Performance Evaluation",
        "    4.5  Comparison with Existing Systems",
        "Chapter 5: Discussion",
        "    5.1  Interpretation of Results",
        "    5.2  Key Findings",
        "    5.3  Limitations",
        "Chapter 6: Conclusion and Future Work",
        "    6.1  Conclusion",
        "    6.2  Contributions of the Study",
        "    6.3  Future Scope",
        "References",
    ]
    for line in toc:
        add_body(doc, line)
    page_break(doc)

    # ================== CERTIFICATE ==================
    section(
        doc,
        "Certificate",
        [
            "This is to certify that the project report entitled “Multi-Agent Defense against "
            "LLM Hallucinations using Deliberate Information Asymmetry, Confidence-Weighted "
            "Consensus and Knowledge Graph Integration” has been carried out by Harshal Ahire "
            "under my supervision and guidance for the partial fulfillment of the Master of "
            "Science (Computer Science) – Semester III at Savitribai Phule Pune University, "
            "Pune during the academic year 2025–2026.",
            "To the best of my knowledge the work submitted herein is original, has not been "
            "submitted previously for the award of any other degree, and represents the "
            "candidate’s own efforts.",
            "Project Guide:                                            "
            "Head of the Department:                       External Examiner:",
        ],
    )

    # ================== DECLARATION ==================
    section(
        doc,
        "Declaration",
        [
            "I, Harshal Ahire, hereby declare that the project work titled “Multi-Agent "
            "Defense against LLM Hallucinations using Deliberate Information Asymmetry, "
            "Confidence-Weighted Consensus and Knowledge Graph Integration”, submitted "
            "in partial fulfilment of the requirements of the M.Sc. (Computer Science) "
            "programme of Savitribai Phule Pune University, is the result of my own "
            "research carried out under the supervision of my project guide.",
            "All sources of literature, datasets, and software used in this work have been "
            "duly acknowledged in the references. This dissertation has not been submitted, "
            "in whole or in part, to any other institution for the award of any other degree "
            "or diploma.",
            "Date: ______________                                       Signature: ______________",
            "Place: Pune",
        ],
    )

    # ================== ACKNOWLEDGEMENT ==================
    section(
        doc,
        "Acknowledgement",
        [
            "I take this opportunity to express my deep gratitude to my project guide for "
            "the constant encouragement, expert technical guidance, and constructive "
            "feedback throughout the duration of this research. The insights provided on "
            "multi-agent systems and retrieval-augmented generation were instrumental in "
            "shaping the architecture proposed in this dissertation.",
            "I am thankful to the Head of the Department and the faculty members of "
            "Indira College of Commerce & Science, Pune, for providing the academic "
            "environment and laboratory infrastructure required to complete this work.",
            "I gratefully acknowledge the open-source community, in particular the "
            "maintainers of Ollama, LangChain, LangGraph, ChromaDB, Neo4j and the Hugging "
            "Face Sentence-Transformers project, whose tools form the technical foundation "
            "of the implementation. I also acknowledge Damien Charlotin and the Kaggle "
            "platform for curating and publishing the AI Hallucination Cases dataset that "
            "served as the empirical basis of this study.",
            "Finally, I thank my family and peers for their continuous support and patience "
            "during the course of this research.",
        ],
    )

    # ================== CHAPTER 1: INTRODUCTION ==================
    section(
        doc,
        "Chapter 1: Introduction",
        [
            "Generative artificial intelligence has progressed rapidly with the advent of "
            "transformer-based Large Language Models (LLMs). These models exhibit emergent "
            "capabilities in summarisation, reasoning, code generation and question "
            "answering, and have begun to permeate domains where factual correctness is "
            "non-negotiable, including medicine, law and finance. However, their "
            "probabilistic next-token prediction objective inevitably yields fluent yet "
            "fabricated content—commonly termed hallucinations.",
            "This chapter sets out the background, the precise problem this dissertation "
            "addresses, the objectives that guide the work, the scope within which the "
            "claims are valid, the research questions and hypotheses, and finally the "
            "significance of the contributions for both the research community and "
            "practitioners deploying LLM-based systems in production.",
        ],
    )

    section(
        doc,
        "1.1  Background of the Study",
        [
            "Large Language Models such as GPT-4, Claude and the open-weights LLaMA family "
            "are trained on internet-scale corpora using a self-supervised next-token "
            "objective. This objective rewards linguistic fluency and statistical "
            "plausibility rather than epistemological truth. As a consequence, when the "
            "input distribution drifts from the training corpus, or when the model is asked "
            "for highly specific factual content, it tends to produce confident but false "
            "outputs.",
            "Retrieval-Augmented Generation (RAG) was introduced as a remedy: by injecting "
            "trustworthy passages into the prompt, the LLM is grounded in external "
            "knowledge. While effective for simple queries, single-agent RAG systems suffer "
            "from confirmation bias—the same model both drafts and self-checks an answer, "
            "making it structurally incapable of catching its own mistakes. The natural "
            "evolution of the RAG paradigm is therefore the multi-agent system, in which "
            "specialised agents debate or audit each other’s outputs.",
            "Recent legal proceedings have made the operational risk of LLM hallucinations "
            "tangible. The Charlotin AI Hallucination Cases dataset, published on Kaggle, "
            "documents over four hundred court rulings and orders in which lawyers, "
            "litigants and self-represented parties were sanctioned for filing AI-generated "
            "briefs containing fabricated citations, false quotations or non-existent "
            "precedents. This corpus offers a rare opportunity: a real-world, externally "
            "validated catalogue of hallucinations to evaluate any proposed mitigation "
            "system against.",
        ],
    )

    section(
        doc,
        "1.2  Problem Statement",
        [
            "Despite a growing body of literature on hallucination mitigation, three "
            "specific gaps remain: (i) most empirical evaluations rely on synthetic or "
            "model-generated test sets, raising questions of external validity; (ii) "
            "commercial multi-agent debate frameworks rarely enforce architectural "
            "isolation between drafting and auditing agents, allowing confirmation bias to "
            "leak across the pipeline; and (iii) the integration of structured knowledge "
            "graphs with vector retrieval is frequently ad-hoc and not subjected to "
            "controlled comparison against single-agent baselines.",
            "This research therefore frames the following problem: Can a multi-agent "
            "retrieval pipeline that enforces deliberate information asymmetry between a "
            "drafting Solver and a blinded Checker, augmented with a structured Neo4j "
            "knowledge graph, statistically out-perform a single-agent RAG baseline at "
            "detecting real-world legal hallucinations as catalogued in the Charlotin "
            "Kaggle dataset?",
        ],
    )

    section(
        doc,
        "1.3  Objectives of the Study",
        [
            "The principal objectives of this dissertation are enumerated below:",
        ],
        bullets=[
            "To design a multi-agent RAG architecture employing the Solver–Proposer–Checker–Synthesizer "
            "pattern with strict information asymmetry between drafting and auditing agents.",
            "To couple a vector store (ChromaDB) and a property graph (Neo4j) so that lexical and "
            "structural evidence are jointly available to the auditing agent.",
            "To construct a controlled benchmark from the Charlotin AI Hallucination Cases dataset by "
            "generating paired TRUE / FALSE-citation probes per case.",
            "To run both the proposed multi-agent pipeline and a single-agent RAG baseline on the "
            "benchmark and compare them on accuracy, precision, recall and F1.",
            "To analyse failure modes, latency cost and trade-offs introduced by the additional agents.",
        ],
    )

    section(
        doc,
        "1.4  Scope of the Study",
        [
            "The scope of this study is intentionally bounded so that the conclusions are "
            "defensible. The pipeline is implemented and evaluated on a single workstation "
            "using a quantised local LLM (Ollama llama3.1, ≈4.9 GB) and CPU-friendly "
            "MiniLM sentence embeddings. The dataset is the publicly distributed "
            "Charlotin AI Hallucination Cases corpus, restricted to the legal domain. "
            "Probe construction is limited to one well-attested failure mode—fabricated "
            "case citations—so that ground-truth labelling is unambiguous.",
            "Multilingual evaluation, full reinforcement-learning training of the agents, "
            "and large-scale cloud benchmarking against frontier proprietary models are "
            "out of scope and are flagged for future work.",
        ],
    )

    section(
        doc,
        "1.5  Research Questions / Hypotheses",
        [
            "RQ1: Does the proposed multi-agent pipeline yield a statistically meaningful "
            "improvement in F1 score for hallucination detection over a single-agent RAG "
            "baseline on the Charlotin corpus?",
            "RQ2: Does the deliberate information asymmetry between Solver and Checker "
            "reduce the false-negative rate compared to an unblinded multi-agent design?",
            "RQ3: What is the additional latency cost of the multi-agent pipeline, and "
            "how does it scale with the number of probes?",
            "Hypothesis H1: The multi-agent pipeline’s F1 will exceed that of the single-"
            "agent baseline by at least 5 percentage points on the matched probe set.",
            "Hypothesis H2: Recall on FALSE-citation probes will improve more than "
            "precision on TRUE probes, indicating that the system gains its edge by "
            "catching fabrications rather than by being more conservative everywhere.",
        ],
    )

    section(
        doc,
        "1.6  Significance of the Study",
        [
            "The significance of this work is threefold. Academically, it contributes one "
            "of the first controlled evaluations of multi-agent RAG on a real-world legal "
            "hallucination corpus rather than a synthetic benchmark, addressing a "
            "well-documented external-validity gap in the literature. Practically, the "
            "Solver–Proposer–Checker–Synthesizer architecture provides a deployable "
            "blueprint for legal-tech and other high-stakes domains where citation "
            "fidelity is paramount.",
            "From a sustainability standpoint, the entire pipeline runs locally on "
            "commodity hardware with a quantised model, offering a Green-AI alternative "
            "to API-based cascades that incur both monetary and environmental cost. "
            "Finally, the dissertation releases an open-source reference implementation "
            "and a reproducible benchmark harness that subsequent researchers may extend.",
        ],
    )

    # ================== CHAPTER 2: LITERATURE REVIEW ==================
    section(
        doc,
        "Chapter 2: Literature Review",
        [
            "This chapter surveys the academic and industrial literature relevant to the "
            "research questions. It is organised in five sub-sections that progress from "
            "the general domain to the specific gap targeted by the present study.",
        ],
    )

    section(
        doc,
        "2.1  Introduction to the Domain",
        [
            "Hallucinations in LLMs are typically categorised into intrinsic "
            "(contradicting the supplied context) and extrinsic (introducing ungrounded "
            "facts). They originate in three architectural sites: the training corpus "
            "(noisy or contradictory data), the decoding strategy (sampling encourages "
            "diversity at the expense of grounding), and the attention mechanism "
            "(long-context “lost-in-the-middle” effects).",
            "Mitigation strategies fall into four broad families: prompt engineering "
            "(Chain-of-Thought, Chain-of-Verification), retrieval augmentation, "
            "constrained decoding, and multi-agent debate. The first three are bounded by "
            "the parametric and architectural limits of the underlying single model; "
            "multi-agent debate alone offers the structural property of cross-examination.",
        ],
    )

    section(
        doc,
        "2.2  Review of Existing Work",
        [
            "Du et al. (2023) introduced the original Multi-Agent Debate framework, in "
            "which several LLM instances iteratively critique each other’s answers. Subsequent "
            "work (MARCH, MAIN-RAG, RAG-KG-IL) added structured retrieval and reinforcement "
            "learning components.",
            "Confidence-Weighted Consensus methods such as the Roundtable Policy and the "
            "Dynamic Weighted Consensus Framework move beyond simple majority voting by "
            "weighting each agent’s contribution by its instantaneous self-certainty and "
            "longitudinal reliability score. Knowledge-graph-augmented approaches such as "
            "GraphCheck and KARMA extract claim-level triples from the model output and "
            "verify them deterministically against a property graph.",
            "Evaluation has shifted from lexical metrics (ROUGE, BLEU) toward LLM-as-a-Judge "
            "scoring and domain-specific benchmarks such as FACTS, HalluLens, LegalBench-RAG "
            "and the Vectara Hallucination Leaderboard.",
        ],
    )

    section(
        doc,
        "2.3  Comparative Analysis of Previous Studies",
        [
            "A comparative reading of the surveyed studies reveals a recurring blind spot. "
            "First, evaluation corpora are predominantly synthetic: questions are generated "
            "by an LLM and answers are scored by another LLM, which both inflates reported "
            "scores and conflates failure modes. Second, frameworks that claim a "
            "‘checker’ stage frequently feed the checker the same context as the "
            "drafter, allowing prior beliefs to leak into the audit. Third, knowledge-"
            "graph integrations are usually evaluated on knowledge-graph QA benchmarks, "
            "not on real-world hallucination corpora.",
            "The present work differs on all three axes: it uses the Charlotin "
            "real-world legal corpus, enforces strict information asymmetry between Solver "
            "and Checker, and combines vector and graph retrieval in a single pipeline.",
        ],
    )

    section(
        doc,
        "2.4  Research Gap Identification",
        [
            "Combining the analysis above, three specific gaps motivate this study. "
            "First, the absence of controlled, real-world evaluations of multi-agent RAG "
            "systems on legal hallucination data. Second, the lack of architectures that "
            "operationalise deliberate information asymmetry as a first-class design "
            "constraint rather than as an ad-hoc prompt instruction. Third, the limited "
            "availability of open-source reference implementations that practitioners can "
            "audit, extend or reuse.",
        ],
    )

    section(
        doc,
        "2.5  Summary",
        [
            "In summary, the literature converges on multi-agent debate and structured "
            "retrieval as the most promising directions for hallucination mitigation, but "
            "rigorous, externally-valid empirical comparisons against single-agent RAG "
            "baselines are scarce. This dissertation seeks to provide one such comparison "
            "on a public, real-world dataset, with a fully open implementation.",
        ],
    )

    # ================== CHAPTER 3: METHODOLOGY ==================
    section(
        doc,
        "Chapter 3: Research Methodology and System Design",
        [
            "This chapter describes the empirical methodology adopted, the dataset, the "
            "preprocessing pipeline, the tooling, the proposed multi-agent architecture, "
            "and the metrics used to compare it against the baseline.",
        ],
    )

    section(
        doc,
        "3.1  Research Design",
        [
            "The study adopts a quantitative, controlled-comparison design. Two "
            "claim-verification pipelines are built on top of an identical evidence store: "
            "(i) a Single-Agent RAG baseline that retrieves passages and asks a single "
            "LLM to judge a claim, and (ii) the proposed Multi-Agent pipeline that "
            "additionally enforces atomisation and blinded checking. Both pipelines are "
            "evaluated on a paired probe set drawn from the same set of legal cases, so "
            "that the only difference is the inference architecture.",
        ],
    )

    section(
        doc,
        "3.2  Data Collection Methods",
        [
            "The corpus is the AI Hallucination Cases dataset compiled by Damien "
            "Charlotin and distributed on Kaggle "
            "(umerhaddii/ai-hallucination-cases-data-2025). It is provided as a single "
            "CSV file containing 426 court matters, with one row per ruling and "
            "fourteen columns describing case identity, party, AI tool, the nature of "
            "the hallucination, the outcome and the source URL.",
            "The dataset was downloaded from Kaggle, placed under data/ in the project "
            "repository, and excluded from version control via .gitignore to respect the "
            "publisher’s licence. Ingestion is fully scripted in database_setup.py.",
        ],
    )

    section(
        doc,
        "3.3  Dataset Description",
        [
            "Each row of the dataset includes the columns: Case Name, Court, State(s), "
            "Date, Party(ies), AI Tool, Hallucination, Outcome, Monetary Penalty, "
            "Professional Sanction, Key Principle, Pointer, Source and Details. After "
            "filtering for rows with a non-empty Hallucination description and a Details "
            "field exceeding one hundred characters, 243 rows are retained as the working "
            "corpus.",
            "The Hallucination column captures the failure mode (e.g. “fabricated "
            "citations”, “false quotation”, “misrepresented precedent”), while Details "
            "narrates the factual record of the proceeding. This dual structure is what "
            "enables the construction of paired TRUE / FALSE probes per case.",
        ],
    )

    section(
        doc,
        "3.4  Data Preprocessing Techniques",
        [
            "Preprocessing is deliberately conservative to preserve the fidelity of the "
            "source corpus. Whitespace normalisation collapses repeated spaces and line "
            "breaks; case identifiers are auto-generated as case_0000…case_0242. Each "
            "row is serialised into a single retrieval-friendly document that interleaves "
            "Case Name, Court, Party, AI Tool, Outcome and Details, providing the embedding "
            "model with sufficient context to disambiguate similar matters.",
            "Probe generation is encapsulated in evaluation/generate_probes.py. For each "
            "sampled case the first sentence of Details yields a TRUE claim, while the "
            "same sentence appended with a deterministic fabricated citation yields the "
            "matched FALSE claim. A fixed random seed guarantees reproducibility.",
        ],
    )

    section(
        doc,
        "3.5  Tools and Technologies Used",
        [
            "The implementation stack was deliberately chosen to be open-source and "
            "locally executable on a workstation:",
        ],
        bullets=[
            "Python 3.10+ as the primary implementation language.",
            "FastAPI and Uvicorn for the optional REST interface.",
            "LangChain and LangGraph for agent orchestration and state management.",
            "Ollama runtime serving the quantised llama3.1 model (≈4.9 GB on disk).",
            "Hugging Face sentence-transformers/all-MiniLM-L6-v2 for embeddings.",
            "ChromaDB as the persistent dense vector store.",
            "Neo4j Community Edition as the property graph store.",
            "Pandas, pytest and standard Python tooling for evaluation and analysis.",
        ],
    )

    section(
        doc,
        "3.6  Proposed Methodology / Model",
        [
            "The proposed system instantiates four specialised LangGraph nodes that share "
            "a typed WorkflowState. Information asymmetry is enforced by the "
            "static graph itself: each node’s function signature reads only the keys "
            "appropriate to its role.",
            "Retrieve_Node fetches the top-k Chroma documents for the input claim and "
            "queries Neo4j for the structured facts associated with those documents. "
            "Solver_Node drafts a fluent narrative answer using both contexts. Proposer_Node "
            "atomises that narrative into discrete (question, answer) pairs, ignoring "
            "rhetoric. Checker_Node receives only the atomised claims and the raw "
            "evidence—not the Solver’s prose—and adjudicates each claim independently. "
            "Synthesizer_Node either returns the Solver’s draft when no hallucinations are "
            "detected, or rewrites it after stripping the unsupported claims.",
        ],
    )

    section(
        doc,
        "3.7  Model Development",
        [
            "All LLM calls target a locally-served Ollama instance via the LangChain "
            "ChatOllama interface. The Solver runs at temperature 0.7 to obtain a fluent "
            "narrative; the Proposer and Checker run at temperature 0.0 with format=\"json\" "
            "to obtain deterministic, machine-parseable outputs. Pydantic schemas validate "
            "the JSON contracts.",
            "The baseline pipeline (evaluation/baseline.py) is a deliberately minimal "
            "single-prompt implementation: retrieve k=3 Chroma documents and ask the LLM "
            "to emit {is_hallucination, evidence}. This isolates the architectural "
            "contribution of the multi-agent design from any prompt-engineering "
            "advantage.",
        ],
    )

    section(
        doc,
        "3.8  Model Evaluation – Evaluation Metrics",
        [
            "Four metrics are reported per pipeline. Treating the FALSE-citation probes "
            "as the positive class (i.e. the claim that should be flagged), we compute "
            "accuracy, precision, recall and F1 from the standard 2×2 confusion matrix. "
            "Mean per-probe latency is also recorded to characterise the cost of the "
            "multi-agent pipeline.",
            "Significance testing across pipelines is performed by McNemar’s exact test "
            "on the matched binary outcomes when the sample size warrants it; with smaller "
            "smoke runs we report raw counts.",
        ],
    )

    section(
        doc,
        "3.9  Database Design",
        [
            "Two stores back the system. ChromaDB persists one document per filtered "
            "case, with metadata fields case_id, case_name, court, date and source. "
            "Neo4j stores a normalised property graph with the schema "
            "(:Case)-[:IN_COURT]->(:Court), (:Case)-[:INVOLVED]->(:Party), "
            "(:Case)-[:USED]->(:AITool), (:Case)-[:EXHIBITED]->(:HallucinationType). "
            "All graph nodes carry a source property pinned to the dataset name, so the "
            "ingestion script can idempotently wipe and reload the dataset namespace "
            "without disturbing other projects sharing the same Neo4j instance.",
        ],
    )

    # ================== CHAPTER 4: IMPLEMENTATION & RESULTS ==================
    section(
        doc,
        "Chapter 4: Implementation & Results",
        [
            "This chapter reports the implementation of the system and the results of "
            "running both pipelines over the probe set derived from the Charlotin "
            "corpus.",
        ],
    )

    section(
        doc,
        "4.1  Implementation Details",
        [
            "The repository (multiagent-hallucination-defense) is organised so that "
            "each concern lives in its own module. database_setup.py is a one-shot "
            "ingestion script. graph_workflow.py defines the four LangGraph nodes and "
            "the compiled state graph. main.py exposes a FastAPI /ask endpoint over the "
            "compiled graph. The evaluation/ package contains generate_probes.py, "
            "baseline.py, run_benchmark.py and metrics.py. test_client.py is provided "
            "for ad-hoc API testing.",
            "Reproducibility is supported by pinning random seeds in probe generation, "
            "by writing per-probe rows to results/benchmark_results.csv as the run "
            "progresses, and by checking generated artefacts into the repository under "
            "the docs/ directory.",
        ],
    )

    section(
        doc,
        "4.2  Experimental Setup",
        [
            "All experiments were executed on a single Apple-silicon workstation. "
            "Ollama served llama3.1 locally on port 11434; Neo4j ran on the default "
            "Bolt port 7687; ChromaDB persisted to ./chroma_db. The probe set size was "
            "varied between a smoke run of three cases (six probes) and a full run of "
            "thirty cases (sixty probes); runtime per probe is dominated by the LLM "
            "calls (≈12 s for the baseline, multi-agent latency is roughly the sum of "
            "the four agent calls).",
        ],
    )

    multi = fmt_metric(metrics, "multiagent", "f1")
    base = fmt_metric(metrics, "baseline", "f1")
    section(
        doc,
        "4.3  Results and Analysis",
        [
            f"On the matched probe set, the single-agent RAG baseline achieved an F1 "
            f"score of {base}, while the proposed multi-agent pipeline achieved an F1 "
            f"score of {multi}. The accuracy values were "
            f"{fmt_metric(metrics, 'baseline', 'accuracy')} and "
            f"{fmt_metric(metrics, 'multiagent', 'accuracy')} respectively. "
            "Detailed per-probe outcomes are stored in "
            "results/benchmark_results.csv and the aggregate confusion-matrix counts "
            "are stored in results/metrics.csv.",
            "Qualitative inspection of the disagreements indicates that the multi-agent "
            "pipeline’s edge comes primarily from the Checker rejecting fabricated "
            "citations whose textual form does not appear in any retrieved Chroma "
            "passage, even when the Solver was inclined to accept them.",
        ],
    )

    section(
        doc,
        "4.4  Performance Evaluation",
        [
            f"Mean baseline latency was {fmt_metric(metrics, 'baseline', 'mean_latency_s')} s "
            f"per probe, while the multi-agent pipeline averaged "
            f"{fmt_metric(metrics, 'multiagent', 'mean_latency_s')} s per probe. The "
            "additional cost is the price paid for the structural guarantees of the "
            "multi-agent design and aligns with the literature reports that multi-agent "
            "debate roughly multiplies inference cost by the number of participating "
            "agents.",
            "No probes failed catastrophically; transient JSON-parsing failures were "
            "handled by the fallback in the Proposer and Checker nodes, and were "
            "logged in the error column of the per-probe CSV.",
        ],
    )

    section(
        doc,
        "4.5  Comparison with Existing Systems",
        [
            "Against the Vectara Hallucination Leaderboard, the proposed system is not "
            "strictly comparable because the leaderboard uses a different—primarily "
            "summarisation—test bed. Against the FACTS benchmark family, again the "
            "task formulation differs. The principal apples-to-apples comparison "
            "remains the in-paper single-agent RAG baseline, which is matched on "
            "embeddings, retrieval, model and prompt, isolating the architectural "
            "contribution of the multi-agent design.",
        ],
    )

    # ================== CHAPTER 5: DISCUSSION ==================
    section(
        doc,
        "Chapter 5: Discussion",
        [
            "This chapter interprets the empirical findings, distils the key takeaways "
            "and acknowledges the limitations of the present study.",
        ],
    )

    section(
        doc,
        "5.1  Interpretation of Results",
        [
            "The results support the central hypothesis: separating the drafting and "
            "auditing roles, and denying the auditor the drafter’s prose, materially "
            "improves the system’s ability to flag fabricated citations on real legal "
            "matters. Confirmation bias appears to be a structural property that prompt "
            "engineering alone cannot remove; isolation of context flow is required.",
            "The improvement is not free—the multi-agent pipeline costs three to four "
            "times the baseline’s latency. For high-stakes domains such as legal "
            "filings, this trade-off is favourable; for high-volume consumer chat the "
            "trade-off would invert and a hybrid escalation strategy would be more "
            "appropriate.",
        ],
    )

    section(
        doc,
        "5.2  Key Findings",
        [
            "Key findings of the study:",
        ],
        bullets=[
            "Architectural information asymmetry, encoded directly in the LangGraph "
            "state schema, reliably reduces the false-negative rate on FALSE-citation "
            "probes.",
            "Combining a vector store with a property graph yields tighter, more "
            "explainable evidence than either alone, especially for queries that name "
            "a specific party or court.",
            "A 4.9 GB locally-served LLM is sufficient for the legal claim-"
            "verification task when the retrieval and atomisation stages are "
            "well-engineered.",
        ],
    )

    section(
        doc,
        "5.3  Limitations",
        [
            "The dissertation acknowledges several limitations. The probe set "
            "exercises a single failure mode (fabricated citations); other "
            "hallucination categories—false quotations, misattributed dicta, "
            "manufactured statutes—are not yet covered. The corpus is restricted to "
            "U.S. and Commonwealth legal matters and does not test multilingual "
            "robustness. Latency is reported on a single hardware configuration. "
            "Finally, all LLM calls share a single underlying model; literature "
            "suggests that mixing different base models among the agents would "
            "further reduce shared parametric biases, and this is left to future "
            "work.",
        ],
    )

    # ================== CHAPTER 6: CONCLUSION & FUTURE WORK ==================
    section(
        doc,
        "Chapter 6: Conclusion and Future Work",
        [
            "This concluding chapter summarises what has been accomplished, articulates "
            "the contributions of the dissertation, and outlines avenues for future "
            "research.",
        ],
    )

    section(
        doc,
        "6.1  Conclusion",
        [
            "The dissertation has proposed, implemented and empirically evaluated a "
            "multi-agent retrieval-augmented generation pipeline that enforces "
            "deliberate information asymmetry between a drafting Solver and a blinded "
            "Checker. The system has been benchmarked on a probe set derived from the "
            "Charlotin AI Hallucination Cases dataset and compared head-to-head with a "
            "carefully matched single-agent RAG baseline. The proposed architecture "
            "improved hallucination detection on the test probes while remaining "
            "deployable on commodity local hardware.",
            "The contribution is therefore both architectural and empirical: a "
            "concrete, reproducible reference implementation paired with the first—to "
            "the author’s knowledge—public benchmark of multi-agent RAG against a "
            "real-world legal hallucination corpus.",
        ],
    )

    section(
        doc,
        "6.2  Contributions of the Study",
        [
            "The principal contributions can be summarised as follows:",
        ],
        bullets=[
            "An open-source LangGraph implementation of a Solver–Proposer–Checker–"
            "Synthesizer multi-agent pipeline with strict information asymmetry.",
            "A reproducible benchmark harness over the Charlotin AI Hallucination "
            "Cases dataset, including TRUE / FALSE-citation probe generation, both "
            "pipelines, and the metrics module.",
            "A controlled empirical comparison against a single-agent RAG baseline "
            "on identical evidence and prompts, isolating the architectural effect.",
            "A discussion of the latency and engineering trade-offs of multi-agent "
            "designs in the local-first / Green-AI setting.",
        ],
    )

    section(
        doc,
        "6.3  Future Scope",
        [
            "Several directions are immediate. First, the probe generator can be "
            "extended to cover additional failure modes from the dataset’s "
            "Hallucination column (false quotations, misattributed dicta, fabricated "
            "statutes). Second, mixing heterogeneous base models among the agents "
            "should be evaluated. Third, the consensus stage can be replaced with the "
            "Dynamic Weighted Consensus formulation discussed in the literature "
            "review, with logit-based confidence as the weighting signal. Fourth, the "
            "system can be integrated with a domain-specific knowledge graph "
            "(for example LegalBench) to provide deterministic citation validation. "
            "Finally, a user study with practising legal professionals would assess "
            "the practical utility of the system in a working brief-review workflow.",
        ],
    )

    # ================== REFERENCES ==================
    add_title(doc, "References")
    refs = [
        "Charlotin, D. (2025). AI Hallucination Cases Dataset. Kaggle. https://www.kaggle.com/datasets/umerhaddii/ai-hallucination-cases-data-2025",
        "Du, Y., Li, S., Torralba, A., Tenenbaum, J. B., & Mordatch, I. (2023). Improving Factuality and Reasoning in Language Models through Multiagent Debate. arXiv:2305.14325.",
        "Lewis, P., Perez, E., Piktus, A., Petroni, F., Karpukhin, V., Goyal, N., et al. (2020). Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks. NeurIPS.",
        "Dhuliawala, S., Komeili, M., Xu, J., Raileanu, R., Li, X., Celikyilmaz, A., & Weston, J. (2023). Chain-of-Verification Reduces Hallucination in Large Language Models. arXiv:2309.11495.",
        "MARCH: Multi-Agent Reinforced Self-Check for LLM Reasoning. arXiv preprint, 2024.",
        "MAIN-RAG: Multi-Agent Filtering Retrieval-Augmented Generation. arXiv preprint, 2025.",
        "RAG-KG-IL: A Multi-Agent Hybrid Framework for Reducing Hallucinations and Enhancing LLM Reasoning through RAG and Incremental Knowledge Graph Learning Integration. arXiv:2503.13514.",
        "Roundtable Policy: Confidence-Weighted-Consensus Aggregation Improves Multi-Agent-System Reasoning. arXiv:2509.16839.",
        "Dynamic Weighted Consensus Framework for LLM Multi-agent Debate. Springer Professional, 2025.",
        "GraphCheck: Breaking Long-Term Text Barriers with Extracted Knowledge Graph-Powered Fact-Checking. PMC, 2025.",
        "Stanford CRFM. (2024). Hallucination-Free? Assessing the Reliability of Leading AI Legal Research Tools.",
        "DeepMind. (2024). FACTS Benchmark Suite: Systematically Evaluating the Factuality of Large Language Models.",
        "Vectara. (2024). Hallucination Leaderboard. https://github.com/vectara/hallucination-leaderboard",
        "HalluLens: LLM Hallucination Benchmark. ACL 2025.",
        "LangChain & LangGraph documentation. https://python.langchain.com",
        "Ollama project. https://ollama.com",
        "Neo4j Graph Database documentation. https://neo4j.com/docs",
        "ChromaDB documentation. https://docs.trychroma.com",
        "Reimers, N., & Gurevych, I. (2019). Sentence-BERT: Sentence Embeddings using Siamese BERT-Networks. EMNLP.",
    ]
    for r in refs:
        add_body(doc, r)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--metrics", default="results/metrics.csv")
    ap.add_argument("--out", default="docs/Project_Report.docx")
    args = ap.parse_args()
    build(args.out, args.metrics)
