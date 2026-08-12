from pathlib import Path
import re

import pandas as pd


# ============================================================
# FILE PATHS
# ============================================================

PROJECT_FOLDER = (
    Path.home()
    / "Desktop"
    / "POLITICAL_CANDIDATE_ANALYTICS"
)

INPUT_FILE = (
    PROJECT_FOLDER
    / "data"
    / "raw"
    / "political_candidate_database_pilot_expanded.xlsx"
)

OUTPUT_EXCEL = (
    PROJECT_FOLDER
    / "data"
    / "processed"
    / "policy_classification_results_expanded.xlsx"
)

OUTPUT_CSV = (
    PROJECT_FOLDER
    / "data"
    / "processed"
    / "policy_classification_results_expanded.csv"
)


# ============================================================
# POLICY KEYWORD DICTIONARY
# ============================================================

POLICY_KEYWORDS = {
    "Economy": [
        "economy",
        "economic",
        "economic development",
        "economic opportunity",
        "cost of living",
        "affordability",
        "affordable",
        "inflation",
        "small business",
        "small businesses",
        "entrepreneur",
        "entrepreneurship",
        "business growth",
        "job growth",
        "workforce development",
        "workforce-development",
        "household costs",
        "price gouging",
        "junk fees",
    ],

    "Education": [
        "education",
        "public education",
        "public school",
        "public schools",
        "school funding",
        "schools",
        "students",
        "student",
        "teachers",
        "teacher",
        "teacher pay",
        "teacher retention",
        "teacher recruitment",
        "literacy",
        "class size",
        "class sizes",
        "higher education",
        "college",
        "university",
        "career and technical education",
        "technical education",
        "special education",
        "special-education",
        "multilingual education",
        "multilingual-education",
        "early childhood education",
        "early-childhood education",
    ],

    "Healthcare": [
        "healthcare",
        "health care",
        "medical care",
        "health coverage",
        "healthcare coverage",
        "insurance",
        "medicaid",
        "medicare",
        "minnesotacare",
        "prescription",
        "prescription drugs",
        "medical costs",
        "community clinic",
        "community clinics",
        "local clinics",
        "preventive care",
        "rural hospital",
        "rural hospitals",
        "patient",
        "patients",
    ],

    "Mental Health": [
        "mental health",
        "mental-health",
        "behavioral health",
        "crisis response",
        "crisis-response",
        "trauma-informed",
        "trauma informed",
        "addiction",
        "substance use",
        "substance-use",
        "counselor",
        "counselors",
        "therapy",
        "dementia",
        "alzheimer",
        "alzheimer's",
        "caregiver",
        "caregivers",
    ],

    "Housing": [
        "housing",
        "affordable housing",
        "homeownership",
        "home ownership",
        "rent",
        "renter",
        "renters",
        "renter protections",
        "tenant",
        "tenants",
        "tenant protections",
        "landlord",
        "landlords",
        "homeless",
        "homelessness",
        "unhoused",
        "eviction",
        "property tax",
        "property taxes",
    ],

    "Public Safety": [
        "public safety",
        "community safety",
        "gun violence",
        "gun-violence",
        "gun safety",
        "firearm",
        "firearms",
        "safe firearm storage",
        "assault weapon",
        "assault weapons",
        "background checks",
        "first responders",
        "first responder",
        "emergency response",
        "police",
        "law enforcement",
        "crime prevention",
        "victim services",
    ],

    "Criminal Justice": [
        "criminal justice",
        "justice reform",
        "justice system",
        "incarceration",
        "reentry",
        "re-entry",
        "rehabilitation",
        "court",
        "courts",
        "prison",
        "prisons",
        "sentencing",
        "formerly incarcerated",
        "second chance",
        "second chances",
        "equal justice",
    ],

    "Environment": [
        "environment",
        "environmental",
        "environmental protection",
        "climate",
        "climate change",
        "clean water",
        "water protection",
        "boundary waters",
        "pollution",
        "polluter",
        "polluters",
        "natural resources",
        "carbon emissions",
        "sustainability",
        "sustainable",
    ],

    "Energy": [
        "renewable energy",
        "clean energy",
        "solar",
        "wind energy",
        "energy efficiency",
        "electricity",
        "utility",
        "utilities",
        "carbon-free",
        "carbon free",
        "data centers",
        "data center",
        "energy-intensive data centers",
    ],

    "Transportation": [
        "transportation",
        "public transportation",
        "public transit",
        "transit",
        "roads",
        "highways",
        "infrastructure",
        "pedestrian",
        "bike",
        "bicycle",
        "rideshare",
        "uber",
        "lyft",
    ],

    "Immigration": [
        "immigration",
        "immigration status",
        "immigrant",
        "immigrants",
        "immigrant rights",
        "immigrant families",
        "refugee",
        "refugees",
        "undocumented",
        "asylum",
        "deportation",
        "ice enforcement",
        "due process",
        "new americans",
    ],

    "Civil Rights": [
        "civil rights",
        "civil liberties",
        "human rights",
        "equal rights",
        "equality",
        "equity",
        "discrimination",
        "racial justice",
        "racial equity",
        "voting rights",
        "disability rights",
        "disability justice",
        "equal opportunity",
        "inclusion",
        "fundamental rights",
    ],

    "LGBTQ+ Rights": [
        "lgbtq",
        "lgbtq+",
        "lgbtq+ rights",
        "transgender",
        "trans rights",
        "trans refuge",
        "marriage equality",
        "gender identity",
        "sexual orientation",
        "queer",
    ],

    "Reproductive Rights": [
        "reproductive freedom",
        "reproductive rights",
        "reproductive healthcare",
        "reproductive health",
        "abortion",
        "fertility treatment",
        "fertility care",
        "ivf",
        "contraception",
        "pregnancy",
        "maternal health",
    ],

    "Labor and Workers": [
        "labor",
        "workers",
        "worker",
        "workers' rights",
        "worker protections",
        "union",
        "unions",
        "collective bargaining",
        "living wage",
        "minimum wage",
        "fair wages",
        "fair salaries",
        "working conditions",
        "paid leave",
        "employee",
        "employees",
    ],

    "Agriculture": [
        "agriculture",
        "agricultural",
        "farmer",
        "farmers",
        "family farm",
        "family farms",
        "farming",
        "crops",
        "livestock",
        "rural economy",
        "sustainable agriculture",
        "food production",
    ],

    "Technology and AI": [
        "artificial intelligence",
        "artificial-intelligence",
        "ai regulation",
        "ai-generated",
        "deepfake",
        "deepfakes",
        "election deepfakes",
        "technology regulation",
        "big tech",
        "data privacy",
        "digital privacy",
        "algorithm",
        "algorithms",
        "social media",
        "chatbot",
        "chatbots",
        "nonconsensual sexual images",
    ],

    "Government and Democracy": [
        "democracy",
        "democratic accountability",
        "election integrity",
        "voting",
        "vote",
        "voters",
        "voting rights",
        "campaign finance",
        "government accountability",
        "transparency",
        "transparent",
        "gerrymandering",
        "public trust",
        "term limits",
        "fraud prevention",
        "corporate fraud",
        "corporate-fraud",
        "auditing",
        "accountability",
    ],

    "Veterans": [
        "veteran",
        "veterans",
        "veterans' services",
        "veterans services",
        "military families",
        "military family",
        "armed forces",
        "national guard",
        "service members",
    ],

    "Childcare and Families": [
        "childcare",
        "child care",
        "working families",
        "family support",
        "families",
        "children",
        "parents",
        "parent",
        "paid family leave",
        "head start",
        "youth",
        "young people",
        "childhood hunger",
        "older adults",
    ],

    "Taxes and Budget": [
        "tax",
        "taxes",
        "taxpayer",
        "taxpayers",
        "property tax",
        "property taxes",
        "state budget",
        "budget",
        "public finance",
        "government spending",
        "state spending",
        "fiscal responsibility",
        "bonding",
        "pensions",
    ],
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):
    """
    Normalize text for policy keyword matching.
    """
    if pd.isna(value):
        return ""

    text = str(value).lower()

    text = text.replace("’", "'")
    text = text.replace("‘", "'")
    text = text.replace("“", '"')
    text = text.replace("”", '"')
    text = text.replace("–", "-")
    text = text.replace("—", "-")

    text = re.sub(
        r"[^a-z0-9\s+\-']",
        " ",
        text,
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    ).strip()

    return text


def phrase_pattern(phrase):
    """
    Convert a keyword or phrase into a flexible regex pattern.

    Spaces and hyphens are treated as interchangeable so phrases
    such as 'mental health' can match 'mental-health'.
    """
    cleaned_phrase = clean_text(phrase)

    phrase_parts = re.split(
        r"[\s\-]+",
        cleaned_phrase,
    )

    phrase_parts = [
        part
        for part in phrase_parts
        if part
    ]

    escaped_parts = [
        re.escape(part)
        for part in phrase_parts
    ]

    flexible_phrase = (
        r"[\s\-]+".join(escaped_parts)
    )

    return (
        rf"(?<!\w)"
        rf"{flexible_phrase}"
        rf"(?!\w)"
    )


def find_keyword_matches(text, keywords):
    """
    Return the unique keywords found in a statement.
    """
    matches = []

    for keyword in keywords:
        pattern = phrase_pattern(keyword)

        if re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        ):
            matches.append(keyword)

    return sorted(
        set(matches)
    )


def calculate_confidence(match_count):
    """
    Convert keyword match count into a confidence category.
    """
    if match_count >= 4:
        return "High"

    if match_count >= 2:
        return "Medium"

    if match_count == 1:
        return "Low"

    return "Unknown"


def calculate_score(
    matches,
    statement_text,
):
    """
    Calculate a transparent rule-based score.

    Longer keyword phrases receive more weight because they are
    generally more policy-specific than single words.
    """
    weighted_score = 0.0

    for keyword in matches:
        word_count = len(
            re.split(
                r"[\s\-]+",
                keyword.strip(),
            )
        )

        if word_count >= 3:
            weighted_score += 3.0

        elif word_count == 2:
            weighted_score += 2.0

        else:
            weighted_score += 1.0

    statement_word_count = max(
        len(statement_text.split()),
        1,
    )

    normalized_score = (
        weighted_score
        / statement_word_count
        * 100
    )

    return round(
        normalized_score,
        2,
    )


def assign_stance(statement_type):
    """
    Assign a cautious preliminary stance.

    Campaign goals and priority statements are marked as Priority.
    Other statement types remain Unclear unless manually reviewed.
    """
    normalized_type = clean_text(
        statement_type
    )

    priority_statement_types = {
        "why running",
        "goals if elected",
        "achievements if elected",
        "areas of concentration",
        "general philosophy",
    }

    if normalized_type in priority_statement_types:
        return "Priority"

    return "Unclear"


def confidence_to_rank(value):
    """
    Convert confidence labels into sortable numeric ranks.
    """
    rank_map = {
        "Unknown": 0,
        "Low": 1,
        "Medium": 2,
        "High": 3,
    }

    return rank_map.get(
        value,
        0,
    )


def rank_to_confidence(value):
    """
    Convert a confidence rank back into a text label.
    """
    label_map = {
        0: "Unknown",
        1: "Low",
        2: "Medium",
        3: "High",
    }

    return label_map.get(
        int(value),
        "Unknown",
    )


# ============================================================
# LOAD INPUT DATA
# ============================================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Input workbook was not found:\n{INPUT_FILE}"
    )

excel_file = pd.ExcelFile(
    INPUT_FILE
)

required_sheets = [
    "Candidates",
    "CandidateStatements",
]

missing_sheets = [
    sheet_name
    for sheet_name in required_sheets
    if sheet_name not in excel_file.sheet_names
]

if missing_sheets:
    raise ValueError(
        "The following required worksheets are missing: "
        + ", ".join(missing_sheets)
    )

candidates = pd.read_excel(
    INPUT_FILE,
    sheet_name="Candidates",
    dtype=object,
).dropna(
    how="all"
).reset_index(
    drop=True
)

statements = pd.read_excel(
    INPUT_FILE,
    sheet_name="CandidateStatements",
    dtype=object,
).dropna(
    how="all"
).reset_index(
    drop=True
)

if statements.empty:
    raise ValueError(
        "CandidateStatements does not contain any records."
    )

required_statement_columns = [
    "Statement_ID",
    "Candidate_ID",
    "Statement_Type",
    "Statement_Text",
]

missing_columns = [
    column
    for column in required_statement_columns
    if column not in statements.columns
]

if missing_columns:
    raise ValueError(
        "CandidateStatements is missing required columns: "
        + ", ".join(missing_columns)
    )


# ============================================================
# REMOVE ROWS WITHOUT A REAL STATEMENT ID
# ============================================================

statements["Statement_ID"] = (
    statements["Statement_ID"]
    .fillna("")
    .astype(str)
    .str.strip()
)

statements = statements[
    statements["Statement_ID"] != ""
].copy()

statements.reset_index(
    drop=True,
    inplace=True,
)


# ============================================================
# BUILD CANDIDATE NAME LOOKUP
# ============================================================

candidate_name_lookup = {}

if (
    "Candidate_ID" in candidates.columns
    and "Ballot_Name" in candidates.columns
):
    for _, row in candidates.iterrows():
        candidate_id = str(
            row.get(
                "Candidate_ID",
                "",
            )
        ).strip()

        ballot_name = str(
            row.get(
                "Ballot_Name",
                "",
            )
        ).strip()

        if candidate_id:
            candidate_name_lookup[
                candidate_id
            ] = ballot_name


# ============================================================
# CLASSIFY POLICY TOPICS
# ============================================================

classification_rows = []
unclassified_rows = []

policy_record_number = 1

for _, statement in statements.iterrows():
    statement_id = str(
        statement.get(
            "Statement_ID",
            "",
        )
    ).strip()

    candidate_id = str(
        statement.get(
            "Candidate_ID",
            "",
        )
    ).strip()

    statement_type = str(
        statement.get(
            "Statement_Type",
            "",
        )
    ).strip()

    statement_text = str(
        statement.get(
            "Statement_Text",
            "",
        )
    ).strip()

    cleaned_statement = clean_text(
        statement_text
    )

    matched_policy_topics = []

    for (
        policy_topic,
        keywords,
    ) in POLICY_KEYWORDS.items():

        matched_keywords = (
            find_keyword_matches(
                cleaned_statement,
                keywords,
            )
        )

        if not matched_keywords:
            continue

        match_count = len(
            matched_keywords
        )

        classification_score = (
            calculate_score(
                matched_keywords,
                cleaned_statement,
            )
        )

        classification_confidence = (
            calculate_confidence(
                match_count
            )
        )

        policy_record_id = (
            f"POL{policy_record_number:06d}"
        )

        classification_rows.append(
            {
                "Policy_Record_ID": (
                    policy_record_id
                ),
                "Candidate_ID": candidate_id,
                "Ballot_Name": (
                    candidate_name_lookup.get(
                        candidate_id,
                        "",
                    )
                ),
                "Statement_ID": statement_id,
                "Statement_Type": statement_type,
                "Policy_Topic": policy_topic,
                "Policy_Subtopic": "",
                "Position_Summary": statement_text,
                "Stance": assign_stance(
                    statement_type
                ),
                "Classification_Method": (
                    "Rule-Based NLP"
                ),
                "Matched_Keywords": "; ".join(
                    matched_keywords
                ),
                "Keyword_Match_Count": (
                    match_count
                ),
                "Classification_Score": (
                    classification_score
                ),
                "Classification_Confidence": (
                    classification_confidence
                ),
                "Human_Reviewed": "No",
                "Reviewer_Notes": "",
            }
        )

        matched_policy_topics.append(
            policy_topic
        )

        policy_record_number += 1

    if not matched_policy_topics:
        unclassified_rows.append(
            {
                "Statement_ID": statement_id,
                "Candidate_ID": candidate_id,
                "Ballot_Name": (
                    candidate_name_lookup.get(
                        candidate_id,
                        "",
                    )
                ),
                "Statement_Type": statement_type,
                "Statement_Text": statement_text,
                "Reason": (
                    "No policy keywords matched "
                    "the current keyword dictionary."
                ),
            }
        )


# ============================================================
# CREATE CLASSIFICATION DATAFRAME
# ============================================================

classification_columns = [
    "Policy_Record_ID",
    "Candidate_ID",
    "Ballot_Name",
    "Statement_ID",
    "Statement_Type",
    "Policy_Topic",
    "Policy_Subtopic",
    "Position_Summary",
    "Stance",
    "Classification_Method",
    "Matched_Keywords",
    "Keyword_Match_Count",
    "Classification_Score",
    "Classification_Confidence",
    "Human_Reviewed",
    "Reviewer_Notes",
]

classifications = pd.DataFrame(
    classification_rows,
    columns=classification_columns,
)

unclassified_columns = [
    "Statement_ID",
    "Candidate_ID",
    "Ballot_Name",
    "Statement_Type",
    "Statement_Text",
    "Reason",
]

unclassified = pd.DataFrame(
    unclassified_rows,
    columns=unclassified_columns,
)


# ============================================================
# CREATE CANDIDATE-TOPIC SUMMARY
# ============================================================

if classifications.empty:
    candidate_topic_summary = pd.DataFrame(
        columns=[
            "Candidate_ID",
            "Ballot_Name",
            "Policy_Topic",
            "Statement_Count",
            "Total_Keyword_Matches",
            "Average_Classification_Score",
            "Maximum_Confidence",
        ]
    )

else:
    summary_working = (
        classifications.copy()
    )

    summary_working[
        "Confidence_Rank"
    ] = summary_working[
        "Classification_Confidence"
    ].apply(
        confidence_to_rank
    )

    candidate_topic_summary = (
        summary_working
        .groupby(
            [
                "Candidate_ID",
                "Ballot_Name",
                "Policy_Topic",
            ],
            as_index=False,
        )
        .agg(
            Statement_Count=(
                "Statement_ID",
                "nunique",
            ),
            Total_Keyword_Matches=(
                "Keyword_Match_Count",
                "sum",
            ),
            Average_Classification_Score=(
                "Classification_Score",
                "mean",
            ),
            Maximum_Confidence_Rank=(
                "Confidence_Rank",
                "max",
            ),
        )
    )

    candidate_topic_summary[
        "Average_Classification_Score"
    ] = candidate_topic_summary[
        "Average_Classification_Score"
    ].round(2)

    candidate_topic_summary[
        "Maximum_Confidence"
    ] = candidate_topic_summary[
        "Maximum_Confidence_Rank"
    ].apply(
        rank_to_confidence
    )

    candidate_topic_summary.drop(
        columns=[
            "Maximum_Confidence_Rank"
        ],
        inplace=True,
    )

    candidate_topic_summary.sort_values(
        by=[
            "Candidate_ID",
            "Total_Keyword_Matches",
            "Policy_Topic",
        ],
        ascending=[
            True,
            False,
            True,
        ],
        inplace=True,
    )

    candidate_topic_summary.reset_index(
        drop=True,
        inplace=True,
    )


# ============================================================
# CREATE TOPIC-FREQUENCY TABLE
# ============================================================

if classifications.empty:
    topic_frequency = pd.DataFrame(
        columns=[
            "Policy_Topic",
            "Candidate_Count",
            "Statement_Count",
            "Classification_Record_Count",
            "Total_Keyword_Matches",
            "Average_Classification_Score",
        ]
    )

else:
    topic_frequency = (
        classifications
        .groupby(
            "Policy_Topic",
            as_index=False,
        )
        .agg(
            Candidate_Count=(
                "Candidate_ID",
                "nunique",
            ),
            Statement_Count=(
                "Statement_ID",
                "nunique",
            ),
            Classification_Record_Count=(
                "Policy_Record_ID",
                "count",
            ),
            Total_Keyword_Matches=(
                "Keyword_Match_Count",
                "sum",
            ),
            Average_Classification_Score=(
                "Classification_Score",
                "mean",
            ),
        )
    )

    topic_frequency[
        "Average_Classification_Score"
    ] = topic_frequency[
        "Average_Classification_Score"
    ].round(2)

    topic_frequency.sort_values(
        by=[
            "Candidate_Count",
            "Total_Keyword_Matches",
            "Policy_Topic",
        ],
        ascending=[
            False,
            False,
            True,
        ],
        inplace=True,
    )

    topic_frequency.reset_index(
        drop=True,
        inplace=True,
    )


# ============================================================
# CREATE CANDIDATE POLICY MATRIX
# ============================================================

if classifications.empty:
    candidate_policy_matrix = pd.DataFrame()

else:
    candidate_policy_matrix = (
        classifications
        .pivot_table(
            index=[
                "Candidate_ID",
                "Ballot_Name",
            ],
            columns="Policy_Topic",
            values="Keyword_Match_Count",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    candidate_policy_matrix.columns.name = None


# ============================================================
# CREATE STATEMENT COVERAGE SUMMARY
# ============================================================

classified_statement_ids = set(
    classifications[
        "Statement_ID"
    ].unique()
) if not classifications.empty else set()

statement_coverage_rows = []

for candidate_id, group in statements.groupby(
    "Candidate_ID"
):
    candidate_statement_ids = set(
        group["Statement_ID"]
    )

    classified_count = len(
        candidate_statement_ids
        & classified_statement_ids
    )

    total_count = len(
        candidate_statement_ids
    )

    if total_count == 0:
        coverage_percent = 0.0

    else:
        coverage_percent = round(
            classified_count
            / total_count
            * 100,
            2,
        )

    candidate_classification_count = 0
    candidate_topic_count = 0

    if not classifications.empty:
        candidate_classifications = (
            classifications[
                classifications[
                    "Candidate_ID"
                ]
                == str(candidate_id)
            ]
        )

        candidate_classification_count = len(
            candidate_classifications
        )

        candidate_topic_count = (
            candidate_classifications[
                "Policy_Topic"
            ].nunique()
        )

    statement_coverage_rows.append(
        {
            "Candidate_ID": candidate_id,
            "Ballot_Name": (
                candidate_name_lookup.get(
                    str(candidate_id),
                    "",
                )
            ),
            "Total_Statements": total_count,
            "Classified_Statements": (
                classified_count
            ),
            "Unclassified_Statements": (
                total_count
                - classified_count
            ),
            "Classification_Coverage_Percent": (
                coverage_percent
            ),
            "Policy_Classification_Records": (
                candidate_classification_count
            ),
            "Unique_Policy_Topics": (
                candidate_topic_count
            ),
        }
    )

statement_coverage = pd.DataFrame(
    statement_coverage_rows
)


# ============================================================
# CREATE KEYWORD DICTIONARY TABLE
# ============================================================

keyword_rows = []

for (
    policy_topic,
    keywords,
) in POLICY_KEYWORDS.items():

    for keyword in keywords:
        keyword_rows.append(
            {
                "Policy_Topic": policy_topic,
                "Keyword_or_Phrase": keyword,
                "Word_Count": len(
                    re.split(
                        r"[\s\-]+",
                        keyword.strip(),
                    )
                ),
            }
        )

keyword_dictionary = pd.DataFrame(
    keyword_rows
)


# ============================================================
# CREATE RUN SUMMARY
# ============================================================

run_summary = pd.DataFrame(
    [
        {
            "Metric": "Input Workbook",
            "Value": str(INPUT_FILE),
        },
        {
            "Metric": "Statements Processed",
            "Value": len(statements),
        },
        {
            "Metric": "Policy Classification Records",
            "Value": len(classifications),
        },
        {
            "Metric": "Unique Policy Topics Detected",
            "Value": (
                classifications[
                    "Policy_Topic"
                ].nunique()
                if not classifications.empty
                else 0
            ),
        },
        {
            "Metric": "Candidates Classified",
            "Value": (
                classifications[
                    "Candidate_ID"
                ].nunique()
                if not classifications.empty
                else 0
            ),
        },
        {
            "Metric": "Unclassified Statements",
            "Value": len(unclassified),
        },
        {
            "Metric": "Classification Method",
            "Value": "Rule-Based NLP",
        },
    ]
)


# ============================================================
# SAVE OUTPUTS
# ============================================================

OUTPUT_EXCEL.parent.mkdir(
    parents=True,
    exist_ok=True,
)

classifications.to_csv(
    OUTPUT_CSV,
    index=False,
    encoding="utf-8-sig",
)

with pd.ExcelWriter(
    OUTPUT_EXCEL,
    engine="openpyxl",
) as writer:

    run_summary.to_excel(
        writer,
        sheet_name="RunSummary",
        index=False,
    )

    classifications.to_excel(
        writer,
        sheet_name="PolicyClassifications",
        index=False,
    )

    candidate_topic_summary.to_excel(
        writer,
        sheet_name="CandidateTopicSummary",
        index=False,
    )

    topic_frequency.to_excel(
        writer,
        sheet_name="TopicFrequency",
        index=False,
    )

    candidate_policy_matrix.to_excel(
        writer,
        sheet_name="CandidatePolicyMatrix",
        index=False,
    )

    statement_coverage.to_excel(
        writer,
        sheet_name="StatementCoverage",
        index=False,
    )

    unclassified.to_excel(
        writer,
        sheet_name="UnclassifiedStatements",
        index=False,
    )

    keyword_dictionary.to_excel(
        writer,
        sheet_name="KeywordDictionary",
        index=False,
    )

    output_workbook = writer.book

    for worksheet in output_workbook.worksheets:
        worksheet.freeze_panes = "A2"

        if worksheet.max_row >= 1:
            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

        for column_cells in worksheet.columns:
            max_length = 0

            column_letter = (
                column_cells[0]
                .column_letter
            )

            for cell in column_cells:
                if cell.value is not None:
                    max_length = max(
                        max_length,
                        len(str(cell.value)),
                    )

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max(
                    max_length + 2,
                    12,
                ),
                80,
            )


# ============================================================
# TERMINAL OUTPUT
# ============================================================

print("=" * 72)
print("EXPANDED POLICY TOPIC CLASSIFICATION COMPLETE")
print("=" * 72)

print(f"Input workbook:\n{INPUT_FILE}")
print()

print(f"Excel output:\n{OUTPUT_EXCEL}")
print()

print(f"CSV output:\n{OUTPUT_CSV}")
print()

print(
    "Statements processed: "
    f"{len(statements)}"
)

print(
    "Policy classification records: "
    f"{len(classifications)}"
)

print(
    "Unique policy topics detected: "
    f"{classifications['Policy_Topic'].nunique() if not classifications.empty else 0}"
)

print(
    "Candidates classified: "
    f"{classifications['Candidate_ID'].nunique() if not classifications.empty else 0}"
)

print(
    "Unclassified statements: "
    f"{len(unclassified)}"
)

if not topic_frequency.empty:
    print()
    print("Most frequently detected topics:")

    for _, row in (
        topic_frequency
        .head(10)
        .iterrows()
    ):
        print(
            f"  - {row['Policy_Topic']}: "
            f"{row['Candidate_Count']} candidate(s), "
            f"{row['Statement_Count']} statement(s), "
            f"{row['Total_Keyword_Matches']} keyword match(es)"
        )

print()

print(
    "Next step: review RunSummary, "
    "PolicyClassifications, CandidateTopicSummary, "
    "CandidatePolicyMatrix, and UnclassifiedStatements."
)