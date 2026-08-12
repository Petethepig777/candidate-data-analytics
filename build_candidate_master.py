from pathlib import Path
from datetime import date
import re

import pandas as pd


# ============================================================
# FILE PATHS
# ============================================================

PROJECT_FOLDER = Path.home() / "Desktop" / "POLITICAL_CANDIDATE_ANALYTICS"

INPUT_FILE = (
    PROJECT_FOLDER
    / "data"
    / "raw"
    / "political_candidate_database_pilot.xlsx"
)

OUTPUT_EXCEL = (
    PROJECT_FOLDER
    / "data"
    / "processed"
    / "candidate_master_clean.xlsx"
)

OUTPUT_CSV = (
    PROJECT_FOLDER
    / "data"
    / "processed"
    / "candidate_master_clean.csv"
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):
    """
    Standardize text by trimming whitespace and collapsing
    repeated internal spaces.
    """
    if pd.isna(value):
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value).strip(),
    )


def clean_phone(value):
    """
    Normalize a U.S. phone number to ###-###-#### when possible.
    Preserve the original cleaned text if it cannot be normalized.
    """
    text = clean_text(value)

    if not text:
        return ""

    digits = re.sub(r"\D", "", text)

    if len(digits) == 10:
        return (
            f"{digits[0:3]}-"
            f"{digits[3:6]}-"
            f"{digits[6:10]}"
        )

    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]

        return (
            f"{digits[0:3]}-"
            f"{digits[3:6]}-"
            f"{digits[6:10]}"
        )

    return text


def clean_url(value):
    """
    Standardize URL text and remove trailing spaces.
    """
    return clean_text(value)


def calculate_age(dob_value, reference_date=None):
    """
    Calculate age from DOB.

    Returns a blank value when DOB is missing or invalid.
    """
    if reference_date is None:
        reference_date = date.today()

    if pd.isna(dob_value) or clean_text(dob_value) == "":
        return ""

    parsed_dob = pd.to_datetime(
        dob_value,
        errors="coerce",
    )

    if pd.isna(parsed_dob):
        return ""

    age = (
        reference_date.year
        - parsed_dob.year
        - (
            (
                reference_date.month,
                reference_date.day,
            )
            < (
                parsed_dob.month,
                parsed_dob.day,
            )
        )
    )

    return int(age)


def count_records_by_candidate(
    dataframe,
    candidate_column="Candidate_ID",
):
    """
    Return a Series with the number of records for each candidate.
    """
    if dataframe.empty:
        return pd.Series(dtype="int64")

    if candidate_column not in dataframe.columns:
        return pd.Series(dtype="int64")

    working = dataframe.copy()

    working[candidate_column] = (
        working[candidate_column]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    working = working[
        working[candidate_column] != ""
    ]

    return working.groupby(
        candidate_column
    ).size()


def count_verified_sources(dataframe):
    """
    Count fully verified sources for each candidate.
    """
    if dataframe.empty:
        return pd.Series(dtype="int64")

    if (
        "Candidate_ID" not in dataframe.columns
        or "Verified" not in dataframe.columns
    ):
        return pd.Series(dtype="int64")

    working = dataframe.copy()

    working["Candidate_ID"] = (
        working["Candidate_ID"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    verified_text = (
        working["Verified"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    verified_rows = working[
        verified_text == "yes"
    ]

    return verified_rows.groupby(
        "Candidate_ID"
    ).size()


def count_high_reliability_sources(dataframe):
    """
    Count high-reliability sources for each candidate.
    """
    if dataframe.empty:
        return pd.Series(dtype="int64")

    if (
        "Candidate_ID" not in dataframe.columns
        or "Reliability_Level" not in dataframe.columns
    ):
        return pd.Series(dtype="int64")

    working = dataframe.copy()

    working["Candidate_ID"] = (
        working["Candidate_ID"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    reliability_text = (
        working["Reliability_Level"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    high_rows = working[
        reliability_text == "high"
    ]

    return high_rows.groupby(
        "Candidate_ID"
    ).size()


def calculate_research_completeness(row, fields):
    """
    Calculate percentage of selected fields that are populated.
    """
    present = 0

    for field in fields:
        value = row.get(field, "")

        if clean_text(value) != "":
            present += 1

    if not fields:
        return 0.0

    return round(
        present / len(fields) * 100,
        2,
    )


# ============================================================
# LOAD WORKBOOK
# ============================================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Input workbook was not found:\n{INPUT_FILE}"
    )

sheet_names = [
    "Candidates",
    "Sources",
    "CandidateFacts",
    "CandidateStatements",
    "PolicyTopics",
    "SocialMedia",
]

tables = {
    sheet_name: pd.read_excel(
        INPUT_FILE,
        sheet_name=sheet_name,
        dtype=object,
    ).dropna(how="all")
    for sheet_name in sheet_names
}

candidates = tables["Candidates"].copy()
sources = tables["Sources"].copy()
facts = tables["CandidateFacts"].copy()
statements = tables["CandidateStatements"].copy()
policy_topics = tables["PolicyTopics"].copy()
social_media = tables["SocialMedia"].copy()


# ============================================================
# STANDARDIZE CANDIDATE DATA
# ============================================================

text_columns = [
    "Candidate_ID",
    "Ballot_Name",
    "Full_Name",
    "First_Name",
    "Middle_Name",
    "Last_Name",
    "Suffix",
    "Party",
    "Office",
    "State",
    "District",
    "Incumbent",
    "Candidate_Status",
    "Birthplace",
    "Residence_City",
    "Residence_State",
    "Campaign_Website",
    "Campaign_Email",
    "Campaign_Phone",
    "Campaign_Address",
    "Research_Status",
    "Researcher",
    "Candidate_Notes",
]

for column in text_columns:
    if column in candidates.columns:
        candidates[column] = candidates[
            column
        ].apply(clean_text)

if "Campaign_Phone" in candidates.columns:
    candidates["Campaign_Phone"] = candidates[
        "Campaign_Phone"
    ].apply(clean_phone)

if "Campaign_Website" in candidates.columns:
    candidates["Campaign_Website"] = candidates[
        "Campaign_Website"
    ].apply(clean_url)


# ============================================================
# STANDARDIZE CATEGORIES
# ============================================================

party_map = {
    "dfl": "Democratic-Farmer-Labor",
    "democratic farmer labor": "Democratic-Farmer-Labor",
    "democratic-farmer-labor": "Democratic-Farmer-Labor",
    "democrat": "Democratic-Farmer-Labor",
    "democratic": "Democratic-Farmer-Labor",
    "gop": "Republican",
    "republican": "Republican",
    "independent": "Independent",
}

if "Party" in candidates.columns:
    candidates["Party"] = candidates[
        "Party"
    ].apply(
        lambda value: party_map.get(
            clean_text(value).lower(),
            clean_text(value),
        )
    )

office_map = {
    "state senator": "State Senator",
    "state representative": "State Representative",
    "senator": "State Senator",
    "representative": "State Representative",
}

if "Office" in candidates.columns:
    candidates["Office"] = candidates[
        "Office"
    ].apply(
        lambda value: office_map.get(
            clean_text(value).lower(),
            clean_text(value),
        )
    )

if "State" in candidates.columns:
    candidates["State"] = candidates[
        "State"
    ].apply(
        lambda value: (
            "Minnesota"
            if clean_text(value).lower()
            in {"mn", "minnesota"}
            else clean_text(value)
        )
    )

if "Residence_State" in candidates.columns:
    candidates["Residence_State"] = candidates[
        "Residence_State"
    ].apply(
        lambda value: (
            "Minnesota"
            if clean_text(value).lower()
            in {"mn", "minnesota"}
            else clean_text(value)
        )
    )


# ============================================================
# STANDARDIZE DATES
# ============================================================

date_columns = [
    "Filing_Date",
    "DOB",
    "Date_First_Researched",
    "Date_Last_Updated",
]

for column in date_columns:
    if column in candidates.columns:
        candidates[column] = pd.to_datetime(
            candidates[column],
            errors="coerce",
        )


# ============================================================
# CREATE DERIVED CANDIDATE FIELDS
# ============================================================

REFERENCE_DATE = date(2026, 8, 5)

candidates["Age_As_Of_2026_08_05"] = candidates[
    "DOB"
].apply(
    lambda value: calculate_age(
        value,
        REFERENCE_DATE,
    )
)

if "Birth_Year" in candidates.columns:
    candidates["Birth_Year"] = pd.to_numeric(
        candidates["Birth_Year"],
        errors="coerce",
    ).astype("Int64")

derived_birth_year = candidates[
    "DOB"
].dt.year.astype("Int64")

if "Birth_Year" not in candidates.columns:
    candidates["Birth_Year"] = derived_birth_year

else:
    candidates["Birth_Year"] = (
        candidates["Birth_Year"]
        .fillna(derived_birth_year)
    )

candidates["District_Label"] = (
    candidates["State"]
    + " "
    + candidates["Office"]
    + " District "
    + candidates["District"]
)


# ============================================================
# ADD RELATED TABLE RECORD COUNTS
# ============================================================

source_counts = count_records_by_candidate(
    sources
)

verified_source_counts = count_verified_sources(
    sources
)

high_reliability_counts = (
    count_high_reliability_sources(
        sources
    )
)

fact_counts = count_records_by_candidate(
    facts
)

statement_counts = count_records_by_candidate(
    statements
)

policy_counts = count_records_by_candidate(
    policy_topics
)

social_counts = count_records_by_candidate(
    social_media
)

candidates["Source_Count"] = (
    candidates["Candidate_ID"]
    .map(source_counts)
    .fillna(0)
    .astype(int)
)

candidates["Verified_Source_Count"] = (
    candidates["Candidate_ID"]
    .map(verified_source_counts)
    .fillna(0)
    .astype(int)
)

candidates["High_Reliability_Source_Count"] = (
    candidates["Candidate_ID"]
    .map(high_reliability_counts)
    .fillna(0)
    .astype(int)
)

candidates["Fact_Count"] = (
    candidates["Candidate_ID"]
    .map(fact_counts)
    .fillna(0)
    .astype(int)
)

candidates["Statement_Count"] = (
    candidates["Candidate_ID"]
    .map(statement_counts)
    .fillna(0)
    .astype(int)
)

candidates["Policy_Record_Count"] = (
    candidates["Candidate_ID"]
    .map(policy_counts)
    .fillna(0)
    .astype(int)
)

candidates["Social_Media_Count"] = (
    candidates["Candidate_ID"]
    .map(social_counts)
    .fillna(0)
    .astype(int)
)


# ============================================================
# CALCULATE MISSINGNESS AND COMPLETENESS
# ============================================================

analysis_fields = [
    "Full_Name",
    "First_Name",
    "Last_Name",
    "Party",
    "Office",
    "State",
    "District",
    "Incumbent",
    "Candidate_Status",
    "Filing_Date",
    "DOB",
    "Birthplace",
    "Residence_City",
    "Campaign_Website",
    "Campaign_Email",
    "Campaign_Phone",
    "Campaign_Address",
]

candidates["Research_Completeness_Percent"] = (
    candidates.apply(
        lambda row: calculate_research_completeness(
            row,
            analysis_fields,
        ),
        axis=1,
    )
)

candidates["Missing_Analysis_Field_Count"] = (
    candidates.apply(
        lambda row: sum(
            clean_text(
                row.get(field, "")
            )
            == ""
            for field in analysis_fields
        ),
        axis=1,
    )
)

candidates["Has_Verified_DOB"] = candidates[
    "DOB"
].apply(
    lambda value: "Yes"
    if not pd.isna(value)
    else "No"
)

candidates["Has_Campaign_Website"] = candidates[
    "Campaign_Website"
].apply(
    lambda value: "Yes"
    if clean_text(value) != ""
    else "No"
)

candidates["Has_Campaign_Email"] = candidates[
    "Campaign_Email"
].apply(
    lambda value: "Yes"
    if clean_text(value) != ""
    else "No"
)

candidates["Has_Campaign_Phone"] = candidates[
    "Campaign_Phone"
].apply(
    lambda value: "Yes"
    if clean_text(value) != ""
    else "No"
)

candidates["Has_Social_Media"] = candidates[
    "Social_Media_Count"
].apply(
    lambda value: "Yes"
    if value > 0
    else "No"
)


# ============================================================
# REORDER OUTPUT COLUMNS
# ============================================================

preferred_columns = [
    "Candidate_ID",
    "Ballot_Name",
    "Full_Name",
    "First_Name",
    "Middle_Name",
    "Last_Name",
    "Suffix",
    "Party",
    "Office",
    "State",
    "District",
    "District_Label",
    "Incumbent",
    "Candidate_Status",
    "Filing_Date",
    "DOB",
    "Birth_Year",
    "Age_As_Of_2026_08_05",
    "Birthplace",
    "Residence_City",
    "Residence_State",
    "Campaign_Website",
    "Campaign_Email",
    "Campaign_Phone",
    "Campaign_Address",
    "Research_Status",
    "Researcher",
    "Date_First_Researched",
    "Date_Last_Updated",
    "Source_Count",
    "Verified_Source_Count",
    "High_Reliability_Source_Count",
    "Fact_Count",
    "Statement_Count",
    "Policy_Record_Count",
    "Social_Media_Count",
    "Research_Completeness_Percent",
    "Missing_Analysis_Field_Count",
    "Has_Verified_DOB",
    "Has_Campaign_Website",
    "Has_Campaign_Email",
    "Has_Campaign_Phone",
    "Has_Social_Media",
    "Candidate_Notes",
]

existing_preferred_columns = [
    column
    for column in preferred_columns
    if column in candidates.columns
]

remaining_columns = [
    column
    for column in candidates.columns
    if column not in existing_preferred_columns
]

candidate_master = candidates[
    existing_preferred_columns
    + remaining_columns
].copy()


# ============================================================
# CREATE SUMMARY TABLE
# ============================================================

summary_rows = [
    {
        "Metric": "Candidate Count",
        "Value": len(candidate_master),
    },
    {
        "Metric": "District Count",
        "Value": candidate_master[
            "District_Label"
        ].nunique(),
    },
    {
        "Metric": "Incumbent Count",
        "Value": int(
            (
                candidate_master[
                    "Incumbent"
                ].str.lower()
                == "yes"
            ).sum()
        ),
    },
    {
        "Metric": "Candidates with Verified DOB",
        "Value": int(
            (
                candidate_master[
                    "Has_Verified_DOB"
                ]
                == "Yes"
            ).sum()
        ),
    },
    {
        "Metric": "Candidates with Campaign Website",
        "Value": int(
            (
                candidate_master[
                    "Has_Campaign_Website"
                ]
                == "Yes"
            ).sum()
        ),
    },
    {
        "Metric": "Candidates with Social Media",
        "Value": int(
            (
                candidate_master[
                    "Has_Social_Media"
                ]
                == "Yes"
            ).sum()
        ),
    },
    {
        "Metric": "Average Research Completeness",
        "Value": round(
            candidate_master[
                "Research_Completeness_Percent"
            ].mean(),
            2,
        ),
    },
    {
        "Metric": "Total Sources",
        "Value": int(
            candidate_master[
                "Source_Count"
            ].sum()
        ),
    },
    {
        "Metric": "Total Facts",
        "Value": int(
            candidate_master[
                "Fact_Count"
            ].sum()
        ),
    },
    {
        "Metric": "Total Statements",
        "Value": int(
            candidate_master[
                "Statement_Count"
            ].sum()
        ),
    },
    {
        "Metric": "Total Policy Records",
        "Value": int(
            candidate_master[
                "Policy_Record_Count"
            ].sum()
        ),
    },
]

summary = pd.DataFrame(
    summary_rows
)


# ============================================================
# SAVE OUTPUTS
# ============================================================

OUTPUT_EXCEL.parent.mkdir(
    parents=True,
    exist_ok=True,
)

candidate_master.to_csv(
    OUTPUT_CSV,
    index=False,
    encoding="utf-8-sig",
    date_format="%Y-%m-%d",
)

with pd.ExcelWriter(
    OUTPUT_EXCEL,
    engine="openpyxl",
    date_format="YYYY-MM-DD",
    datetime_format="YYYY-MM-DD",
) as writer:

    candidate_master.to_excel(
        writer,
        sheet_name="CandidateMaster",
        index=False,
    )

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False,
    )

    sources.to_excel(
        writer,
        sheet_name="Sources",
        index=False,
    )

    facts.to_excel(
        writer,
        sheet_name="CandidateFacts",
        index=False,
    )

    statements.to_excel(
        writer,
        sheet_name="CandidateStatements",
        index=False,
    )

    policy_topics.to_excel(
        writer,
        sheet_name="PolicyTopics",
        index=False,
    )

    social_media.to_excel(
        writer,
        sheet_name="SocialMedia",
        index=False,
    )

    workbook = writer.book

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
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
                max(max_length + 2, 12),
                60,
            )


# ============================================================
# TERMINAL OUTPUT
# ============================================================

print("=" * 72)
print("CANDIDATE MASTER DATASET CREATED")
print("=" * 72)

print(f"Input workbook:\n{INPUT_FILE}")
print()

print(f"Excel output:\n{OUTPUT_EXCEL}")
print()

print(f"CSV output:\n{OUTPUT_CSV}")
print()

print(f"Candidates processed: {len(candidate_master)}")

print(
    "Average research completeness: "
    f"{candidate_master['Research_Completeness_Percent'].mean():.2f}%"
)

print(
    "Candidates with verified DOB: "
    f"{(candidate_master['Has_Verified_DOB'] == 'Yes').sum()}"
)

print(
    "Candidates with social media: "
    f"{(candidate_master['Has_Social_Media'] == 'Yes').sum()}"
)

print()
print(
    "Next step: review CandidateMaster and Summary "
    "in candidate_master_clean.xlsx."
)