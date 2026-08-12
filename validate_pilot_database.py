from pathlib import Path
from datetime import date, datetime
from urllib.parse import urlparse

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

OUTPUT_FILE = (
    PROJECT_FOLDER
    / "data"
    / "processed"
    / "pilot_validation_report.xlsx"
)


# ============================================================
# CONFIGURATION
# ============================================================

REQUIRED_SHEETS = [
    "Candidates",
    "Sources",
    "CandidateFacts",
    "CandidateStatements",
    "PolicyTopics",
    "SocialMedia",
]

REQUIRED_CANDIDATE_FIELDS = [
    "Candidate_ID",
    "Ballot_Name",
    "Party",
    "Office",
    "State",
    "District",
    "Incumbent",
    "Candidate_Status",
    "Research_Status",
]

IMPORTANT_CANDIDATE_FIELDS = [
    "Full_Name",
    "First_Name",
    "Last_Name",
    "Filing_Date",
    "DOB",
    "Birthplace",
    "Residence_City",
    "Campaign_Website",
    "Campaign_Email",
    "Campaign_Phone",
    "Campaign_Address",
]

VALID_URL_COLUMNS = {
    "Candidates": ["Campaign_Website"],
    "Sources": ["URL", "Archive_URL"],
    "SocialMedia": ["Profile_URL"],
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def is_missing(value):
    """
    Return True when a value should be treated as missing.
    """
    if pd.isna(value):
        return True

    if isinstance(value, str):
        return value.strip() == ""

    return False


def is_valid_url(value):
    """
    Validate basic HTTP or HTTPS URL structure.

    Blank URLs are permitted because some URL fields are optional.
    """
    if is_missing(value):
        return True

    try:
        parsed = urlparse(str(value).strip())

        return (
            parsed.scheme in {"http", "https"}
            and bool(parsed.netloc)
        )

    except Exception:
        return False


def normalize_text(value):
    """
    Convert a value into standardized lowercase text.
    """
    if is_missing(value):
        return ""

    return " ".join(
        str(value).strip().lower().split()
    )


def add_issue(
    issues,
    severity,
    table_name,
    record_id,
    field_name,
    issue_type,
    message,
):
    """
    Add one issue to the validation issue list.
    """
    issues.append(
        {
            "Severity": severity,
            "Table": table_name,
            "Record_ID": record_id,
            "Field": field_name,
            "Issue_Type": issue_type,
            "Message": message,
        }
    )


def calculate_percentage(numerator, denominator):
    """
    Safely calculate a percentage.
    """
    if denominator == 0:
        return 0.0

    return round(
        (numerator / denominator) * 100,
        2,
    )


def count_candidate_records(
    dataframe,
    candidate_id,
):
    """
    Count records associated with one candidate.

    This function safely handles empty DataFrames.
    """
    if dataframe.empty:
        return 0

    if "Candidate_ID" not in dataframe.columns:
        return 0

    normalized_candidate_id = str(
        candidate_id
    ).strip()

    return int(
        (
            dataframe["Candidate_ID"]
            .fillna("")
            .astype(str)
            .str.strip()
            == normalized_candidate_id
        ).sum()
    )


# ============================================================
# LOAD WORKBOOK
# ============================================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Pilot workbook was not found:\n{INPUT_FILE}"
    )

excel_file = pd.ExcelFile(INPUT_FILE)

missing_sheets = [
    sheet_name
    for sheet_name in REQUIRED_SHEETS
    if sheet_name not in excel_file.sheet_names
]

if missing_sheets:
    raise ValueError(
        "The following required worksheets are missing: "
        + ", ".join(missing_sheets)
    )

tables = {
    sheet_name: pd.read_excel(
        INPUT_FILE,
        sheet_name=sheet_name,
        dtype=object,
    )
    for sheet_name in REQUIRED_SHEETS
}


# ============================================================
# REMOVE COMPLETELY BLANK ROWS
# ============================================================

for table_name, dataframe in tables.items():
    tables[table_name] = (
        dataframe
        .dropna(how="all")
        .reset_index(drop=True)
    )

candidates = tables["Candidates"]
sources = tables["Sources"]
facts = tables["CandidateFacts"]
statements = tables["CandidateStatements"]
policy_topics = tables["PolicyTopics"]
social_media = tables["SocialMedia"]

issues = []


# ============================================================
# CHECK REQUIRED CANDIDATE COLUMNS
# ============================================================

for field_name in REQUIRED_CANDIDATE_FIELDS:
    if field_name not in candidates.columns:
        add_issue(
            issues,
            "Critical",
            "Candidates",
            "",
            field_name,
            "Missing Column",
            f"Required column '{field_name}' is missing.",
        )


# ============================================================
# CHECK REQUIRED CANDIDATE VALUES
# ============================================================

for row_index, row in candidates.iterrows():
    candidate_id = row.get(
        "Candidate_ID",
        f"Candidate row {row_index + 2}",
    )

    for field_name in REQUIRED_CANDIDATE_FIELDS:
        if field_name not in candidates.columns:
            continue

        if is_missing(row.get(field_name)):
            add_issue(
                issues,
                "Error",
                "Candidates",
                candidate_id,
                field_name,
                "Missing Required Value",
                f"Required field '{field_name}' is blank.",
            )


# ============================================================
# CHECK DUPLICATE PRIMARY KEYS
# ============================================================

primary_keys = {
    "Candidates": "Candidate_ID",
    "Sources": "Source_ID",
    "CandidateFacts": "Fact_ID",
    "CandidateStatements": "Statement_ID",
    "PolicyTopics": "Policy_Record_ID",
    "SocialMedia": "Social_ID",
}

for table_name, key_column in primary_keys.items():
    dataframe = tables[table_name]

    if key_column not in dataframe.columns:
        add_issue(
            issues,
            "Critical",
            table_name,
            "",
            key_column,
            "Missing Primary Key Column",
            f"Primary-key column '{key_column}' is missing.",
        )
        continue

    # This form safely handles an empty DataFrame.
    nonblank_mask = ~dataframe[
        key_column
    ].apply(is_missing)

    nonblank_keys = dataframe.loc[
        nonblank_mask
    ].copy()

    if nonblank_keys.empty:
        continue

    duplicate_mask = nonblank_keys[
        key_column
    ].duplicated(keep=False)

    duplicate_rows = nonblank_keys.loc[
        duplicate_mask
    ]

    for _, duplicate_row in duplicate_rows.iterrows():
        duplicate_id = duplicate_row[
            key_column
        ]

        add_issue(
            issues,
            "Critical",
            table_name,
            duplicate_id,
            key_column,
            "Duplicate Primary Key",
            f"Duplicate primary key found: {duplicate_id}",
        )


# ============================================================
# CHECK MISSING PRIMARY KEYS
# ============================================================

for table_name, key_column in primary_keys.items():
    dataframe = tables[table_name]

    if key_column not in dataframe.columns:
        continue

    for row_index, row in dataframe.iterrows():
        if is_missing(row.get(key_column)):
            add_issue(
                issues,
                "Critical",
                table_name,
                f"Row {row_index + 2}",
                key_column,
                "Missing Primary Key",
                f"Primary-key field '{key_column}' is blank.",
            )


# ============================================================
# CHECK DUPLICATE CANDIDATE NAMES
# ============================================================

if (
    not candidates.empty
    and "Ballot_Name" in candidates.columns
):
    normalized_names = candidates[
        "Ballot_Name"
    ].apply(normalize_text)

    duplicate_name_mask = (
        normalized_names.ne("")
        & normalized_names.duplicated(
            keep=False
        )
    )

    duplicate_candidate_rows = candidates.loc[
        duplicate_name_mask
    ]

    for _, row in duplicate_candidate_rows.iterrows():
        add_issue(
            issues,
            "Warning",
            "Candidates",
            row.get("Candidate_ID", ""),
            "Ballot_Name",
            "Possible Duplicate Candidate",
            (
                "Another candidate record has the same "
                "normalized ballot name: "
                f"{row.get('Ballot_Name')}"
            ),
        )


# ============================================================
# BUILD PRIMARY-KEY SETS
# ============================================================

candidate_ids = set()

if "Candidate_ID" in candidates.columns:
    candidate_ids = {
        str(value).strip()
        for value in candidates[
            "Candidate_ID"
        ]
        if not is_missing(value)
    }

source_ids = set()

if "Source_ID" in sources.columns:
    source_ids = {
        str(value).strip()
        for value in sources[
            "Source_ID"
        ]
        if not is_missing(value)
    }

statement_ids = set()

if "Statement_ID" in statements.columns:
    statement_ids = {
        str(value).strip()
        for value in statements[
            "Statement_ID"
        ]
        if not is_missing(value)
    }


# ============================================================
# CHECK CANDIDATE FOREIGN-KEY RELATIONSHIPS
# ============================================================

candidate_reference_tables = [
    (
        "Sources",
        sources,
        "Candidate_ID",
        "Source_ID",
    ),
    (
        "CandidateFacts",
        facts,
        "Candidate_ID",
        "Fact_ID",
    ),
    (
        "CandidateStatements",
        statements,
        "Candidate_ID",
        "Statement_ID",
    ),
    (
        "PolicyTopics",
        policy_topics,
        "Candidate_ID",
        "Policy_Record_ID",
    ),
    (
        "SocialMedia",
        social_media,
        "Candidate_ID",
        "Social_ID",
    ),
]

for (
    table_name,
    dataframe,
    foreign_key,
    record_key,
) in candidate_reference_tables:

    if dataframe.empty:
        continue

    if foreign_key not in dataframe.columns:
        add_issue(
            issues,
            "Critical",
            table_name,
            "",
            foreign_key,
            "Missing Foreign Key Column",
            (
                f"Foreign-key column "
                f"'{foreign_key}' is missing."
            ),
        )
        continue

    for row_index, row in dataframe.iterrows():
        candidate_id = row.get(foreign_key)

        if is_missing(candidate_id):
            add_issue(
                issues,
                "Error",
                table_name,
                row.get(
                    record_key,
                    f"Row {row_index + 2}",
                ),
                foreign_key,
                "Missing Candidate Reference",
                "Candidate_ID is blank.",
            )
            continue

        candidate_id = str(
            candidate_id
        ).strip()

        if candidate_id not in candidate_ids:
            add_issue(
                issues,
                "Critical",
                table_name,
                row.get(
                    record_key,
                    f"Row {row_index + 2}",
                ),
                foreign_key,
                "Invalid Candidate Reference",
                (
                    f"Candidate_ID '{candidate_id}' "
                    "does not exist in the "
                    "Candidates table."
                ),
            )


# ============================================================
# CHECK SOURCE FOREIGN-KEY RELATIONSHIPS
# ============================================================

source_reference_tables = [
    (
        "CandidateFacts",
        facts,
        "Source_ID",
        "Fact_ID",
    ),
    (
        "CandidateStatements",
        statements,
        "Primary_Source_ID",
        "Statement_ID",
    ),
]

for (
    table_name,
    dataframe,
    source_field,
    record_key,
) in source_reference_tables:

    if dataframe.empty:
        continue

    if source_field not in dataframe.columns:
        add_issue(
            issues,
            "Critical",
            table_name,
            "",
            source_field,
            "Missing Source Reference Column",
            (
                f"Source-reference column "
                f"'{source_field}' is missing."
            ),
        )
        continue

    for row_index, row in dataframe.iterrows():
        source_id = row.get(source_field)

        if is_missing(source_id):
            add_issue(
                issues,
                "Warning",
                table_name,
                row.get(
                    record_key,
                    f"Row {row_index + 2}",
                ),
                source_field,
                "Missing Source Reference",
                (
                    f"Source reference "
                    f"'{source_field}' is blank."
                ),
            )
            continue

        source_id = str(source_id).strip()

        if source_id not in source_ids:
            add_issue(
                issues,
                "Critical",
                table_name,
                row.get(
                    record_key,
                    f"Row {row_index + 2}",
                ),
                source_field,
                "Invalid Source Reference",
                (
                    f"Source_ID '{source_id}' "
                    "does not exist in the "
                    "Sources table."
                ),
            )


# ============================================================
# CHECK POLICY-TO-STATEMENT REFERENCES
# ============================================================

if (
    not policy_topics.empty
    and "Statement_ID" in policy_topics.columns
):
    for row_index, row in policy_topics.iterrows():
        statement_id = row.get("Statement_ID")

        if is_missing(statement_id):
            continue

        statement_id = str(
            statement_id
        ).strip()

        if statement_id not in statement_ids:
            add_issue(
                issues,
                "Critical",
                "PolicyTopics",
                row.get(
                    "Policy_Record_ID",
                    f"Row {row_index + 2}",
                ),
                "Statement_ID",
                "Invalid Statement Reference",
                (
                    f"Statement_ID '{statement_id}' "
                    "does not exist in the "
                    "CandidateStatements table."
                ),
            )


# ============================================================
# CHECK URL FORMAT
# ============================================================

for table_name, url_columns in VALID_URL_COLUMNS.items():
    dataframe = tables[table_name]

    if dataframe.empty:
        continue

    for url_column in url_columns:
        if url_column not in dataframe.columns:
            continue

        for row_index, row in dataframe.iterrows():
            value = row.get(url_column)

            if not is_valid_url(value):
                record_key = primary_keys.get(
                    table_name
                )

                if record_key:
                    record_id = row.get(
                        record_key,
                        f"Row {row_index + 2}",
                    )
                else:
                    record_id = (
                        f"Row {row_index + 2}"
                    )

                add_issue(
                    issues,
                    "Error",
                    table_name,
                    record_id,
                    url_column,
                    "Invalid URL",
                    f"Invalid URL format: {value}",
                )


# ============================================================
# CHECK DOB AND BIRTH-YEAR CONSISTENCY
# ============================================================

if not candidates.empty:
    for _, row in candidates.iterrows():
        candidate_id = row.get(
            "Candidate_ID",
            "",
        )

        dob_value = row.get("DOB")
        birth_year_value = row.get(
            "Birth_Year"
        )

        if is_missing(dob_value):
            continue

        try:
            parsed_dob = pd.to_datetime(
                dob_value,
                errors="raise",
            )

            dob_year = int(parsed_dob.year)

            if not is_missing(
                birth_year_value
            ):
                try:
                    stored_birth_year = int(
                        float(
                            birth_year_value
                        )
                    )

                    if stored_birth_year != dob_year:
                        add_issue(
                            issues,
                            "Error",
                            "Candidates",
                            candidate_id,
                            "Birth_Year",
                            "DOB Conflict",
                            (
                                f"DOB year is "
                                f"{dob_year}, but "
                                f"Birth_Year is "
                                f"{stored_birth_year}."
                            ),
                        )

                except (
                    TypeError,
                    ValueError,
                ):
                    add_issue(
                        issues,
                        "Error",
                        "Candidates",
                        candidate_id,
                        "Birth_Year",
                        "Invalid Birth Year",
                        (
                            "Birth_Year cannot "
                            "be converted to an "
                            "integer: "
                            f"{birth_year_value}"
                        ),
                    )

            if (
                dob_year < 1900
                or dob_year > date.today().year
            ):
                add_issue(
                    issues,
                    "Error",
                    "Candidates",
                    candidate_id,
                    "DOB",
                    "Implausible DOB",
                    (
                        "DOB year appears "
                        f"implausible: {dob_year}"
                    ),
                )

        except Exception:
            add_issue(
                issues,
                "Error",
                "Candidates",
                candidate_id,
                "DOB",
                "Invalid Date",
                (
                    "DOB could not be parsed: "
                    f"{dob_value}"
                ),
            )


# ============================================================
# CHECK FILING DATE FORMAT
# ============================================================

if (
    not candidates.empty
    and "Filing_Date" in candidates.columns
):
    for _, row in candidates.iterrows():
        filing_date = row.get(
            "Filing_Date"
        )

        if is_missing(filing_date):
            continue

        try:
            parsed_filing_date = pd.to_datetime(
                filing_date,
                errors="raise",
            )

            if parsed_filing_date.year < 2000:
                add_issue(
                    issues,
                    "Warning",
                    "Candidates",
                    row.get(
                        "Candidate_ID",
                        "",
                    ),
                    "Filing_Date",
                    "Unusual Filing Date",
                    (
                        "Filing date appears "
                        "unusually old: "
                        f"{filing_date}"
                    ),
                )

        except Exception:
            add_issue(
                issues,
                "Error",
                "Candidates",
                row.get(
                    "Candidate_ID",
                    "",
                ),
                "Filing_Date",
                "Invalid Date",
                (
                    "Filing_Date could not "
                    f"be parsed: {filing_date}"
                ),
            )


# ============================================================
# CHECK SOURCE QUALITY
# ============================================================

if not sources.empty:
    for _, row in sources.iterrows():
        source_id = row.get(
            "Source_ID",
            "",
        )

        source_type = row.get(
            "Source_Type"
        )

        reliability = row.get(
            "Reliability_Level"
        )

        verified = row.get(
            "Verified"
        )

        link_status = row.get(
            "Link_Status"
        )

        if is_missing(source_type):
            add_issue(
                issues,
                "Error",
                "Sources",
                source_id,
                "Source_Type",
                "Missing Source Classification",
                "Source_Type is blank.",
            )

        if is_missing(reliability):
            add_issue(
                issues,
                "Warning",
                "Sources",
                source_id,
                "Reliability_Level",
                "Missing Reliability Rating",
                "Reliability_Level is blank.",
            )

        if normalize_text(
            verified
        ) != "yes":
            add_issue(
                issues,
                "Warning",
                "Sources",
                source_id,
                "Verified",
                "Source Not Fully Verified",
                (
                    "Source verification "
                    f"status is: {verified}"
                ),
            )

        if normalize_text(
            link_status
        ) in {
            "broken",
            "blocked",
            "not checked",
        }:
            add_issue(
                issues,
                "Warning",
                "Sources",
                source_id,
                "Link_Status",
                "Source Link Issue",
                (
                    "Source link status is: "
                    f"{link_status}"
                ),
            )


# ============================================================
# CHECK FACT VERIFICATION
# ============================================================

if not facts.empty:
    for _, row in facts.iterrows():
        fact_id = row.get(
            "Fact_ID",
            "",
        )

        status = normalize_text(
            row.get(
                "Verification_Status"
            )
        )

        inferred = normalize_text(
            row.get("Is_Inferred")
        )

        conflict = normalize_text(
            row.get(
                "Conflicting_Source_Flag"
            )
        )

        missing_flag = normalize_text(
            row.get(
                "Missing_Value_Flag"
            )
        )

        if status in {
            "unverified",
            "conflicting",
            "not found",
        }:
            add_issue(
                issues,
                "Warning",
                "CandidateFacts",
                fact_id,
                "Verification_Status",
                "Low-Confidence Fact",
                (
                    "Fact requires review "
                    "because its verification "
                    "status is "
                    f"'{row.get('Verification_Status')}'."
                ),
            )

        if inferred == "yes":
            add_issue(
                issues,
                "Information",
                "CandidateFacts",
                fact_id,
                "Is_Inferred",
                "Inferred Fact",
                (
                    "Fact was derived or inferred "
                    "rather than directly stated."
                ),
            )

        if conflict == "yes":
            add_issue(
                issues,
                "Error",
                "CandidateFacts",
                fact_id,
                "Conflicting_Source_Flag",
                "Conflicting Sources",
                (
                    "Fact has conflicting "
                    "source information."
                ),
            )

        if missing_flag == "yes":
            add_issue(
                issues,
                "Information",
                "CandidateFacts",
                fact_id,
                "Missing_Value_Flag",
                "Documented Missing Value",
                (
                    "The record documents that "
                    "this candidate fact was "
                    "not found."
                ),
            )


# ============================================================
# CHECK CANDIDATE STATEMENT REVIEW STATUS
# ============================================================

if not statements.empty:
    for _, row in statements.iterrows():
        statement_id = row.get(
            "Statement_ID",
            "",
        )

        verification = normalize_text(
            row.get("Verified")
        )

        review_status = normalize_text(
            row.get("Review_Status")
        )

        method = normalize_text(
            row.get(
                "Manual_or_Generated"
            )
        )

        if verification in {
            "",
            "no",
        }:
            add_issue(
                issues,
                "Warning",
                "CandidateStatements",
                statement_id,
                "Verified",
                "Statement Not Verified",
                (
                    "Statement has not been "
                    "verified against its source."
                ),
            )

        if review_status in {
            "",
            "not reviewed",
            "needs review",
            "rejected",
            "needs update",
        }:
            add_issue(
                issues,
                "Warning",
                "CandidateStatements",
                statement_id,
                "Review_Status",
                "Statement Review Needed",
                (
                    "Statement review status is: "
                    f"{row.get('Review_Status')}"
                ),
            )

        if method == "ai-assisted":
            add_issue(
                issues,
                "Information",
                "CandidateStatements",
                statement_id,
                "Manual_or_Generated",
                "AI-Assisted Statement",
                (
                    "Statement was created with "
                    "AI assistance and should "
                    "retain human review."
                ),
            )


# ============================================================
# CANDIDATE COMPLETENESS SCORES
# ============================================================

quality_rows = []

for _, candidate in candidates.iterrows():
    candidate_id = candidate.get(
        "Candidate_ID",
        "",
    )

    ballot_name = candidate.get(
        "Ballot_Name",
        "",
    )

    required_present = sum(
        not is_missing(
            candidate.get(field)
        )
        for field in REQUIRED_CANDIDATE_FIELDS
        if field in candidates.columns
    )

    important_present = sum(
        not is_missing(
            candidate.get(field)
        )
        for field in IMPORTANT_CANDIDATE_FIELDS
        if field in candidates.columns
    )

    required_total = len(
        REQUIRED_CANDIDATE_FIELDS
    )

    important_total = len(
        IMPORTANT_CANDIDATE_FIELDS
    )

    required_score = calculate_percentage(
        required_present,
        required_total,
    )

    important_score = calculate_percentage(
        important_present,
        important_total,
    )

    overall_score = round(
        required_score * 0.60
        + important_score * 0.40,
        2,
    )

    candidate_source_count = (
        count_candidate_records(
            sources,
            candidate_id,
        )
    )

    candidate_fact_count = (
        count_candidate_records(
            facts,
            candidate_id,
        )
    )

    candidate_statement_count = (
        count_candidate_records(
            statements,
            candidate_id,
        )
    )

    candidate_policy_count = (
        count_candidate_records(
            policy_topics,
            candidate_id,
        )
    )

    candidate_social_count = (
        count_candidate_records(
            social_media,
            candidate_id,
        )
    )

    if overall_score >= 90:
        quality_level = "Excellent"

    elif overall_score >= 75:
        quality_level = "Good"

    elif overall_score >= 60:
        quality_level = "Fair"

    else:
        quality_level = (
            "Needs Improvement"
        )

    quality_rows.append(
        {
            "Candidate_ID": candidate_id,
            "Ballot_Name": ballot_name,
            "Required_Fields_Present": (
                required_present
            ),
            "Required_Fields_Total": (
                required_total
            ),
            "Required_Completeness_Percent": (
                required_score
            ),
            "Important_Fields_Present": (
                important_present
            ),
            "Important_Fields_Total": (
                important_total
            ),
            "Important_Completeness_Percent": (
                important_score
            ),
            "Overall_Quality_Score": (
                overall_score
            ),
            "Quality_Level": quality_level,
            "Source_Count": (
                candidate_source_count
            ),
            "Fact_Count": (
                candidate_fact_count
            ),
            "Statement_Count": (
                candidate_statement_count
            ),
            "Policy_Record_Count": (
                candidate_policy_count
            ),
            "Social_Media_Count": (
                candidate_social_count
            ),
        }
    )

candidate_quality = pd.DataFrame(
    quality_rows
)


# ============================================================
# TABLE STATISTICS
# ============================================================

table_statistics_rows = []

for table_name, dataframe in tables.items():
    row_count = len(dataframe)
    column_count = len(
        dataframe.columns
    )

    total_cells = (
        row_count * column_count
    )

    if total_cells == 0:
        missing_cells = 0
        completeness = 0.0

    else:
        missing_cells = int(
            dataframe.apply(
                lambda column: column.apply(
                    is_missing
                )
            )
            .sum()
            .sum()
        )

        completeness = (
            calculate_percentage(
                total_cells
                - missing_cells,
                total_cells,
            )
        )

    table_statistics_rows.append(
        {
            "Table": table_name,
            "Rows": row_count,
            "Columns": column_count,
            "Total_Cells": total_cells,
            "Missing_Cells": (
                missing_cells
            ),
            "Completeness_Percent": (
                completeness
            ),
        }
    )

table_statistics = pd.DataFrame(
    table_statistics_rows
)


# ============================================================
# VALIDATION ISSUE TABLE
# ============================================================

validation_issues = pd.DataFrame(
    issues
)

expected_issue_columns = [
    "Severity",
    "Table",
    "Record_ID",
    "Field",
    "Issue_Type",
    "Message",
]

if validation_issues.empty:
    validation_issues = pd.DataFrame(
        columns=expected_issue_columns
    )

else:
    validation_issues = (
        validation_issues[
            expected_issue_columns
        ]
    )


# ============================================================
# VALIDATION SUMMARY
# ============================================================

severity_counts = (
    validation_issues[
        "Severity"
    ]
    .value_counts()
    .to_dict()
)

critical_count = severity_counts.get(
    "Critical",
    0,
)

error_count = severity_counts.get(
    "Error",
    0,
)

warning_count = severity_counts.get(
    "Warning",
    0,
)

information_count = severity_counts.get(
    "Information",
    0,
)

if critical_count > 0:
    validation_status = "Failed"

elif error_count > 0:
    validation_status = (
        "Passed with Errors"
    )

elif warning_count > 0:
    validation_status = (
        "Passed with Warnings"
    )

else:
    validation_status = "Passed"

if candidate_quality.empty:
    average_quality_score = 0.0

else:
    average_quality_score = round(
        candidate_quality[
            "Overall_Quality_Score"
        ].mean(),
        2,
    )

validation_summary = pd.DataFrame(
    [
        {
            "Metric": "Validation Status",
            "Value": validation_status,
        },
        {
            "Metric": "Candidates",
            "Value": len(candidates),
        },
        {
            "Metric": "Sources",
            "Value": len(sources),
        },
        {
            "Metric": "Candidate Facts",
            "Value": len(facts),
        },
        {
            "Metric": "Candidate Statements",
            "Value": len(statements),
        },
        {
            "Metric": "Policy Records",
            "Value": len(policy_topics),
        },
        {
            "Metric": "Social Media Records",
            "Value": len(social_media),
        },
        {
            "Metric": "Critical Issues",
            "Value": critical_count,
        },
        {
            "Metric": "Errors",
            "Value": error_count,
        },
        {
            "Metric": "Warnings",
            "Value": warning_count,
        },
        {
            "Metric": "Informational Issues",
            "Value": information_count,
        },
        {
            "Metric": (
                "Average Candidate Quality Score"
            ),
            "Value": average_quality_score,
        },
        {
            "Metric": "Validation Run Date",
            "Value": datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
        },
    ]
)


# ============================================================
# SAVE VALIDATION REPORT
# ============================================================

OUTPUT_FILE.parent.mkdir(
    parents=True,
    exist_ok=True,
)

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl",
) as writer:

    validation_summary.to_excel(
        writer,
        sheet_name="ValidationSummary",
        index=False,
    )

    candidate_quality.to_excel(
        writer,
        sheet_name="CandidateQuality",
        index=False,
    )

    validation_issues.to_excel(
        writer,
        sheet_name="ValidationIssues",
        index=False,
    )

    table_statistics.to_excel(
        writer,
        sheet_name="TableStatistics",
        index=False,
    )

    output_workbook = writer.book

    for worksheet in (
        output_workbook.worksheets
    ):
        worksheet.freeze_panes = "A2"

        if worksheet.max_row >= 1:
            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

        for column_cells in (
            worksheet.columns
        ):
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
                60,
            )


# ============================================================
# TERMINAL OUTPUT
# ============================================================

print("=" * 72)
print("PILOT DATABASE VALIDATION COMPLETE")
print("=" * 72)

print(f"Input file:\n{INPUT_FILE}")
print()

print(
    f"Validation report:\n{OUTPUT_FILE}"
)
print()

print(
    f"Validation status: "
    f"{validation_status}"
)

print(
    f"Candidates checked: "
    f"{len(candidates)}"
)

print(
    f"Critical issues: "
    f"{critical_count}"
)

print(
    f"Errors: {error_count}"
)

print(
    f"Warnings: {warning_count}"
)

print(
    "Informational issues: "
    f"{information_count}"
)

print(
    "Average candidate quality score: "
    f"{average_quality_score}%"
)

print()

print(
    "Next step: open the validation report "
    "and review ValidationSummary, "
    "CandidateQuality, and ValidationIssues."
)