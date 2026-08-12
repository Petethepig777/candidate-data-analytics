
## Project Overview

This project was developed to transform unstructured candidate information into a structured analytical dataset that supports policy analysis and candidate comparison.

The workflow includes:

- Structured candidate database design
- Data validation and quality assessment
- Policy topic classification using rule-based NLP
- Human review and classifier evaluation
- Candidate feature engineering
- Candidate similarity analysis
- Proof-of-concept supervised machine learning

---

## Project Workflow

```
Public Candidate Information
            │
            ▼
    Data Collection & Cleaning
            │
            ▼
   Structured Candidate Database
            │
            ▼
   Policy Topic Classification
            │
            ▼
 Human Validation & Evaluation
            │
            ▼
 Feature Engineering
            │
            ▼
 Candidate Similarity Analysis
            │
            ▼
 TF-IDF + Logistic Regression
```

---

## Technologies

- Python
- pandas
- NumPy
- openpyxl
- scikit-learn

---

## Analytical Methods

### Data Engineering

- Relational database design
- Data validation
- Duplicate detection
- Missing value handling
- Source verification
- Research completeness scoring

### Exploratory Data Analysis

- Candidate research coverage
- Policy topic frequency
- Source quality analysis
- Candidate comparison

### Natural Language Processing

- Text preprocessing
- Keyword extraction
- Rule-based multi-label policy classification
- Policy taxonomy construction

### Feature Engineering

Candidate-level numerical features including:

- Policy Diversity
- Total Policy Weight
- Dominant Policy
- Average Policy Weight

### Similarity Analysis

- Cosine Similarity
- Jaccard Similarity
- Combined Similarity Score

### Machine Learning (Proof of Concept)

A small supervised machine learning experiment was implemented using:

- TF-IDF text vectorization
- Logistic Regression (One-vs-Rest)
- Leave-One-Out Cross Validation
- Precision
- Recall
- F1 Score
- Hamming Accuracy

The ML component serves as a proof-of-concept due to the limited pilot dataset.

---

## Repository Structure

```
.
├── analyze_candidate_similarity.py
├── build_candidate_comparison.py
├── build_candidate_features.py
├── build_candidate_master.py
├── build_dashboard_dataset.py
├── classify_policy_topics.py
├── create_analytics_report.py
├── create_classifier_review_sample.py
├── create_database_template.py
├── evaluate_classifier.py
├── expand_pilot_statements.py
├── load_pilot_candidates.py
├── train_ml_policy_classifier.py
├── validate_pilot_database.py
└── README.md
```

---

## Data Availability

The original datasets, generated reports, and intermediate outputs are **not included** in this repository.

This repository focuses on the implementation of the analytical workflow rather than distributing the research dataset.

---

## Disclaimer

This project is intended for educational and portfolio purposes. Candidate information analyzed by the pipeline should originate from publicly available sources. Any conclusions produced by the analytical pipeline should not be interpreted as political endorsements or official evaluations.

---

## Future Improvements

- Larger training dataset
- Transformer-based NLP models
- Interactive dashboard
- Automated data collection pipeline
- Temporal policy trend analysis
