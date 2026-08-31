````markdown
# Approval Category Processor

A Streamlit application that processes an Input Excel file against a Reference Excel file and generates an Output Excel file.

## Input File

The Input Excel file must contain:

- CustomerName
- Comment

## Reference File

The Reference Excel file must contain:

- FunctionName
- Comment1
- Comment2
- Comment3
- etc.

The application automatically detects all columns starting with "Comment".

## Output

The application generates one Excel file containing:

### Output
Matched CustomerName, FunctionName, ApprovalCategory and Comment.

### Not Match
Comments that could not be matched with the Reference file.

### Reference
The original Reference data.

### Criteria
Approval category criteria.

## Run Locally

Install the required packages:

```bash
pip install -r requirements.txt
````

Run the application:

```bash
streamlit run app.py
```

## Streamlit Cloud

Upload the repository to GitHub and deploy the application using Streamlit Community Cloud.

Main file:

```text
app.py
```

```
```
