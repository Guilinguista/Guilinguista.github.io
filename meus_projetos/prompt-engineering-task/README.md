# Informative Article Automation

## Overview
This script automates the creation of an informative article based on the content of a provided website. The article is generated using Cohere's API, saved as a PowerPoint presentation, and sent via email.

## How to Use

### Prerequisites
- Python 3.10 or higher.
- Dependencies listed in `requirements.txt`.
- A valid Cohere API key set as an environment variable (`COHERE_API_KEY`).
- Email settings configured as environment variables (`EMAIL_ADDRESS` and `EMAIL_PASSWORD`).

### Setup
1. Clone the repository and activate your virtual environment.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Configure your Cohere API key:
   ```bash
   export COHERE_API_KEY="your_cohere_api_key"
   ```
4. Configure your email credentials:
   ```bash
   export EMAIL_ADDRESS="your_email@example.com"
   export EMAIL_PASSWORD="your_email_password"
   ```

### Execution
Run the script by providing the website and email file:
```bash
python main_content.py --website "https://example.com" --emails emails.txt
```

### Output
- A `summary.pptx` file will be generated containing the informative article.
- The article will be sent via email to the specified recipients.