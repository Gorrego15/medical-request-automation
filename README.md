# Medical Request Automation System

This Python-based automation tool streamlines the process of requesting Medical Records and Medical Bills (MR/MB) for legal case management. It automates the generation of customized request letters and the editing of signed HIPAA authorizations.

## Features

- **Automated PDF Editing**: Uses `PyMuPDF` to inject realistic, randomized audit trails and specific service dates into signed HIPAA documents.
- **Dynamic Word Templates**: Automatically updates Microsoft Word request letters with provider names, addresses, and client details.
- **Batch Processing**: Reads provider lists from Excel to generate dozens of personalized request packages in seconds.
- **Seamless Merging**: Combines the request letter and the specific HIPAA authorization into a single, professional PDF ready for mailing or faxing.

## Prerequisites

- **Python 3.9+**
- **Microsoft Word**: Required for the `docx2pdf` conversion module (compatible with Windows and macOS).

## Installation

1. Clone the repository:
   ```bash
   git clone [https://github.com/Gorrego15/medical-request-automation.git](https://github.com/Gorrego15/medical-request-automation.git)