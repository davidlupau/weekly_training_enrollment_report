# Weekly Training Enrollment Report Generator

A Python-based automation solution that transforms messy leadership program enrollment data into clean, actionable reports for facilitators and HR partners.

## Overview
This project automates the weekly reporting process for the Essentials of Leadership (EoL) program, a global leadership development initiative. It processes raw enrollment data and generates standardized reports, saving approximately 1 hour of manual work per week.

## Problem Statement
The EoL program involves:
- 1,000 leaders participating annually
- 30 facilitators delivering content globally
- 250 sessions scheduled each year

The manual reporting process required:
1. Extracting raw data from the LMS (sample in `input_data.xlsx`)
2. Cleaning and standardizing the data
3. Resolving duplicate enrollments
4. Creating facilitator and HR reports
5. Tracking manager communications

This process was time-consuming, error-prone, and required specialized knowledge.

## Solution
This automation pipeline processes raw enrollment data and generates three outputs:
1. Detailed facilitator report with participant information
2. Session summary report for HR partners
3. Manager communication tracking list

## Processing Flow

```mermaid
flowchart TD
    A[User Uploads the File]
    B[Delete Row 1]
    C[Delete Specific Useless Columns]
    D[Rename Columns to Python Standard (_ instead of spaces)]
    E[Convert employee_id to Numbers]
    F[Format Offering_Start_Date (e.g. Jan-11, remove time)]
    G[Convert "Enrolled - Pending Approval" to "Pending Approval"]
    H[Filter rows: Keep only Enrolled or Pending Approval]
    I[Check for duplicate enrollments (same module)]
    I1[Same Attendance Status? Keep latest date]
    I2[Different Status? Delete 'Did Not Attend']
    J[Check for duplicates again]
    K[Create file for facilitators (1 row per person)]
    L[Save file: EoL_enrollment_report_YYYYMMDD]
    M[Save file with session list]
    N[Check/update file: Manager email delivery status]

    A --> B --> C --> D --> E --> F --> G --> H --> I
    I --> I1 --> I2 --> J --> K --> L --> M --> N
```

## Repository Structure
- `main.py`: Main execution script with the processing flow
- `functions.py`: Utility functions for data processing
- `input_data.xlsx`: Sample input data (anonymized)
- `requirements.txt`: Required Python packages
- `LICENSE`: License information

## Technical Highlights
- **Date format standardization**: Handles multiple international date formats
- **Intelligent deduplication**: Applies business logic to resolve duplicate enrollments
- **Excel output formatting**: Creates properly formatted reports for business users
- **Persistent tracking**: Maintains manager communication history

## Installation and Usage
### Prerequisites
- Python 3.8+
- Packages listed in `requirements.txt`

### Setup
```bash
# Clone the repository
git clone https://github.com/davidlupau/weekly_training_enrollment_report
cd weekly_training_enrollment_report

# Install dependencies
pip install -r requirements.txt
```

### Running the Report Generator
```bash
python main.py path/to/input_data.xlsx
```

## Business Benefits
- 52+ hours saved annually on manual processing
- Improved data quality and consistency
- Enhanced facilitator experience
- Operational resilience through automation
- Scalable solution adaptable to other programs

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

## License
MIT
This project is licensed under the MIT License - see the LICENSE file for details.
