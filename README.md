# API2

API2 is a Python-based API application that generates a downloadable Excel report with monthly attendance data. The application integrates with Supabase to retrieve data and uses Flask as the web framework along with openpyxl for Excel file creation. It is ready for deployment on platforms like Heroku.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Deployment](#deployment)
- [Contributing](#contributing)
- [License](#license)

## Overview

The API application provides an endpoint `/download_excel` that:
- Connects to a Supabase backend using environment-defined credentials.
- Executes a stored procedure (`get_monthly_attendance_counts`) to retrieve monthly attendance counts.
- Dynamically generates an Excel file with formatted headers and data rows.
- Sends the generated file as a download to the client.

## Features

- **Data Retrieval:** Uses Supabase RPC to fetch monthly attendance counts.
- **Excel Report Generation:** Dynamically creates a well-formatted Excel file using openpyxl.
- **Easy Deployment:** Includes a `Procfile` for seamless deployment on Heroku.
- **Error Handling:** Returns JSON error messages with appropriate HTTP status in case of failures.

## Installation

### Prerequisites

- Python 3.7 or later
- [Pip](https://pip.pypa.io/en/stable/installation/)

### Clone the Repository

```bash
git clone https://github.com/mohamed-ben-kaddour/API2.git
cd API2
