## AppTrackPro

### Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Technical Stack](#technical-stack)
- [Project Structure](#project-structure)
  - [Directory Breakdown](#directory-breakdown)
  - [File Descriptions](#file-descriptions)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)

### Overview

**AppTrackPro** is a comprehensive job application tracking tool designed to help users manage and monitor their job applications efficiently. Built with Python's Tkinter library, AppTrackPro offers a user-friendly graphical interface, seamless integration with Google Sheets for data synchronization, and robust configuration management to ensure personalized user experiences.

### Features

- **Add Applications:** Easily input and store details of job applications, including company name, position, application portal URL, date applied, and status.
- **View & Edit Applications:** Interactive Treeview to display all applications with options to search, edit, delete, and update statuses.
- **Google Sheets Integration:** Synchronize your application data with Google Sheets for cloud-based access and backup.
- **Personal Information Management:** Store and manage personal details securely, with easy clipboard access for quick data entry.
- **Theming:** Switch between Dark and Light modes to suit your visual preferences.
- **Drag-and-Drop Functionality:** Upload essential files effortlessly using drag-and-drop features.
- **Configuration Management:** Comprehensive settings dialog to manage Google Sync, file paths, and other configurations.

### Technical Stack

- **Programming Language:** Python 3.x
- **Libraries & Frameworks:**
  - `Tkinter` for GUI development
  - `pandas` for data manipulation
  - `openpyxl` for Excel file operations
  - `google-auth` and `google-api-python-client` for Google Sheets integration
  - `Pillow` for image processing
  - `windnd` for drag-and-drop functionality
  - `appdirs` for managing user-specific application data directories
- **File Formats:** Excel (`.xlsx`), JSON (`.json`), ICO (`.ico`)

### Project Structure

```
AppTrack
├── assets/
│   ├── app_icon.png
│   ├── applications.png
│   ├── google_sync.png
│   ├── upload_json.png
│   ├── upload_sheets_id.png
│   └── upload_xlsx.png
├── config/
│   └── settings_manager.py
├── data/
├── src/
│   ├── gui/
│   │   └── main_window.py
│   └── utils/
│       ├── utils.py
│       ├── file_io.py
│       └── google_sheets.py
├── app.py
├── hook-tkinterdnd2.py
└── requirements.txt
```

#### Directory Breakdown

- **assets/**: Contains all static assets like icons and images used in the application.
- **config/**: Houses configuration management scripts, ensuring settings are loaded and saved correctly.
- **data/**: Stores data files such as the `Applications.xlsx` which holds all job application records.
- **src/**: Source code divided into:
  - **gui/**: GUI-related scripts, primarily the main application window.
  - **utils/**: Utility scripts for file I/O and Google Sheets operations.
- **app.py**: The entry point of the application, initializing logging and launching the main window.

#### File Descriptions

- **app.py**: Initializes the application, sets up logging, configures environment variables, and launches the Tkinter main loop.

- **config/settings_manager.py**:
  - Manages application configurations, ensuring settings are stored in user-specific directories.
  - Handles themes, Google Sync settings, file paths, and initializes default configurations if none exist.
  - Utilizes the `appdirs` library to determine appropriate directories for storing configuration and data files.

- **src/gui/main_window.py**:
  - Defines the `AppTrackPro` class, inheriting from `tk.Tk`, which sets up the main application window.
  - Implements the user interface, including tabs for adding applications, viewing/editing applications, and managing personal information.
  - Handles interactions such as adding new applications, editing existing ones, syncing with Google Sheets, and theming.

- **src/utils/file_io.py**:
  - Provides functions to read from and write to the `Applications.xlsx` file using `pandas` and `openpyxl`.
  - Ensures data consistency and handles cases where the Excel file might be missing or corrupted.

- **src/utils/google_sheets.py**:
  - Manages synchronization between the local `Applications.xlsx` and Google Sheets.
  - Includes functions to read data from Google Sheets, write data to Google Sheets, and delete specific rows.
  - Utilizes Google’s APIs for authentication and data manipulation.

- **assets/**:
  - **app_icon.ico**: The main application icon.
  - **upload_json.png**, **upload_sheets_id.png**, **upload_xlsx.png**: Icons used in the settings dialog for uploading respective files.

### Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/Maged-Abuzaid/AppTrackPro.git
   cd AppTrackPro
   ```

2. **Create a Virtual Environment (Optional but Recommended)**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Set Up Configuration**
   - Ensure that the `config/` and `data/` directories are properly initialized. The application will create necessary files on the first run.

### Usage

1. **Run the Application**
   ```bash
   python app.py
   ```

2. **Adding a Job Application**
   - Navigate to the "Add Application" tab.
   - Fill in the Company, Position, and Application Portal URL fields.
   - Click "Submit" to save the application.

3. **Viewing and Editing Applications**
   - Switch to the "View/Edit Applications" tab to see all your applications.
   - Use the search bar to filter applications.
   - Right-click on any row to delete or copy it.
   - Double-click on cells to edit their contents.

4. **Google Sheets Synchronization**
   - Navigate to Settings (⚙️) in the menu bar.
   - Configure Google Sync by providing the necessary API credentials and Spreadsheet ID.
   - Enable Google Sync to automatically synchronize your data.

### Configuration

**Google Sync Setup:**

1. **Create a Google API Project**
   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Create a new project.
   - Enable the Google Sheets API for your project.

2. **Create Service Account Credentials**
   - In the Google Cloud Console, navigate to **APIs & Services > Credentials**.
   - Click on **Create Credentials > Service Account**.
   - Follow the prompts to create a service account.
   - Once created, generate a JSON key and download it.

3. **Share Your Spreadsheet**
   - Create a new Google Sheet or use an existing one.
   - Share the sheet with the service account's email address (found in the JSON key file).

4. **Configure AppTrackPro**
   - Open AppTrackPro and go to Settings (⚙️).
   - Upload the `service_account.json` file.
   - Enter your Spreadsheet ID.
   - (Optional) Upload an existing `Applications.xlsx` with the required columns: Company, Position, Application Portal URL, Date Applied, and Status.
   - Save the settings and restart the application.

### Contributing

Contributions are welcome! Please follow these steps:

1. **Fork the Repository**

2. **Create a Feature Branch**
   ```bash
   git checkout -b feature/YourFeature
   ```

3. **Commit Your Changes**
   ```bash
   git commit -m "Add Your Feature"
   ```

4. **Push to the Branch**
   ```bash
   git push origin feature/YourFeature
   ```

5. **Open a Pull Request**

### License

Distributed under the MIT License. See `LICENSE` for more information.

### Contact

- **Maged Abuzaid** - [LinkedIn](https://www.linkedin.com/in/maged-abuzaid/) - MagedM.Abuzaid@gmail.com
- **Project Link:** [https://github.com/yourusername/AppTrackPro](https://github.com/yourusername/AppTrackPro)

---

# Executable README.md

# AppTrackPro

![AppTrackPro Logo](assets/app_icon.ico)

## Introduction

**AppTrackPro** is a user-friendly application designed to help you manage and track your job applications efficiently. With a sleek interface and powerful features, AppTrackPro ensures you stay organized and never miss an opportunity.

## Features

- **Add Applications:** Easily input details of your job applications.
- **View & Edit:** Manage your applications with options to search, edit, and delete entries.
- **Google Sync:** Seamlessly synchronize your application data with Google Sheets.
- **Personal Info Management:** Store and access your personal information securely.
- **Theming:** Choose between Dark and Light modes to suit your preference.
- **Drag-and-Drop:** Effortlessly upload essential files using drag-and-drop functionality.

## Getting Started

### Prerequisites

- **Operating System:** Windows 10 or later
- **Python:** Not required (executables are bundled)
- **Internet Connection:** Required for Google Sync

### Downloading the Executable

1. **Download the .exe File**
   - Visit the [AppTrackPro Releases](https://github.com/yourusername/AppTrackPro/releases) page on GitHub.
   - Download the latest `AppTrackPro.exe` file.

2. **Run the Application**
   - Double-click the downloaded `AppTrackPro.exe` file.
   - The application will launch, displaying the main window.

### Setting Up and Configuring Google Sync

To enable Google Sync and synchronize your application data with Google Sheets, follow these steps:

#### 1. Setting Up a Google API

1. **Create a Google Cloud Project**
   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Click on **Select a project** and then **New Project**.
   - Enter a project name and click **Create**.

2. **Enable Google Sheets API**
   - In the Google Cloud Console, navigate to **APIs & Services > Library**.
   - Search for "Google Sheets API" and click **Enable**.

3. **Create Service Account Credentials**
   - Go to **APIs & Services > Credentials**.
   - Click on **Create Credentials > Service Account**.
   - Provide a name and description for the service account.
   - Click **Create and Continue**, then **Done**.
   - Once the service account is created, click on it to open its details.
   - Navigate to the **Keys** tab and click **Add Key > Create New Key**.
   - Select **JSON** and click **Create** to download the `service_account.json` file.

4. **Share Your Google Sheet with the Service Account**
   - Open your Google Sheet or create a new one.
   - Click on **Share** and enter the service account's email address (found in the `service_account.json` file).
   - Assign **Editor** permissions and click **Send**.

#### 2. Uploading Required Files

1. **Service Account JSON File**
   - In AppTrackPro, navigate to **Settings** (⚙️ icon in the menu bar).
   - Click on **Google Sync Configuration**.
   - Upload the `service_account.json` file you downloaded earlier.

2. **Google Sheets ID**
   - Open your Google Sheet and copy the Spreadsheet ID from the URL.
     - Example URL: `https://docs.google.com/spreadsheets/d/your_spreadsheet_id/edit#gid=0`
     - The part between `/d/` and `/edit` is your Spreadsheet ID.
   - In AppTrackPro's **Google Sync Configuration**, enter the Spreadsheet ID.

#### 3. (Optional) Uploading an Existing Applications.xlsx

If you have an existing `Applications.xlsx` file with the following columns:
- **Company**
- **Position**
- **Application Portal URL**
- **Date Applied**
- **Status**

You can upload it to AppTrackPro:

1. **Navigate to Settings**
   - Click on the **Settings** (⚙️) icon in the menu bar.
   - Go to **Google Sync Configuration**.

2. **Upload Applications.xlsx**
   - Click on the designated area to upload the `Applications.xlsx` file.
   - Alternatively, drag and drop the file into the upload zone.

### Using AppTrackPro

1. **Adding a New Application**
   - Open AppTrackPro.
   - Navigate to the **Add Application** tab.
   - Fill in the **Company**, **Position**, and **Application Portal URL** fields.
   - Click **Submit** to save the application.

2. **Viewing and Editing Applications**
   - Switch to the **View/Edit Applications** tab.
   - Use the search bar to filter applications.
   - Right-click on any application to delete or copy it.
   - Double-click on cells to edit their contents directly.

3. **Managing Personal Information**
   - Go to the **Clipboard** tab.
   - Edit your personal information as needed.
   - Click **Save** to store changes.

4. **Switching Themes**
   - Click on the **Settings** (⚙️) icon in the menu bar.
   - Select **Switch Theme** to toggle between Dark and Light modes.

### Troubleshooting

- **Google Sync Issues:**
  - Ensure that the `service_account.json` file is correctly uploaded.
  - Verify that the Spreadsheet ID is accurate.
  - Make sure the service account has Editor access to the Google Sheet.
  - Check your internet connection.

- **Application Not Launching:**
  - Ensure that you have downloaded the correct executable for your operating system.
  - Disable any antivirus software that might be blocking the application.
  - Re-download the executable in case the file was corrupted during download.

- **Data Not Syncing:**
  - Confirm that Google Sync is enabled in Settings.
  - Check if there are any errors displayed in the application logs.

### Support

For further assistance, please contact [your.email@example.com](mailto:your.email@example.com) or visit the [AppTrackPro GitHub Issues](https://github.com/yourusername/AppTrackPro/issues) page.

---