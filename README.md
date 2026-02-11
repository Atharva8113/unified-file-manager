# Unified File Manager System

A powerful automation tool that streamlines your workflow by automatically organizing OOC uploads, moving job folders to billing, and sorting loose files.

## Quick Start

*   **Requirements:** Windows OS (No Python installation needed).
*   **Installation:** Download the folder and locate the `.exe` file (likely named `UnifiedFileManager.exe` or `unified_file_managerv2.exe`).

## How to Use

1.  **Launch the App:** Double-click the executable file to start the Unified File Manager.
2.  **Dashboard Overview:** You will see a dashboard monitoring three main services:
    *   **OOC Upload:** Moves files from the upload folder to their specific job folders.
    *   **Job Mover:** Detects "Out of Charge" triggers and moves job folders to the billing directory.
    *   **Loose Files:** Scans for loose files in billing folders and places them into their matching job subfolders.
3.  **Start Automation:** Click the green **"▶ Start All Services"** button to begin all tasks simultaneously. You can also start each service individually using the "Start" buttons next to their names.
4.  **Monitor Progress:** Watch the "Live Statistics" cards to see real-time counts of moved files and any errors encountered.
5.  **View Logs:** In the "Logs & Management" section, click the buttons (e.g., "📄 Job Move Log") to view detailed records of all actions taken by the system.
6.  **Undo Mistakes:** If a job folder was moved incorrectly, click **"🔄 Open Revert Manager"**, select the job, and click "Revert" to move it back to its original location.
7.  **Stop Automation:** Click the red **"⏹ Stop All Services"** button when you are finished.

## Common Issues

*   **Windows Defender Warning:** If Windows Defender prevents the app from starting, click **'More Info'** -> **'Run Anyway'**. This happens because it is a custom internal tool not signed by Microsoft.
*   **Network Access:** Ensure you are connected to the office network and have access to the **Y:** and **Z:** drives. The application will show errors if it cannot reach these folders.

## Contact

For support, feature requests, or to update folder mappings, please contact the IT Team.
