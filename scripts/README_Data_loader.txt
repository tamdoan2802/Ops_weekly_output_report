================================================================================
  README — Data_loader (MyDaily Automator)
================================================================================

This folder contains scripts to automate downloading the Workload Excel file
from https://mydaily.myteamsolution.com.vn/ using Playwright.

Because the website requires a login and specific clicks, we use a two-step 
process to configure the script.

--------------------------------------------------------------------------------
STEP 1: INSTALL PREREQUISITES
--------------------------------------------------------------------------------
Open a terminal in this scripts/ folder and run:
  pip install playwright
  playwright install chromium

--------------------------------------------------------------------------------
STEP 2: RUN THE AUTOMATION
--------------------------------------------------------------------------------
Whenever you need to pull fresh data, just run:
  python Data_loader.py

The script will launch, log in using the embedded credentials, perform the steps, 
and save the Excel file directly into:
  G:\My Drive\Dữ liệu nhân sự\Workload\Construction Team\
