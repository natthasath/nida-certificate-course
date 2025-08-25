# 🎉 NIDA Certificate Course

NIDA Certificate Course is a project for managing training assessments that automatically generates personalized course completion certificates, provides instant download, and sends them directly to participants via email.

![version](https://img.shields.io/badge/version-1.0-blue)
![rating](https://img.shields.io/badge/rating-★★★★★-yellow)
![uptime](https://img.shields.io/badge/uptime-100%25-brightgreen)

### ✅ Requirements

- [Canva](https://www.canva.com/)
- [Github Pages](https://docs.github.com/en/pages)
- [Google Apps Script](https://script.google.com/)

### 🚀 Prod Setup

- Change path certificate `../templates/certificate-00.svg` in html file
- Config `GOOGLE_APPS_SCRIPT_URL` in html file
- Deploy code on Github Pages
- Config `SPREADSHEET_ID` in gs file
- Config `CERTIFICATE_FOLDER_ID` in gs file
- Deploy code on Google Apps Script

### 🏆 Run

- Run test function `setupHeaders` in gs file
- Run test function `testDoPostWithCertificateAndEmail` in gs file