# Outlook Credit Card Sanitizer Add-in

This Outlook add-in detects and masks credit card numbers in incoming emails, displaying only the last four digits to protect sensitive information.

The functionality is useful for banks and financial institutions that often receive credit card PAN numbers over email from customers, yet are subject to PCI regulations where the auditor will look for unmasked card numbers in various systems including email.

## Features
- Detects common credit card number formats (Visa, MasterCard, Amex, etc.).
- Validates numbers using the Luhn algorithm to avoid false positives.
- Masks all but the last four digits (e.g., `1234-5678-9012-3456` becomes `XXXX-XXXX-XXXX-3456`).
- Processes emails in real-time as they are viewed in Outlook.

## Prerequisites
- Node.js and npm installed for local development.
- A web server to host the add-in files (e.g., localhost:3000).
- Outlook client (Desktop or Web) supporting add-ins.

## Installation
1. Clone this repository:
   ```bash
   git clone https://github.com/samiptoivonen/credit-card-sanitizer.git
   ```
2. Navigate to the project directory and install dependencies:
   ```bash
   cd credit-card-sanitizer
   npm install
   ```
3. Start a local web server:
   ```bash
   npm start
   ```
   Ensure the server runs on `https://localhost:3000` with a valid SSL certificate (self-signed is acceptable for development).

4. Sideload the add-in in Outlook:
   - In Outlook, go to **File > Manage Add-ins** (or equivalent in your client).
   - Choose **Add from File** and select the `manifest.xml` file.
   - Follow prompts to install the add-in.

## Usage
- The add-in automatically activates when viewing an email in Outlook.
- It scans the email body for credit card numbers and masks them before display.
- No user interaction is required; the process is transparent.

## Development Notes
- The add-in uses the Office.js API to interact with Outlook.
- The `main.js` script handles email body retrieval, sanitization, and updating.
- The regular expression in `sanitizeCreditCardNumbers` matches 13-16 digit numbers, covering most credit card formats.
- The Luhn algorithm ensures only valid credit card numbers are masked.
- Replace `https://localhost:3000` in `manifest.xml` with your production server URL before deployment.

## Deployment
- Host the `index.html` and `main.js` files on a secure server (HTTPS required).
- Update the `manifest.xml` with your server's URLs for `SourceLocation`, `IconUrl`, and `HighResolutionIconUrl`.
- Distribute the manifest file to users or publish it to the Microsoft 365 Admin Center for organization-wide deployment.

## Contributing
Contributions are welcome! Please submit pull requests or open issues on GitHub.

## License
MIT License. See `LICENSE` for details.
