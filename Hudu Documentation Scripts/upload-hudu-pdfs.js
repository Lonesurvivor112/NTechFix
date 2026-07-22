const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  /*
    Hudu Bulk PDF Upload Script

    Description:
    This script uploads all PDF files from a local directory into a specific Hudu
    Knowledge Base folder using Playwright browser automation.

    Notes:
    - Browser runs in visible mode because SSO login requires manual interaction.
    - Update the configuration variables below before running.
    - Test with a small folder of PDFs first before running a large batch.
  */

  // =========================
  // Configuration
  // =========================

  // Local directory containing PDF files
  const pdfDirectory = 'X:\\Downloads\\renamed_orders';

  // Base Hudu URL
  const huduBaseUrl = 'https://domain.huducloud.com';

  // Hudu KB folder ID
  const folderId = 30;

  // Manual SSO login wait time in milliseconds
  // 60000 = 60 seconds
  const manualLoginDelay = 60000;

  // Delay after each upload for stability
  const uploadStabilityDelay = 3000;

  // =========================
  // Derived URLs
  // =========================

  const loginUrl = `${huduBaseUrl}/users/sign_in`;
  const targetFolderUrl = `${huduBaseUrl}/kba?folder=${folderId}`;

  // Track the current file being uploaded for better error reporting
  let currentFile = '';

  try {
    // Validate that the PDF directory exists
    if (!fs.existsSync(pdfDirectory)) {
      console.error('PDF directory does not exist:', pdfDirectory);
      return;
    }

    // Read all .pdf files from the directory
    const pdfFiles = fs
      .readdirSync(pdfDirectory)
      .filter(file => file.toLowerCase().endsWith('.pdf'));

    if (pdfFiles.length === 0) {
      console.log('No PDF files found in the directory:', pdfDirectory);
      return;
    }

    console.log(`Found ${pdfFiles.length} PDF files to upload.`);
    console.log(`Source directory: ${pdfDirectory}`);
    console.log(`Target Hudu folder: ${targetFolderUrl}`);

    // Launch browser
    const browser = await chromium.launch({
      headless: false
    });

    const page = await browser.newPage();

    try {
      // Navigate to the Hudu login page
      console.log('Opening Hudu login page...');
      await page.goto(loginUrl);

      // Click the SSO link
      console.log('Clicking Single Sign On link...');
      await page.getByRole('link', { name: 'Use Single Sign On (SSO)' }).click();

      // Wait for manual SSO login
      console.log(`Please complete the SSO login within ${manualLoginDelay / 1000} seconds...`);
      await page.waitForTimeout(manualLoginDelay);

      // Navigate to target KB folder
      console.log('Navigating to target Hudu folder...');
      await page.goto(targetFolderUrl);

      // Loop through each PDF file
      for (const pdfFile of pdfFiles) {
        currentFile = pdfFile;

        const pdfPath = path.join(pdfDirectory, pdfFile);

        console.log('----------------------------------------');
        console.log(`Uploading: ${pdfFile}`);
        console.log(`Full path: ${pdfPath}`);

        // Click "New Article"
        await page.getByText('New Article').click();

        // Click "Upload PDF" link in the dropdown
        await page.getByRole('link', { name: ' Upload PDF' }).click();

        // Wait for the upload modal/dropzone to appear
        await page.waitForSelector('text=Drop files here to upload', {
          state: 'visible'
        });

        // Click the dropzone
        await page.getByText('Drop files here to upload').click();

        // Upload the PDF file
        const fileInput = page.locator('input[type="file"]');
        await fileInput.setInputFiles(pdfPath);

        // Click the "Upload PDF" button to complete the upload
        await page
          .locator('a.button.button--secondary.button--large', {
            hasText: 'Upload PDF'
          })
          .click();

        // Wait briefly for upload/post-processing to complete
        await page.waitForTimeout(uploadStabilityDelay);

        console.log(`Successfully uploaded: ${pdfFile}`);

        // Navigate back to the target folder before processing the next file
        await page.goto(targetFolderUrl);
      }

      console.log('----------------------------------------');
      console.log('All PDF uploads completed successfully.');
    } catch (error) {
      console.error('----------------------------------------');
      console.error(`Error during upload of: ${currentFile || 'unknown file'}`);
      console.error(error);
    } finally {
      await browser.close();
      console.log('Browser closed.');
      console.log('Upload process completed.');
    }
  } catch (error) {
    console.error('Fatal script error:');
    console.error(error);
  }
})();
