// export-usage.js

const { chromium } = require('@playwright/test');
const path = require('path');
const fs = require('fs');

const SITES = [
  'https://contoso.sharepoint.com/sites/Operations',
  'https://contoso.sharepoint.com/sites/ClinicalServices',
  'https://contoso.sharepoint.com/sites/SocialServices',
  'https://contoso.sharepoint.com/sites/CaseManagement',
  'https://contoso.sharepoint.com/sites/LevelOfCare',
  'https://contoso.sharepoint.com/sites/HumanResources',
  'https://contoso.sharepoint.com/sites/Facilities',
  'https://contoso.sharepoint.com/sites/Leadership',
  'https://contoso.sharepoint.com/sites/FoodServices',
  'https://contoso.sharepoint.com/sites/RevenueCycle',
  'https://contoso.sharepoint.com/sites/ResidentialPrograms',
  'https://contoso.sharepoint.com/sites/VocationalServices',
  'https://contoso.sharepoint.com/sites/RecreationalTherapy',
  'https://contoso.sharepoint.com/sites/OccupationalTherapy',
  'https://contoso.sharepoint.com/sites/Nursing',
  'https://contoso.sharepoint.com/sites/Quality',
  'https://contoso.sharepoint.com/sites/DayPrograms',
  'https://contoso.sharepoint.com/sites/CareCoordination',
  'https://contoso.sharepoint.com/sites/CompanyNews',
  'https://contoso.sharepoint.com/sites/SpecialtyClinic',
  'https://contoso.sharepoint.com/sites/BoardOfTrustees',
  'https://contoso.sharepoint.com/sites/RegionalManagement',
  'https://contoso.sharepoint.com/sites/Admissions',
  'https://contoso.sharepoint.com/sites/Administration',
  'https://contoso.sharepoint.com/sites/BoardDocuments',
  'https://contoso.sharepoint.com/sites/BusinessAnalytics',
  'https://contoso.sharepoint.com/sites/Finance',
  'https://contoso.sharepoint.com/sites/StaffScheduling',
  'https://contoso.sharepoint.com/sites/PatientCare'
];

const OUT_DIR = path.resolve('./downloads');
const TIME_RANGE = '90'; // '7', '30', or '90'

function siteSlug(url) {
  const u = new URL(url);
  const seg = u.pathname.replace(/\/sites\//, '').replace(/\/$/, '') || 'root';
  return seg.replace(/[^a-zA-Z0-9._-]/g, '_');
}

(async () => {
  fs.mkdirSync(OUT_DIR, { recursive: true });

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  // One-time interactive login
  await page.goto('https://contoso.sharepoint.com');

  console.log('\n>>> Log in + complete MFA in the browser, then press Enter here...');
  await new Promise(r => process.stdin.once('data', r));

  const results = { ok: [], failed: [] };

  for (const [i, siteUrl] of SITES.entries()) {
    const label = `[${i + 1}/${SITES.length}] ${siteUrl}`;
    console.log(label);

    try {
      const usageUrl = `${siteUrl.replace(/\/$/, '')}/_layouts/15/siteanalytics.aspx?view=19`;

      const resp = await page.goto(usageUrl, {
        waitUntil: 'domcontentloaded',
        timeout: 30000
      });

      if (resp && !resp.ok() && resp.status() !== 200) {
        throw new Error(`HTTP ${resp.status()}`);
      }

      const currentUrl = page.url();

      if (/AccessDenied\.aspx|error\.aspx|signin/i.test(currentUrl)) {
        throw new Error('Access denied or redirected to sign-in');
      }

      const tab = page.getByRole('tab', {
        name: new RegExp(`last ${TIME_RANGE} days`, 'i')
      });

      if (await tab.count()) {
        await tab.first().click().catch(() => {});
        await page.waitForTimeout(750);
      }

      const downloadBtn = page.getByRole('button', {
        name: /download .* site usage/i
      });

      try {
        await downloadBtn.waitFor({
          state: 'visible',
          timeout: 15000
        });
      } catch {
        const bodyText = await page
          .locator('body')
          .innerText()
          .catch(() => '');

        if (/sorry.*access|don't have access|permission/i.test(bodyText)) {
          throw new Error('No access to site usage');
        }

        throw new Error('Download button not found (page may have changed)');
      }

      const downloadPromise = page.waitForEvent('download', {
        timeout: 30000
      });

      await downloadBtn.click();

      const download = await downloadPromise;

      const suggested = download.suggestedFilename();
      const ext = path.extname(suggested) || '.xlsx';

      const outPath = path.join(
        OUT_DIR,
        `${siteSlug(siteUrl)}${ext}`
      );

      await download.saveAs(outPath);

      console.log(`   ✓ saved ${outPath}`);
      results.ok.push(siteUrl);

    } catch (err) {
      const msg = err.message.split('\n')[0].slice(0, 120);

      console.log(`   ✗ skipped: ${msg}`);

      results.failed.push({
        siteUrl,
        error: msg
      });
    }

    await page.waitForTimeout(1000);
  }

  console.log(`\nDone. ${results.ok.length} ok, ${results.failed.length} skipped.`);

  if (results.failed.length) {
    const failPath = path.join(OUT_DIR, '_failed.txt');

    fs.writeFileSync(
      failPath,
      results.failed.map(f => `${f.siteUrl}\t${f.error}`).join('\n')
    );

    console.log(`Skipped sites written to ${failPath}`);

    results.failed.forEach(f =>
      console.log(`  - ${f.siteUrl}: ${f.error}`)
    );
  }

  await browser.close();
})();