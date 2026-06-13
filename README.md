# StatementIQ — App

LP capital account statement generator for US PE & VC funds.
Fund managers log in, enter LP data, and generate professional Excel statements — downloaded instantly and emailed to the LP.

## Repository structure

```
index.html          ← The full web app (frontend)
api/
  generate.py       ← Python serverless function (Excel generator + email)
requirements.txt    ← Python dependencies
vercel.json         ← Vercel routing config
.gitignore
README.md
```

## How it works

1. Fund manager opens the app → enters LP details, NAV, capital calls, distributions
2. Clicks **Generate** → browser sends data to `/api/generate` (Python on Vercel)
3. Python builds a formatted Excel file using `openpyxl`
4. Excel downloads instantly in the browser
5. Python simultaneously emails the file to the LP's email address

## Deploy to Vercel via GitHub

1. Push this repo to GitHub
2. Go to [vercel.com](https://vercel.com) → Add New Project
3. Import this GitHub repository
4. Framework Preset → **Other**
5. Click **Deploy**

Your app will be live at `https://your-project.vercel.app`

## Required environment variables

Set these in Vercel → Project → Settings → Environment Variables:

| Variable    | Description                              | Example                  |
|-------------|------------------------------------------|--------------------------|
| `SMTP_HOST` | Your email provider's SMTP server        | `smtp.gmail.com`         |
| `SMTP_PORT` | SMTP port (usually 587)                  | `587`                    |
| `SMTP_USER` | The email address emails are sent from   | `statements@yourfund.com`|
| `SMTP_PASS` | Email password or App Password           | `your-app-password`      |
| `FROM_NAME` | Display name on outgoing emails          | `Alpine Equity Partners` |

### Gmail setup (recommended for getting started)

1. Go to your Google Account → Security → Turn on 2-Step Verification
2. Search **App Passwords** → Create one for "Mail"
3. Copy the 16-character password → paste as `SMTP_PASS`

### Other providers

| Provider       | SMTP_HOST                  | SMTP_PORT |
|----------------|----------------------------|-----------|
| Gmail          | `smtp.gmail.com`           | 587       |
| Outlook/O365   | `smtp.office365.com`       | 587       |
| Zoho Mail      | `smtp.zoho.com`            | 587       |
| SendGrid       | `smtp.sendgrid.net`        | 587       |

After adding environment variables → Vercel → Deployments → Redeploy.

## Connect your domain

1. Vercel → Project → Settings → Domains
2. Add e.g. `app.statementiq.com`
3. Set CNAME at your registrar pointing to `cname.vercel-dns.com`

## Link landing page → app

In your landing page repo (`index.html`), update all CTA buttons to point to this app's URL:

```html
<a href="https://app.statementiq.com">Start free trial</a>
```

## Local development

```bash
# Install Python dependencies
pip install openpyxl

# Test the API locally
python -m http.server 3000

# Or use Vercel CLI
npm install -g vercel
vercel dev
```

## Currencies supported

USD · EUR · GBP · CAD · AUD · CHF · SGD · JPY

Each LP can have their own currency. The Excel statement uses the correct symbol and formatting automatically.
