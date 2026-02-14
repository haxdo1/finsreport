# How to Deploy to Vercel

Since your project is a static site (HTML, CSS, and JS), it is perfect for Vercel. Here are the two best ways to host it:

## Method 1: Via GitHub (Recommended)
1. Initialize a Git repository in `c:\xampp\htdocs\Prrm`.
2. Push your code to a new **GitHub repository**.
3. Go to [Vercel.com](https://vercel.com) and click **"New Project"**.
4. Import your GitHub repository.
5. Vercel will automatically detect the static files and deploy them.

## Method 2: Via Vercel CLI (Quickest)
1. Open up your terminal.
2. Navigate to your project: `cd c:\xampp\htdocs\Prrm`.
3. Install Vercel CLI: `npm i -g vercel`.
4. Run the command: `vercel`.
5. Follow the prompts to link your account and deploy.

### Pre-configured for you:
I have already added a `vercel.json` file in your project directory to ensure clean URLs and proper caching.

### ⚠️ IMPORTANT: Data Files
For the **Combined Data** and **Search** features to work instantly on the live site, you **MUST** upload your Excel files (`Proposal.xlsx` and `Reimbursement.xlsx`) to the repository.
1. Ensure these files are in the same folder as `index.html`.
2. Commit and push them to GitHub.
3. Vercel will serve them as static files which the App can fetch.
